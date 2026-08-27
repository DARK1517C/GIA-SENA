"""
Servicio de seguimiento de GIA.

Centraliza el estado operativo de los momentos de seguimiento usando:
- las fechas derivadas centralizadas en services.date_rules;
- las actividades institucionales FUP-01/FUP-02/FUP-03;
- el estado de la evidencia asociada a cada momento.

No introduce una segunda fuente de verdad ni una tabla paralela para el
seguimiento en esta fase. Los momentos se consideran atendidos cuando la
evidencia correspondiente llega a APROBADO.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from typing import Any

from sqlalchemy import and_

from extensions import db
from models import Apprentice, TrainingGroup, EvidenceSubmission
from models.evidence import (
    EVIDENCE_STATUS_NOT_SUBMITTED,
    EVIDENCE_STATUS_PENDING_REVIEW,
    EVIDENCE_STATUS_REQUIRES_CORRECTION,
    EVIDENCE_STATUS_APPROVED,
)


FOLLOWUP_STATUS_NOT_STARTED = "not_started"
FOLLOWUP_STATUS_PENDING = "pending"
FOLLOWUP_STATUS_IN_PROGRESS = "in_progress"
FOLLOWUP_STATUS_OVERDUE = "overdue"
FOLLOWUP_STATUS_PENDING_REVIEW = "pending_review"
FOLLOWUP_STATUS_REQUIRES_CORRECTION = "requires_correction"
FOLLOWUP_STATUS_COMPLETED = "completed"

FOLLOWUP_STATUS_LABELS = {
    FOLLOWUP_STATUS_NOT_STARTED: "No iniciado",
    FOLLOWUP_STATUS_PENDING: "Pendiente",
    FOLLOWUP_STATUS_IN_PROGRESS: "En curso",
    FOLLOWUP_STATUS_OVERDUE: "Vencido",
    FOLLOWUP_STATUS_PENDING_REVIEW: "En revisión",
    FOLLOWUP_STATUS_REQUIRES_CORRECTION: "Requiere corrección",
    FOLLOWUP_STATUS_COMPLETED: "Aprobado",
}

FOLLOWUP_MOMENTS = (
    ("Momento 1", "FUP-01", "followup_moment1_start", "followup_moment1_end"),
    ("Momento 2", "FUP-02", "followup_moment2_start", "followup_moment2_end"),
    ("Momento 3", "FUP-03", "followup_moment3_start", "followup_moment3_end"),
    # Momento 4 remains visible as optional/unconfigured, but is not used in
    # the operative status calculation until its institutional rule exists.
    ("Momento 4", "FUP-04", "followup_moment4_start", "followup_moment4_end"),
)

OPERATIVE_MOMENTS = FOLLOWUP_MOMENTS[:3]


def _today() -> date:
    return date.today()


def _normalize_date(value: Any) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
    return None


def calculate_followup_ranges(practice_start_date, practice_end_date):
    from services.date_rules import calculate_followup_ranges_from_ep
    return calculate_followup_ranges_from_ep(practice_start_date, practice_end_date)


def validate_followup_dates(practice_start_date, practice_end_date):
    warnings = []
    start = _normalize_date(practice_start_date)
    end = _normalize_date(practice_end_date)
    if start is None:
        warnings.append("La fecha de inicio de la etapa productiva no existe.")
    if end is None:
        warnings.append("La fecha final de la etapa productiva no existe.")
    if start and end and end <= start:
        warnings.append("La fecha final debe ser posterior a la fecha inicial.")
    return warnings


def _ensure_apprentice_submissions(apprentice):
    """Garantiza que existan las actividades institucionales FUP para el aprendiz."""
    if not apprentice:
        return
    from services.evidence_service import ensure_submissions_for_apprentice
    ensure_submissions_for_apprentice(apprentice)
    db.session.flush()


def _normalize_followup_code(code: str | None) -> str | None:
    """Normaliza códigos internos FUP_01 a códigos lógicos FUP-01."""
    if not code:
        return None

    value = str(code).strip().upper().replace("_", "-")

    if value in {"FUP-01", "FUP-02", "FUP-03", "FUP-04"}:
        return value

    return None


def _moment_submission_map(apprentice) -> dict[str, EvidenceSubmission]:
    """Devuelve la última entrega por código lógico FUP-01..FUP-03."""
    try:
        query = (
            EvidenceSubmission.query
            .join(EvidenceSubmission.activity)
            .filter(
                EvidenceSubmission.apprentice_id == apprentice.id,
                EvidenceSubmission.is_latest.is_(True),
            )
        )
        submissions = query.all()
    except Exception:
        # Evitamos romper el detalle si una base antigua no puede resolver la
        # relación; la UI mostrará el momento como no atendido.
        submissions = []

    result: dict[str, EvidenceSubmission] = {}
    operative_codes = {item[1] for item in OPERATIVE_MOMENTS}

    for submission in submissions:
        raw_code = getattr(submission.activity, "code", None)
        code = _normalize_followup_code(raw_code)
        if code in operative_codes:
            result[code] = submission

    return result


def _date_status(start: date | None, end: date | None, today: date) -> str:
    if not start or not end:
        return FOLLOWUP_STATUS_NOT_STARTED
    if today < start:
        return FOLLOWUP_STATUS_PENDING
    if start <= today <= end:
        return FOLLOWUP_STATUS_IN_PROGRESS
    return FOLLOWUP_STATUS_OVERDUE


def _moment_status(start, end, submission, today):
    if submission is not None:
        if submission.status == EVIDENCE_STATUS_APPROVED:
            return FOLLOWUP_STATUS_COMPLETED
        if submission.status == EVIDENCE_STATUS_REQUIRES_CORRECTION:
            return FOLLOWUP_STATUS_REQUIRES_CORRECTION
        if submission.status == EVIDENCE_STATUS_PENDING_REVIEW:
            return FOLLOWUP_STATUS_PENDING_REVIEW
        if submission.status == EVIDENCE_STATUS_NOT_SUBMITTED:
            return _date_status(start, end, today)
    return _date_status(start, end, today)


def get_apprentice_followup(apprentice, *, ensure_submissions=False) -> list[dict[str, Any]]:
    """Construye el estado de M1-M3 y deja M4 explícitamente no configurado."""
    if apprentice is None:
        return []

    if ensure_submissions:
        _ensure_apprentice_submissions(apprentice)

    today = _today()
    submission_map = _moment_submission_map(apprentice)
    rows: list[dict[str, Any]] = []

    for label, code, start_field, end_field in FOLLOWUP_MOMENTS:
        start = _normalize_date(getattr(apprentice, start_field, None))
        end = _normalize_date(getattr(apprentice, end_field, None))
        submission = submission_map.get(code)

        if label == "Momento 4" and not start and not end:
            status = FOLLOWUP_STATUS_NOT_STARTED
            configurable = False
        else:
            status = _moment_status(start, end, submission, today)
            configurable = True

        rows.append({
            "label": label,
            "code": code,
            "start": start,
            "end": end,
            "status": status,
            "status_label": FOLLOWUP_STATUS_LABELS[status],
            "submission": submission,
            "submission_id": getattr(submission, "id", None),
            "activity_title": getattr(getattr(submission, "activity", None), "title", None),
            "configurable": configurable,
            "is_operational": label != "Momento 4",
        })

    return rows


def get_next_followup(apprentice):
    rows = get_apprentice_followup(apprentice)
    for row in rows:
        if not row["is_operational"] or not row["configurable"]:
            continue
        if row["status"] != FOLLOWUP_STATUS_COMPLETED:
            return row
    return None


def get_followup_status(apprentice):
    rows = get_apprentice_followup(apprentice)
    if not rows:
        return FOLLOWUP_STATUS_NOT_STARTED

    operative = [row for row in rows if row["is_operational"]]
    if not operative:
        return FOLLOWUP_STATUS_NOT_STARTED

    statuses = [row["status"] for row in operative]
    if all(status == FOLLOWUP_STATUS_COMPLETED for status in statuses):
        return FOLLOWUP_STATUS_COMPLETED
    if FOLLOWUP_STATUS_REQUIRES_CORRECTION in statuses:
        return FOLLOWUP_STATUS_REQUIRES_CORRECTION
    if FOLLOWUP_STATUS_PENDING_REVIEW in statuses:
        return FOLLOWUP_STATUS_PENDING_REVIEW
    if FOLLOWUP_STATUS_IN_PROGRESS in statuses:
        return FOLLOWUP_STATUS_IN_PROGRESS
    if FOLLOWUP_STATUS_OVERDUE in statuses:
        return FOLLOWUP_STATUS_OVERDUE
    if FOLLOWUP_STATUS_PENDING in statuses:
        return FOLLOWUP_STATUS_PENDING
    return FOLLOWUP_STATUS_NOT_STARTED


def get_upcoming_followups(apprentices, days=7):
    today = _today()
    limit = today + timedelta(days=days)
    upcoming = []
    for apprentice in apprentices:
        next_followup = get_next_followup(apprentice)
        if not next_followup or not next_followup.get("start"):
            continue
        start = next_followup["start"]
        if today <= start <= limit:
            upcoming.append({
                "apprentice": apprentice,
                "moment": next_followup["label"],
                "start": start,
                "end": next_followup["end"],
                "days_remaining": (start - today).days,
                "status": next_followup["status"],
            })
    upcoming.sort(key=lambda item: item["start"])
    return upcoming


def get_overdue_followups(apprentices):
    today = _today()
    overdue = []
    for apprentice in apprentices:
        rows = get_apprentice_followup(apprentice)
        for row in rows:
            if not row["is_operational"] or row["status"] != FOLLOWUP_STATUS_OVERDUE:
                continue
            if not row["end"]:
                continue
            overdue.append({
                "apprentice": apprentice,
                "moment": row["label"],
                "start": row["start"],
                "end": row["end"],
                "days_overdue": (today - row["end"]).days,
                "status": row["status"],
            })
    overdue.sort(key=lambda item: item["days_overdue"], reverse=True)
    return overdue


def build_followup_alerts(apprentices, upcoming_days=7):
    alerts = []
    for item in get_upcoming_followups(apprentices, days=upcoming_days):
        apprentice = item["apprentice"]
        alerts.append({
            "type": "info",
            "moment": item["moment"],
            "apprentice": apprentice,
            "message": (
                f"{apprentice.first_names} {apprentice.last_names} "
                f"tiene {item['moment']} a partir del "
                f"{item['start'].strftime('%d/%m/%Y')}."
            ),
            "start": item["start"],
            "end": item["end"],
        })
    for item in get_overdue_followups(apprentices):
        apprentice = item["apprentice"]
        alerts.append({
            "type": "warning",
            "moment": item["moment"],
            "apprentice": apprentice,
            "message": (
                f"{apprentice.first_names} {apprentice.last_names} "
                f"tiene vencido {item['moment']}."
            ),
            "start": item["start"],
            "end": item["end"],
        })
    return alerts


def get_group_followups(group):
    if group is None:
        return []
    apprentices = Apprentice.query.filter_by(group_number=group.group_number).all()
    rows = []
    for apprentice in apprentices:
        details = get_apprentice_followup(apprentice)
        rows.append({
            "apprentice": apprentice,
            "status": get_followup_status(apprentice),
            "status_label": FOLLOWUP_STATUS_LABELS[get_followup_status(apprentice)],
            "moments": details,
            "next_followup": next((m for m in details if m["is_operational"] and m["status"] != FOLLOWUP_STATUS_COMPLETED), None),
        })
    return rows


def get_followup_dashboard(target):
    """Resumen para detalle de aprendiz o grupo."""
    if isinstance(target, Apprentice):
        moments = get_apprentice_followup(target)
        status = get_followup_status(target)
        return {
            "total_apprentices": 1,
            "status": status,
            "status_label": FOLLOWUP_STATUS_LABELS[status],
            "moments": moments,
            "next_followup": next((m for m in moments if m["is_operational"] and m["status"] != FOLLOWUP_STATUS_COMPLETED), None),
            "upcoming": [],
            "overdue": [],
        }

    rows = get_group_followups(target)
    statuses = [row["status"] for row in rows]
    return {
        "total_apprentices": len(rows),
        "not_started": sum(s == FOLLOWUP_STATUS_NOT_STARTED for s in statuses),
        "pending": sum(s == FOLLOWUP_STATUS_PENDING for s in statuses),
        "in_progress": sum(s == FOLLOWUP_STATUS_IN_PROGRESS for s in statuses),
        "pending_review": sum(s == FOLLOWUP_STATUS_PENDING_REVIEW for s in statuses),
        "requires_correction": sum(s == FOLLOWUP_STATUS_REQUIRES_CORRECTION for s in statuses),
        "overdue": sum(s == FOLLOWUP_STATUS_OVERDUE for s in statuses),
        "completed": sum(s == FOLLOWUP_STATUS_COMPLETED for s in statuses),
        "rows": rows,
    }


def get_group_followup_summary(group):
    dashboard = get_followup_dashboard(group)
    rows = dashboard.get("rows", [])
    next_items = [row["next_followup"] for row in rows if row.get("next_followup")]
    next_items.sort(key=lambda item: item.get("start") or date.max)
    return {
        "total": dashboard.get("total_apprentices", 0),
        "upcoming": len([item for item in next_items if item.get("start") and item["start"] >= _today()]),
        "overdue": dashboard.get("overdue", 0),
        "completed": dashboard.get("completed", 0),
        "next_followup": next_items[0] if next_items else None,
        "rows": rows,
        "pending_review": dashboard.get("pending_review", 0),
        "requires_correction": dashboard.get("requires_correction", 0),
        "in_progress": dashboard.get("in_progress", 0),
        "pending": dashboard.get("pending", 0),
    }


__all__ = [
    "FOLLOWUP_STATUS_NOT_STARTED",
    "FOLLOWUP_STATUS_PENDING",
    "FOLLOWUP_STATUS_IN_PROGRESS",
    "FOLLOWUP_STATUS_OVERDUE",
    "FOLLOWUP_STATUS_PENDING_REVIEW",
    "FOLLOWUP_STATUS_REQUIRES_CORRECTION",
    "FOLLOWUP_STATUS_COMPLETED",
    "FOLLOWUP_STATUS_LABELS",
    "FOLLOWUP_MOMENTS",
    "calculate_followup_ranges",
    "validate_followup_dates",
    "get_apprentice_followup",
    "get_next_followup",
    "get_followup_status",
    "get_upcoming_followups",
    "get_overdue_followups",
    "build_followup_alerts",
    "get_group_followups",
    "get_followup_dashboard",
    "get_group_followup_summary",
]
