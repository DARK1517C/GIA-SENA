from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any

from sqlalchemy import and_

from extensions import db
from models import Apprentice, CertificationReview, EvidenceSubmission
from models.certification import CERTIFICATION_REVIEW_APPROVED, CERTIFICATION_REVIEW_REJECTED
from models.evidence import EVIDENCE_STATUS_APPROVED
from services.evidence_service import ensure_submissions_for_apprentice
from services.followup_service import get_apprentice_followup

CERTIFICATION_CATEGORY_CODE = "certification_requirements"


@dataclass(frozen=True)
class CertificationCheck:
    label: str
    code: str
    status: str
    ok: bool
    submission_id: int | None
    title: str | None



def _latest_by_activity(apprentice_id: int):
    rows = (
        EvidenceSubmission.query
        .join(EvidenceSubmission.activity)
        .join(EvidenceSubmission.activity.property.mapper.class_.category)
        .filter(
            EvidenceSubmission.apprentice_id == apprentice_id,
            EvidenceSubmission.is_latest.is_(True),
        )
        .all()
    )
    return rows


def build_certification_checklist(apprentice: Apprentice) -> dict[str, Any]:
    ensure_submissions_for_apprentice(apprentice)
    db.session.flush()

    submissions = _latest_by_activity(apprentice.id)
    certification_rows = []
    followup_rows = get_apprentice_followup(apprentice, ensure_submissions=False)

    for submission in submissions:
        category = getattr(submission.activity, "category", None)
        if getattr(category, "code", None) != CERTIFICATION_CATEGORY_CODE:
            continue
        certification_rows.append(
            CertificationCheck(
                label=getattr(category, "name", "Requisitos de certificación"),
                code=getattr(submission.activity, "code", ""),
                status=submission.status,
                ok=submission.status == EVIDENCE_STATUS_APPROVED,
                submission_id=submission.id,
                title=getattr(submission.activity, "title", None),
            )
        )

    # La certificación exige que los momentos operativos de seguimiento estén aprobados.
    followup_operational = [r for r in followup_rows if r.get("is_operational")]
    followup_ok = bool(followup_operational) and all(
        row.get("status") == "completed" for row in followup_operational
    )

    certification_rows.sort(key=lambda x: x.code)
    all_required = bool(certification_rows) and all(r.ok for r in certification_rows)
    ready = all_required and followup_ok and not apprentice.is_certified

    latest_review = (
        CertificationReview.query
        .filter_by(apprentice_id=apprentice.id)
        .order_by(CertificationReview.created_at.desc())
        .first()
    )

    return {
        "certification_requirements": certification_rows,
        "followup": followup_rows,
        "followup_ok": followup_ok,
        "requirements_ok": all_required,
        "ready": ready,
        "latest_review": latest_review,
        "is_certified": apprentice.is_certified,
    }


def approve_certification(apprentice: Apprentice, reviewer, notes: str | None = None) -> CertificationReview:
    checklist = build_certification_checklist(apprentice)
    if apprentice.is_certified:
        raise ValueError("El aprendiz ya figura como certificado.")
    if not checklist["ready"]:
        raise ValueError("El aprendiz aún no cumple todos los requisitos de certificación.")

    review = CertificationReview(
        apprentice_id=apprentice.id,
        reviewer_id=reviewer.id,
        status=CERTIFICATION_REVIEW_APPROVED,
        notes=(notes or "").strip() or None,
        reviewed_at=datetime.now(timezone.utc),
    )
    apprentice.sofia_status = "CERTIFICADO"
    apprentice.evaluation_date = datetime.now(timezone.utc).date().isoformat()
    db.session.add(review)
    db.session.flush()
    return review


def reject_certification(apprentice: Apprentice, reviewer, notes: str) -> CertificationReview:
    text = (notes or "").strip()
    if not text:
        raise ValueError("Debe registrar un motivo para no aprobar la certificación.")
    review = CertificationReview(
        apprentice_id=apprentice.id,
        reviewer_id=reviewer.id,
        status=CERTIFICATION_REVIEW_REJECTED,
        notes=text,
        reviewed_at=datetime.now(timezone.utc),
    )
    db.session.add(review)
    db.session.flush()
    return review
