"""
services/date_rules.py

Single source of truth para fechas derivadas de GIA.

Regla de diseño:
- Las fechas que llegan desde una fuente institucional (p.ej. Excel) son datos
  de hecho y tienen prioridad.
- El sistema solo calcula fechas derivadas cuando existe la fecha fuente
  suficiente.
- No se infieren duraciones de formación únicamente por nivel. La duración
  definitiva debe provenir del programa/diseño curricular.

Estado 0.4.10:
- Vigencia de grupo: fin de formación + 6 meses (regla institucional dada para GIA).
- Momento 1: ventana de 15 días desde el inicio de etapa productiva.
- Momento 2: ventana de 15 días centrada en la mitad del periodo EP.
- Momento 3: ventana de 15 días que termina en el fin de EP.
- Momento 4: NO se calcula hasta contar con una regla institucional explícita.
"""

from __future__ import annotations

from datetime import date, timedelta
from typing import Any

from .utils_base import add_months, parse_date_value, format_date_value

FOLLOWUP_WINDOW_DAYS = 15
GROUP_VALIDITY_MONTHS = 6


def _fmt(value: date | None) -> str:
    return value.strftime("%d/%m/%Y") if value else ""


def calculate_group_validity_from_training_end(training_end_date: Any) -> str:
    end = parse_date_value(training_end_date)
    return _fmt(add_months(end, GROUP_VALIDITY_MONTHS)) if end else ""


def calculate_followup_ranges_from_ep(practice_start_date: Any, practice_end_date: Any) -> dict[str, str]:
    start = parse_date_value(practice_start_date)
    end = parse_date_value(practice_end_date)
    empty = {
        f"followup_moment{i}_{suffix}": ""
        for i in range(1, 5)
        for suffix in ("start", "end")
    }
    if not start or not end or end <= start:
        return empty

    ranges = dict(empty)
    # Momento 1: primeros 15 días de EP.
    m1_end = min(start + timedelta(days=FOLLOWUP_WINDOW_DAYS - 1), end)
    ranges["followup_moment1_start"] = _fmt(start)
    ranges["followup_moment1_end"] = _fmt(m1_end)

    # Momento 2: ventana de 15 días centrada en el punto medio del periodo.
    total_days = (end - start).days
    midpoint = start + timedelta(days=total_days // 2)
    half_left = (FOLLOWUP_WINDOW_DAYS - 1) // 2
    m2_start = max(start, midpoint - timedelta(days=half_left))
    m2_end = min(end, m2_start + timedelta(days=FOLLOWUP_WINDOW_DAYS - 1))
    if (m2_end - m2_start).days + 1 < FOLLOWUP_WINDOW_DAYS:
        m2_start = max(start, m2_end - timedelta(days=FOLLOWUP_WINDOW_DAYS - 1))
    ranges["followup_moment2_start"] = _fmt(m2_start)
    ranges["followup_moment2_end"] = _fmt(m2_end)

    # Momento 3: últimos 15 días de EP.
    m3_start = max(start, end - timedelta(days=FOLLOWUP_WINDOW_DAYS - 1))
    ranges["followup_moment3_start"] = _fmt(m3_start)
    ranges["followup_moment3_end"] = _fmt(end)

    # Momento 4 queda deliberadamente vacío hasta contar con la regla oficial.
    return ranges


def audit_date_consistency(*, group_start=None, ep_start=None, training_end=None, group_validity=None, practice_start=None, practice_end=None) -> list[str]:
    warnings: list[str] = []
    gs = parse_date_value(group_start)
    ep = parse_date_value(ep_start)
    te = parse_date_value(training_end)
    gv = parse_date_value(group_validity)
    ps = parse_date_value(practice_start)
    pe = parse_date_value(practice_end)

    if gs and te and te < gs:
        warnings.append("La fecha fin de formación es anterior al inicio de formación.")
    if ep and te and ep > te:
        warnings.append("El inicio de etapa productiva es posterior al fin de formación.")
    if ps and pe and pe <= ps:
        warnings.append("La fecha fin de etapa productiva del aprendiz debe ser posterior al inicio.")
    if te and gv:
        expected = add_months(te, GROUP_VALIDITY_MONTHS)
        if expected and gv != expected:
            warnings.append("La vigencia del grupo no coincide con fin de formación + 6 meses.")
    if ep and ps and ep != ps:
        warnings.append("El inicio de EP del aprendiz difiere del inicio de EP del grupo.")
    return warnings


def build_derived_group_dates(record: dict[str, Any]) -> dict[str, Any]:
    result = dict(record)
    # Never overwrite an imported/explicit validity.
    if not result.get("group_validity") and result.get("training_end_date"):
        result["group_validity"] = calculate_group_validity_from_training_end(result["training_end_date"])
    return result


def build_derived_apprentice_dates(record: dict[str, Any]) -> dict[str, Any]:
    result = dict(record)
    ranges = calculate_followup_ranges_from_ep(
        result.get("practice_start_date"),
        result.get("practice_end_date"),
    )
    for key, value in ranges.items():
        if not result.get(key) and value:
            result[key] = value
    return result


__all__ = [
    "FOLLOWUP_WINDOW_DAYS",
    "GROUP_VALIDITY_MONTHS",
    "calculate_group_validity_from_training_end",
    "calculate_followup_ranges_from_ep",
    "audit_date_consistency",
    "build_derived_group_dates",
    "build_derived_apprentice_dates",
]
