"""
services/instructor_assignment.py

Fase A — asignación de instructores basada en el texto existente.

TrainingGroup todavía almacena los instructores como texto. Este servicio
centraliza la comparación de esos textos con la identidad del usuario
autenticado sin introducir todavía una relación User -> TrainingGroup.

La normalización de instructores reales queda para una fase posterior.
"""

from __future__ import annotations

import unicodedata


def normalize_instructor_name(value) -> str:
    """Normaliza un nombre para comparación segura, sin alterar el dato guardado."""
    if value is None:
        return ""

    text = " ".join(str(value).strip().split())
    if not text:
        return ""

    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        char for char in text
        if not unicodedata.combining(char)
    )

    return text.casefold()


def instructor_names_match(left, right) -> bool:
    """Indica si dos representaciones textuales identifican al mismo instructor."""
    left_normalized = normalize_instructor_name(left)
    right_normalized = normalize_instructor_name(right)

    if not left_normalized or not right_normalized:
        return False

    return left_normalized == right_normalized


def get_followup_group_ids(groups, instructor_name) -> list[int]:
    """
    Devuelve IDs de grupos cuyo instructor de seguimiento corresponde al nombre.

    Se usa en Fase A porque la asignación continúa siendo texto. La función
    permite que rutas de grupos, aprendices y estadísticas compartan exactamente
    la misma regla de comparación.
    """
    if not normalize_instructor_name(instructor_name):
        return []

    return [
        group.id
        for group in groups
        if instructor_names_match(
            getattr(group, "followup_instructor", None),
            instructor_name,
        )
    ]
