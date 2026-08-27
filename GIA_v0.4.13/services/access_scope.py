"""Alcance de acceso por registro para GIA.

Fase C del Problema 6:
- centraliza el alcance por grupos, aprendices y evidencias;
- mantiene la asignación de instructor como texto mientras el modelo no tenga FK;
- separa alcance (qué registros puede ver) de permisos (qué acciones puede ejecutar).
"""

from __future__ import annotations

from flask import abort
from flask_login import current_user
from sqlalchemy import func

from extensions import db
from models import Apprentice, EvidenceSubmission, TrainingGroup
from services.permissions import (
    ROLE_APPRENTICE,
    ROLE_CENTER_STAFF,
    ROLE_CERTIFIER,
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_SUPPORT,
)

GLOBAL_GROUP_VIEW_ROLES = frozenset({
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_CENTER_STAFF,
    ROLE_CERTIFIER,
    ROLE_SUPPORT,
})

ASSIGNED_GROUP_ROLES = frozenset({
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
})


def current_role() -> str | None:
    return getattr(current_user, "role", None)


def is_apprentice() -> bool:
    return current_role() == ROLE_APPRENTICE


def is_global_group_view() -> bool:
    return current_role() in GLOBAL_GROUP_VIEW_ROLES


def instructor_identity() -> str:
    return (
        getattr(current_user, "full_name", None)
        or getattr(current_user, "login_identifier", None)
        or ""
    ).strip()


def group_belongs_to_current_instructor(group: TrainingGroup) -> bool:
    """Comprueba el alcance del instructor sin inventar una FK inexistente."""
    assigned_name = (getattr(group, "followup_instructor", None) or "").strip()
    current_name = instructor_identity()
    if not assigned_name or not current_name:
        return False
    return assigned_name.casefold() == current_name.casefold()


def visible_groups_query():
    """Consulta base de grupos visibles para el usuario autenticado."""
    if is_global_group_view():
        return TrainingGroup.query

    if current_role() in ASSIGNED_GROUP_ROLES:
        name = instructor_identity()
        if not name:
            return TrainingGroup.query.filter(False)
        return TrainingGroup.query.filter(
            func.lower(func.trim(TrainingGroup.followup_instructor)) == name.lower()
        )

    return TrainingGroup.query.filter(False)


def visible_group_ids() -> list[int]:
    try:
        return [
            group_id
            for (group_id,) in visible_groups_query()
            .with_entities(TrainingGroup.id)
            .all()
        ]
    except Exception:
        db.session.rollback()
        return []


def can_view_group(group: TrainingGroup) -> bool:
    if is_global_group_view():
        return True
    return current_role() in ASSIGNED_GROUP_ROLES and group_belongs_to_current_instructor(group)


def can_manage_group(group: TrainingGroup) -> bool:
    # Soporte y líder tienen administración global; el instructor normal
    # permanece estrictamente limitado a sus grupos asignados.
    if current_role() in {ROLE_SUPPORT, ROLE_LEAD_FOLLOWUP_INSTRUCTOR}:
        return True
    return current_role() == ROLE_FOLLOWUP_INSTRUCTOR and group_belongs_to_current_instructor(group)


def can_manage_all_groups() -> bool:
    return current_role() in {ROLE_SUPPORT, ROLE_LEAD_FOLLOWUP_INSTRUCTOR}


def visible_apprentices_query():
    """Consulta base de aprendices respetando el alcance de grupos."""
    if is_apprentice():
        return Apprentice.query.filter(
            Apprentice.student_user_id == current_user.id
        )

    if is_global_group_view():
        return Apprentice.query

    group_ids = visible_group_ids()
    if not group_ids:
        return Apprentice.query.filter(False)

    return Apprentice.query.filter(
        Apprentice.group_id.in_(group_ids)
    )


def can_view_apprentice(apprentice: Apprentice) -> bool:
    if is_apprentice():
        return apprentice.student_user_id == current_user.id
    if is_global_group_view():
        return True
    group_id = getattr(apprentice, "group_id", None)
    return group_id in set(visible_group_ids()) if group_id else False


def visible_submissions_query():
    """Consulta base de evidencias visibles.

    Aprendiz: únicamente sus entregas.
    Roles institucionales globales: todas.
    Instructor/líder: únicamente evidencias de sus grupos.
    """
    if is_apprentice():
        apprentice = Apprentice.query.filter_by(student_user_id=current_user.id).first()
        if apprentice is None:
            return EvidenceSubmission.query.filter(False)
        return EvidenceSubmission.query.filter_by(apprentice_id=apprentice.id)

    if is_global_group_view():
        return EvidenceSubmission.query

    group_ids = visible_group_ids()
    if not group_ids:
        return EvidenceSubmission.query.filter(False)

    return EvidenceSubmission.query.filter(
        EvidenceSubmission.group_id.in_(group_ids)
    )


def can_view_submission(submission: EvidenceSubmission) -> bool:
    apprentice = getattr(submission, "apprentice", None)
    if apprentice is None:
        return False
    return can_view_apprentice(apprentice)


def require_submission_scope(submission: EvidenceSubmission) -> None:
    if not can_view_submission(submission):
        abort(403)
