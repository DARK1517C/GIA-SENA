from __future__ import annotations

from flask import url_for
from sqlalchemy import func

from extensions import db
from models import Apprentice, Notification, User


def create_notification(*, user_id: int, notification_type: str, title: str, message: str, url: str | None = None) -> Notification:
    notification = Notification(
        user_id=user_id,
        notification_type=notification_type,
        title=title,
        message=message,
        url=url,
    )
    db.session.add(notification)
    return notification


def notify_apprentice_correction(submission) -> None:
    apprentice = submission.apprentice
    if not apprentice or not apprentice.student_user_id:
        return
    create_notification(
        user_id=apprentice.student_user_id,
        notification_type="EVIDENCE_CORRECTION",
        title="Se solicitó corrección de una evidencia",
        message=(
            f"El instructor indicó correcciones para “{submission.activity.title}”. "
            "Revisa la observación y vuelve a entregar la evidencia."
        ),
        url=url_for("evidences.detail", submission_id=submission.id),
    )


def notify_apprentice_approval(submission) -> None:
    apprentice = submission.apprentice
    if not apprentice or not apprentice.student_user_id:
        return
    create_notification(
        user_id=apprentice.student_user_id,
        notification_type="EVIDENCE_APPROVED",
        title="Evidencia aprobada",
        message=f"La evidencia “{submission.activity.title}” fue aprobada.",
        url=url_for("evidences.detail", submission_id=submission.id),
    )


def notify_reviewer_resubmission(submission) -> None:
    apprentice = submission.apprentice
    if not apprentice:
        return
    email = apprentice.followup_instructor_email
    if not email:
        return
    reviewer = User.query.filter(func.lower(User.email) == email.lower()).first()
    if not reviewer:
        return
    create_notification(
        user_id=reviewer.id,
        notification_type="EVIDENCE_RESUBMITTED",
        title="Evidencia reentregada",
        message=f"El aprendiz reentregó “{submission.activity.title}” para revisión.",
        url=url_for("evidences.detail", submission_id=submission.id),
    )


def notify_evidence_comment(submission, *, author) -> None:
    """Notifica a la contraparte sobre un comentario visible en la evidencia."""
    if not author:
        return

    if getattr(author, "role", None) == "APPRENTICE":
        apprentice = submission.apprentice
        email = getattr(apprentice, "followup_instructor_email", None) if apprentice else None
        if not email:
            return
        recipient = User.query.filter(func.lower(User.email) == email.lower()).first()
        if not recipient:
            return
        create_notification(
            user_id=recipient.id,
            notification_type="EVIDENCE_COMMENT",
            title="Nuevo comentario en una evidencia",
            message=(
                f"El aprendiz agregó un comentario en “{submission.activity.title}”."
            ),
            url=url_for("evidences.detail", submission_id=submission.id),
        )
        return

    apprentice = submission.apprentice
    if not apprentice or not apprentice.student_user_id:
        return
    create_notification(
        user_id=apprentice.student_user_id,
        notification_type="EVIDENCE_COMMENT",
        title="Nuevo comentario en una evidencia",
        message=(
            f"El instructor agregó un comentario en “{submission.activity.title}”."
        ),
        url=url_for("evidences.detail", submission_id=submission.id),
    )


def notify_apprentice_certification_approved(apprentice: Apprentice) -> None:
    if not apprentice or not apprentice.student_user_id:
        return
    create_notification(
        user_id=apprentice.student_user_id,
        notification_type="CERTIFICATION_APPROVED",
        title="Certificación aprobada",
        message="Tu proceso de certificación fue aprobado correctamente.",
        url=url_for("apprentices.detail", id=apprentice.id),
    )


def notify_apprentice_certification_rejected(apprentice: Apprentice) -> None:
    if not apprentice or not apprentice.student_user_id:
        return
    create_notification(
        user_id=apprentice.student_user_id,
        notification_type="CERTIFICATION_REJECTED",
        title="Revisión de certificación no aprobada",
        message="Tu revisión de certificación no fue aprobada. Revisa el detalle y las observaciones registradas.",
        url=url_for("certification.detail", apprentice_id=apprentice.id),
    )
