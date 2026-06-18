# routes/dashboard.py
from flask import Blueprint, render_template, current_app
from flask_login import login_required, current_user
from sqlalchemy import func
from extensions import db
from models import Apprentice, TrainingGroup, EvidenceSubmission
from services.evidence_service import ensure_submissions_for_apprentice, summarize_submissions

dashboard_bp = Blueprint("dashboard", __name__, url_prefix="/")


@dashboard_bp.route("/")
@login_required
def index():
    """
    Vista principal del dashboard.
    Calcula estadísticas generales y, si el usuario es aprendiz, datos específicos.
    Devuelve siempre valores por defecto seguros para evitar errores en la plantilla.
    """
    try:
        total_apprentices = db.session.query(func.count(Apprentice.id)).scalar() or 0
    except Exception:
        current_app.logger.exception("Error contando aprendices")
        total_apprentices = 0

    try:
        total_groups = db.session.query(func.count(TrainingGroup.id)).scalar() or 0
    except Exception:
        current_app.logger.exception("Error contando grupos")
        total_groups = 0

    # Estadísticas adicionales con defensas por si los campos no existen
    try:
        # Ajusta los filtros a los nombres reales de tus columnas (ej. 'sofia_status', 'status', etc.)
        in_training = db.session.query(func.count(Apprentice.id)).filter(
            getattr(Apprentice, "sofia_status", None) == "En lectiva"
        ).scalar() or 0
    except Exception:
        current_app.logger.debug("Campo 'sofia_status' no disponible o error en consulta", exc_info=True)
        in_training = 0

    try:
        in_practice = db.session.query(func.count(Apprentice.id)).filter(
            getattr(Apprentice, "sofia_status", None) == "En practica"
        ).scalar() or 0
    except Exception:
        current_app.logger.debug("Campo 'sofia_status' no disponible o error en consulta", exc_info=True)
        in_practice = 0

    try:
        enabled = db.session.query(func.count(Apprentice.id)).filter(
            getattr(Apprentice, "enabled", None) == True
        ).scalar() or 0
    except Exception:
        current_app.logger.debug("Campo 'enabled' no disponible o error en consulta", exc_info=True)
        enabled = 0

    # Alternativas / certificados (ajusta nombres de columnas si son distintos)
    try:
        with_alternative = db.session.query(func.count(Apprentice.id)).filter(
            getattr(Apprentice, "has_alternative", None) == True
        ).scalar() or 0
    except Exception:
        current_app.logger.debug("Campo 'has_alternative' no disponible o error en consulta", exc_info=True)
        with_alternative = 0

    try:
        certified = db.session.query(func.count(Apprentice.id)).filter(
            getattr(Apprentice, "certified", None) == True
        ).scalar() or 0
    except Exception:
        current_app.logger.debug("Campo 'certified' no disponible o error en consulta", exc_info=True)
        certified = 0

    without_alternative = max(0, (total_apprentices - with_alternative))

    stats = {
        "total_groups": total_groups,
        "total_apprentices": total_apprentices,
        "in_training": in_training,
        "in_practice": in_practice,
        "enabled": enabled,
        "with_alternative": with_alternative,
        "without_alternative": without_alternative,
        "certified": certified,
    }

    # Datos específicos para usuarios con rol 'aprendiz'
    student_dashboard = None
    try:
        if getattr(current_user, "role", None) == "aprendiz":
            pending = 0
            uploaded = 0
            group_number = getattr(current_user, "group_number", None) or getattr(current_user, "group", None) or "—"
            last_upload = None
            evidence_summary = {}

            try:
                apprentice_record = db.session.query(Apprentice).filter(
                    Apprentice.student_user_id == getattr(current_user, "id", None)
                ).first()
                if apprentice_record:
                    ensure_submissions_for_apprentice(apprentice_record)
                    db.session.commit()
                    submissions = EvidenceSubmission.query.filter_by(apprentice_id=apprentice_record.id).all()
                    pending = sum(1 for item in submissions if item.status in ["no_entregado", "pendiente_aprobacion"])
                    uploaded = sum(1 for item in submissions if item.status != "no_entregado")
                    group_number = getattr(apprentice_record, "group_number", group_number)
                    uploads = [item.uploaded_at for item in submissions if item.uploaded_at]
                    last_upload = max(uploads) if uploads else None
                    evidence_summary = summarize_submissions(submissions)
            except Exception:
                current_app.logger.debug("No se pudo buscar Apprentice relacionado al usuario", exc_info=True)

            student_dashboard = {
                "pending": pending or 0,
                "uploaded": uploaded or 0,
                "group_number": group_number,
                "last_upload": last_upload,
                "evidence_summary": evidence_summary,
            }
    except Exception:
        current_app.logger.exception("Error construyendo student_dashboard")
        student_dashboard = {"pending": 0, "uploaded": 0, "group_number": "—", "last_upload": None, "evidence_summary": {}}

    # Renderizar plantilla con valores seguros
    return render_template("dashboard/index.html", stats=stats, student_dashboard=student_dashboard)
