# routes/dashboard.py
from flask import Blueprint, render_template, current_app
from flask_login import login_required, current_user
from sqlalchemy import func
from extensions import db
from models import Apprentice, TrainingGroup

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
            # Intenta obtener campos del modelo User/Apprentice relacionados al usuario actual.
            # Ajusta los nombres (pending_logs, uploaded_logs, group_number, last_upload) a tu modelo.
            pending = getattr(current_user, "pending_logs", None)
            uploaded = getattr(current_user, "uploaded_logs", None)
            group_number = getattr(current_user, "group_number", None) or getattr(current_user, "group", None) or "—"
            last_upload = getattr(current_user, "last_upload", None)

            # Si current_user no tiene esos atributos, intenta buscar un registro Apprentice relacionado
            if pending is None and uploaded is None:
                try:
                    apprentice_record = db.session.query(Apprentice).filter(
                        getattr(Apprentice, "student_user_id", None) == getattr(current_user, "id", None)
                    ).first()
                    if apprentice_record:
                        pending = getattr(apprentice_record, "pending_logs", 0)
                        uploaded = getattr(apprentice_record, "uploaded_logs", 0)
                        group_number = getattr(apprentice_record, "group_number", group_number)
                        last_upload = getattr(apprentice_record, "last_upload", last_upload)
                except Exception:
                    current_app.logger.debug("No se pudo buscar Apprentice relacionado al usuario", exc_info=True)

            student_dashboard = {
                "pending": pending or 0,
                "uploaded": uploaded or 0,
                "group_number": group_number,
                "last_upload": last_upload,
            }
    except Exception:
        current_app.logger.exception("Error construyendo student_dashboard")
        student_dashboard = {"pending": 0, "uploaded": 0, "group_number": "—", "last_upload": None}

    # Renderizar plantilla con valores seguros
    return render_template("dashboard/index.html", stats=stats, student_dashboard=student_dashboard)
