# routes/dashboard.py

from flask import Blueprint, render_template, current_app, redirect, url_for
from flask_login import login_required, current_user
from sqlalchemy import func

from extensions import db
from models import Apprentice, TrainingGroup, EvidenceSubmission

from catalogs.apprentice import EpModality, SofiaStatus
from catalogs.common_catalogs import ProgramLevel
from catalogs.display import get_label

from services.access_scope import visible_group_ids
from services.followup_service import build_followup_alerts, get_followup_status
from services.permissions import (
    ROLE_FOLLOWUP_INSTRUCTOR as ROLE_INSTRUCTOR_SEGUIMIENTO,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR as ROLE_INSTRUCTOR_SEGUIMIENTO_LIDER,
    ROLE_CENTER_STAFF as ROLE_ADMINISTRATIVO,
    ROLE_CERTIFIER as ROLE_CERTIFICADOR,
    ROLE_SUPPORT as ROLE_SOPORTE,
    ROLE_APPRENTICE,
    GLOBAL_ROLES,
)


dashboard_bp = Blueprint("dashboard", __name__, url_prefix="/")


# =============================================================================
# ROLES
# =============================================================================

# =============================================================================
# HELPERS GENERALES
# =============================================================================

def _current_role():
    """Devuelve el rol canónico del usuario autenticado."""
    return getattr(current_user, "role", None)


def _is_followup_instructor():
    return _current_role() == ROLE_INSTRUCTOR_SEGUIMIENTO


def _is_followup_leader():
    return _current_role() == ROLE_INSTRUCTOR_SEGUIMIENTO_LIDER


def _is_global_role():
    """
    Roles con visión administrativa/global.

    El instructor de seguimiento líder conserva su carácter administrativo
    para las estadísticas, aunque también pueda tener grupos específicos
    asignados.
    """
    return _current_role() in GLOBAL_ROLES


def _safe_count(query):
    """Ejecuta un COUNT y devuelve siempre un entero seguro."""
    try:
        return query.scalar() or 0
    except Exception:
        current_app.logger.exception("Error ejecutando estadística COUNT")
        return 0


def _safe_label(catalog, value):
    """
    Obtiene la etiqueta oficial de presentación desde catalogs/display.py.

    Nunca crea etiquetas paralelas dentro del dashboard.
    """
    try:
        if value is None or value == "":
            return "Sin especificar"

        return get_label(catalog, value)
    except Exception:
        current_app.logger.debug(
            "No se pudo obtener etiqueta para %s=%r",
            getattr(catalog, "__name__", str(catalog)),
            value,
            exc_info=True,
        )

        # Fallback solamente para evitar romper el dashboard.
        return str(value)


def _catalog_stat_item(catalog, value, count):
    """
    Estructura uniforme para las estadísticas de catálogos.

    Resultado:
        {
            "value": valor_canónico,
            "label": etiqueta_presentación,
            "count": cantidad
        }
    """
    return {
        "value": value,
        "label": _safe_label(catalog, value),
        "count": int(count or 0),
    }


# =============================================================================
# DETERMINACIÓN DEL ALCANCE
# =============================================================================

def _get_assigned_group_ids():
    """
    Obtiene los grupos asignados al instructor de seguimiento.

    La consulta intenta utilizar las relaciones/campos existentes en
    TrainingGroup sin asumir un único nombre de columna.

    Si el modelo todavía no expone la asignación, devuelve [] para evitar
    convertir accidentalmente la estadística en una estadística global.
    """
    # El modelo actual mantiene la asignación del instructor como texto.
    # Reutilizamos la misma fuente de alcance que protege Grupos/Aprendices/Evidencias
    # para que el dashboard no se quede vacío cuando el instructor sí tiene grupos.
    try:
        return list(visible_group_ids())
    except Exception:
        current_app.logger.exception(
            "No fue posible obtener los grupos visibles del instructor."
        )
        return []


def _get_scope():
    """
    Determina el alcance de las estadísticas.

    Retorna:
        {
            "type": "global" | "assigned",
            "label": str,
            "group_ids": [...]
        }
    """

    role = _current_role()

    # Instructor de seguimiento:
    # únicamente sus grupos asignados.
    if role == ROLE_INSTRUCTOR_SEGUIMIENTO:
        group_ids = _get_assigned_group_ids()

        return {
            "type": "assigned",
            "label": "Mis grupos y aprendices asignados",
            "group_ids": group_ids,
        }

    # Instructor de seguimiento líder:
    # visión administrativa.
    #
    # Puede tener grupos específicos a su cargo, pero su rol de líder
    # conserva la visión administrativa establecida para el dashboard.
    if role == ROLE_INSTRUCTOR_SEGUIMIENTO_LIDER:
        return {
            "type": "global",
            "label": "Vista administrativa",
            "group_ids": None,
        }

    # Roles administrativos/globales.
    return {
        "type": "global",
        "label": "Estadísticas generales",
        "group_ids": None,
    }


# =============================================================================
# CONSULTAS DE APRENDICES
# =============================================================================

def _base_apprentice_query(scope):
    """
    Construye la consulta base de aprendices según el alcance.
    """
    query = db.session.query(Apprentice)

    if scope["type"] == "assigned":
        group_ids = scope["group_ids"]

        # Si el instructor no tiene grupos asignados, el resultado debe ser
        # vacío, nunca global.
        if not group_ids:
            return query.filter(Apprentice.id == -1)

        group_field = getattr(Apprentice, "training_group_id", None)

        if group_field is not None:
            return query.filter(group_field.in_(group_ids))

        group_number_field = getattr(Apprentice, "group_number", None)
        group_number_model_field = getattr(TrainingGroup, "group_number", None)

        if group_number_field is not None and group_number_model_field is not None:
            return (
                query.join(
                    TrainingGroup,
                    group_number_field == group_number_model_field,
                )
                .filter(TrainingGroup.id.in_(group_ids))
            )

        # No existe una relación que permita garantizar el alcance.
        return query.filter(Apprentice.id == -1)

    return query


# =============================================================================
# CONSULTAS DE GRUPOS
# =============================================================================

def _base_group_query(scope):
    """
    Construye la consulta base de grupos según el alcance.
    """
    query = db.session.query(TrainingGroup)

    if scope["type"] == "assigned":
        group_ids = scope["group_ids"]

        if not group_ids:
            return query.filter(TrainingGroup.id == -1)

        return query.filter(TrainingGroup.id.in_(group_ids))

    return query


# =============================================================================
# ESTADÍSTICAS DE APRENDICES
# =============================================================================

def _apprentice_statistics(scope):
    """
    Estadísticas generales de aprendices dentro del alcance correspondiente.
    """

    query = _base_apprentice_query(scope)

    total_apprentices = _safe_count(
        query.with_entities(func.count(Apprentice.id))
    )

    # -------------------------------------------------------------------------
    # Modalidad de etapa productiva
    # -------------------------------------------------------------------------

    by_ep_modality = []

    ep_field = getattr(Apprentice, "ep_modality", None)

    if ep_field is not None:
        try:
            # Una modalidad de etapa productiva describe la situación
            # operativa de un aprendiz no certificado. Una vez certificado,
            # deja de pertenecer a las estadísticas derivadas de EP.
            modality_query = query
            if sofia_field is not None:
                modality_query = modality_query.filter(
                    sofia_field != SofiaStatus.CERTIFICADO.value
                )

            rows = (
                modality_query.with_entities(
                    ep_field,
                    func.count(Apprentice.id),
                )
                .filter(
                    ep_field.is_not(None),
                    func.trim(ep_field) != "",
                )
                .group_by(ep_field)
                .order_by(ep_field)
                .all()
            )

            for value, count in rows:
                by_ep_modality.append(
                    _catalog_stat_item(
                        EpModality,
                        value,
                        count,
                    )
                )

        except Exception:
            current_app.logger.exception(
                "Error calculando aprendices por modalidad de etapa productiva"
            )

    # -------------------------------------------------------------------------
    # Estado académico / habilitación / alternativa / certificación
    # -------------------------------------------------------------------------
    # En la versión actual no se utilizan "En lectiva" y "En productiva"
    # como indicadores del dashboard. La situación operativa se deriva de:
    # - certificado: sofia_status == CERTIFICADO
    # - con alternativa / habilitado: tiene modalidad de etapa productiva
    #   y aún no está certificado
    # - sin alternativa: no tiene modalidad y aún no está certificado
    # Esto mantiene las categorías mutuamente excluyentes.

    certified = 0
    with_alternative = 0

    sofia_field = getattr(Apprentice, "sofia_status", None)
    ep_field = getattr(Apprentice, "ep_modality", None)

    if sofia_field is not None:
        certified = _safe_count(
            query.filter(
                sofia_field == SofiaStatus.CERTIFICADO.value
            ).with_entities(func.count(Apprentice.id))
        )

    if ep_field is not None:
        with_alternative = _safe_count(
            query.filter(
                ep_field.is_not(None),
                func.trim(ep_field) != "",
                sofia_field != SofiaStatus.CERTIFICADO.value
                if sofia_field is not None else True,
            ).with_entities(func.count(Apprentice.id))
        )

    without_alternative = max(
        0,
        total_apprentices - certified - with_alternative,
    )

    # En el modelo vigente, la presencia de una modalidad de etapa productiva
    # es la evidencia persistida de que el aprendiz ya tiene alternativa.
    enabled = with_alternative

    return {
        "total_apprentices": total_apprentices,
        "enabled": enabled,
        "with_alternative": with_alternative,
        "without_alternative": without_alternative,
        "certified": certified,
        "by_ep_modality": by_ep_modality,
    }


# =============================================================================
# ESTADÍSTICAS DE GRUPOS
# =============================================================================

def _group_statistics(scope):
    """
    Estadísticas de grupos dentro del alcance correspondiente.
    """

    query = _base_group_query(scope)

    total_groups = _safe_count(
        query.with_entities(func.count(TrainingGroup.id))
    )

    # -------------------------------------------------------------------------
    # Grupos por nivel de formación
    # -------------------------------------------------------------------------

    by_program_level = []

    level_field = getattr(TrainingGroup, "program_level", None)

    if level_field is not None:
        try:
            rows = (
                query.with_entities(
                    level_field,
                    func.count(TrainingGroup.id),
                )
                .group_by(level_field)
                .order_by(level_field)
                .all()
            )

            for value, count in rows:
                by_program_level.append(
                    _catalog_stat_item(
                        ProgramLevel,
                        value,
                        count,
                    )
                )

        except Exception:
            current_app.logger.exception(
                "Error calculando grupos por nivel de formación"
            )

    return {
        "total_groups": total_groups,
        "by_program_level": by_program_level,
    }


# =============================================================================
# RUTA PRINCIPAL
# =============================================================================

@dashboard_bp.route("/")
@login_required
def index():
    """
    Dashboard principal de GIA.

    Alcances:

    - FOLLOW_UP_INSTRUCTOR:
        estadísticas únicamente de sus grupos y aprendices asignados.

    - FOLLOW_UP_INSTRUCTOR_lider:
        visión administrativa/global.

    - CENTER_STAFF:
        estadísticas globales.

    - CERTIFIER:
        estadísticas globales.

    - SUPPORT:
        estadísticas globales.

    - SUPPORT:
        estadísticas globales.

    Las estadísticas de catálogos se entregan con:
        value
        label
        count

    utilizando exclusivamente catalogs.display.get_label().
    """

    try:
        role = _current_role()

        # El aprendiz no utiliza el dashboard administrativo. Su superficie
        # operativa canónica es Evidencias.
        if role == ROLE_APPRENTICE:
            return redirect(url_for("evidences.index"))

        # ---------------------------------------------------------------------
        # Dashboard administrativo
        # ---------------------------------------------------------------------

        scope = _get_scope()

        apprentice_stats = _apprentice_statistics(scope)
        group_stats = _group_statistics(scope)

        stats = {
            "total_groups": group_stats["total_groups"],
            "total_apprentices": apprentice_stats["total_apprentices"],
            "enabled": apprentice_stats["enabled"],
            "with_alternative": apprentice_stats["with_alternative"],
            "without_alternative": apprentice_stats["without_alternative"],
            "certified": apprentice_stats["certified"],
            "by_ep_modality": apprentice_stats["by_ep_modality"],
            "by_program_level": group_stats["by_program_level"],
        }

        visible_apprentices = _base_apprentice_query(scope).all()
        followup_alerts = build_followup_alerts(visible_apprentices, upcoming_days=7)
        followup_statuses = [get_followup_status(a) for a in visible_apprentices]
        followup_summary = {
            "not_started": sum(s == "not_started" for s in followup_statuses),
            "pending": sum(s == "pending" for s in followup_statuses),
            "in_progress": sum(s == "in_progress" for s in followup_statuses),
            "pending_review": sum(s == "pending_review" for s in followup_statuses),
            "requires_correction": sum(s == "requires_correction" for s in followup_statuses),
            "overdue": sum(s == "overdue" for s in followup_statuses),
            "completed": sum(s == "completed" for s in followup_statuses),
        }

        return render_template(
            "dashboard/index.html",
            stats=stats,
            scope=scope,
            by_ep_modality=apprentice_stats["by_ep_modality"],
            by_program_level=group_stats["by_program_level"],
            followup_summary=followup_summary,
            followup_alerts=followup_alerts,
        )

    except Exception:
        current_app.logger.exception(
            "Error construyendo dashboard"
        )

        # Valores seguros para no romper la interfaz.
        safe_stats = {
            "total_groups": 0,
            "total_apprentices": 0,
            "enabled": 0,
            "with_alternative": 0,
            "without_alternative": 0,
            "certified": 0,
            "by_ep_modality": [],
            "by_program_level": [],
        }

        return render_template(
            "dashboard/index.html",
            stats=safe_stats,
            scope={
                "type": "global",
                "label": "Estadísticas generales",
            },
            by_ep_modality=[],
            by_program_level=[],
            followup_summary={},
            followup_alerts=[],
        )