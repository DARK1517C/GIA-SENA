# =============================================================================
# routes/groups.py
# =============================================================================

from datetime import datetime, timezone
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from io import BytesIO
from types import SimpleNamespace

from flask import (
    Blueprint,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from flask_login import current_user, login_required
from sqlalchemy import case, func, or_
from werkzeug.utils import secure_filename

from extensions import db

from models import Apprentice, TrainingGroup

from catalogs.common_catalogs import ProgramLevel
from catalogs.apprentice import EpModality
from catalogs.display import get_label
from catalogs.training_group import (
    GroupModality,
    GroupMunicipality,
    GroupStatus,
)

from services.auth_helpers import permission_required
from services.evidence_service import (
    ensure_group_submissions,
    ensure_template_activities_for_group,
)
from services.excel_export import (
    export_reference_workbook,
    format_value,
)
from services.excel_import import import_reference_workbook
from services.followup_service import get_group_followup_summary
from services.date_rules import build_derived_group_dates
from services.access_scope import (
    can_manage_all_groups as _scope_can_manage_all_groups,
    can_manage_group as _scope_can_manage_group,
    can_view_group as _scope_can_view_group,
    group_belongs_to_current_instructor as _scope_group_belongs_to_current_instructor,
    instructor_identity as _scope_group_identity,
    visible_groups_query as _scope_visible_groups_query,
)
from services.permissions import (
    ROLE_APPRENTICE,
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_CERTIFIER,
    ROLE_CENTER_STAFF,
    ROLE_SUPPORT,
)


# =============================================================================
# Blueprint
# =============================================================================

groups_bp = Blueprint(
    "groups",
    __name__,
    url_prefix="/groups",
)


# =============================================================================
# Roles canónicos
# =============================================================================

# =============================================================================
# Campos del formulario
# =============================================================================

GROUP_FORM_FIELDS = [
    "group_number",
    "program_name",
    "program_level",
    "modality",
    "municipality",
    "group_validity",
    "sofia_group_status",
    "group_start_date",
    "training_end_date",
    "ep_start_date",
    "lead_instructor",
    "followup_instructor",
    "apprentices_enabled",
    "apprentices_certified",
]


CARD_FIELDS = [
    "group_number",
    "program_name",
    "program_level",
    "municipality",
    "modality",
    "lead_instructor",
    "followup_instructor",
    "group_start_date",
    "training_end_date",
    "ep_start_date",
    "sofia_group_status",
]


# =============================================================================
# Helpers de roles
# =============================================================================

def _current_role():
    return getattr(current_user, "role", None)


def _is_followup_instructor():
    return _current_role() == ROLE_FOLLOWUP_INSTRUCTOR


def _is_lead_followup_instructor():
    return _current_role() == ROLE_LEAD_FOLLOWUP_INSTRUCTOR


def _is_center_staff():
    return _current_role() == ROLE_CENTER_STAFF


def _is_support():
    return _current_role() == ROLE_SUPPORT


def _is_certifier():
    return _current_role() == ROLE_CERTIFIER


def _is_apprentice():
    return _current_role() == ROLE_APPRENTICE


def _has_global_group_view():
    """
    Usuarios con visión global del módulo de grupos.

    El instructor de seguimiento y el líder NO entran aquí:
    ambos trabajan operativamente con sus grupos asignados.
    """
    return _current_role() in {
        ROLE_CENTER_STAFF,
        ROLE_SUPPORT,
        ROLE_CERTIFIER,
    }


def _has_group_management():
    """
    Roles que pueden crear/modificar/eliminar/importar grupos.

    El instructor líder conserva capacidades administrativas del módulo,
    pero su listado operativo continúa segmentado a sus grupos asignados.
    """
    return _current_role() in {
        ROLE_FOLLOWUP_INSTRUCTOR,
        ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
        ROLE_SUPPORT,
    }


def _can_manage_all_groups():
    return _scope_can_manage_all_groups()


def _can_manage_group(group):
    return _scope_can_manage_group(group)


def _can_view_group(group):
    return _scope_can_view_group(group)


def _group_belongs_to_followup_instructor(group):
    return _scope_group_belongs_to_current_instructor(group)


def _require_group_management():
    if _has_group_management():
        return True

    flash(
        "No tienes permisos para realizar esta acción.",
        "warning",
    )
    return False


def _require_global_management():
    if _can_manage_all_groups():
        return True

    flash(
        "Esta acción requiere permisos globales de soporte.",
        "warning",
    )
    return False


# =============================================================================
# Helpers de formulario
# =============================================================================

def _group_form_data():
    """
    Obtiene y limpia los datos enviados desde el formulario.
    """
    return {
        field: request.form.get(field, "").strip()
        for field in GROUP_FORM_FIELDS
    }


# =============================================================================
# Helpers de catálogos
# =============================================================================

def _catalog_choices(catalog):
    """
    Devuelve opciones de catálogo en formato:

        {
            "value": "...",
            "label": "..."
        }

    La plantilla no necesita conocer el Enum ni interpretar sus valores.
    """
    choices = []

    for item in catalog:
        choices.append(
            {
                "value": item.value,
                "label": get_label(catalog, item.value),
            }
        )

    return choices


def _group_catalogs():
    """
    Catálogos oficiales utilizados por el módulo de grupos.
    """
    return {
        "program_levels": _catalog_choices(ProgramLevel),
        "modalities": _catalog_choices(GroupModality),
        "sofia_statuses": _catalog_choices(GroupStatus),
        "municipalities": _catalog_choices(GroupMunicipality),
    }


# =============================================================================
# Helpers de presentación
# =============================================================================

def _safe_catalog_label(catalog, value):
    """
    Obtiene una etiqueta sin romper la página si existe un dato antiguo,
    vacío o no reconocido.

    Esto permite que datos históricos no normalizados sigan siendo visibles.
    """
    if value is None or value == "":
        return ""

    try:
        return get_label(catalog, value)
    except Exception:
        return str(value)


def _local_now_label():
    """Fecha/hora de actualización mostrada en el huso institucional."""
    tz_name = current_app.config.get("DISPLAY_TIMEZONE", "America/Bogota")
    try:
        display_tz = ZoneInfo(tz_name)
    except ZoneInfoNotFoundError:
        display_tz = timezone.utc
    return datetime.now(timezone.utc).astimezone(display_tz).strftime("%d/%m/%Y %H:%M:%S")


def _program_level_css(value):
    """
    Clase visual preparada por backend.

    La plantilla no interpreta si el valor corresponde a TECNICO,
    TECNOLOGO, etc.
    """
    mapping = {
        ProgramLevel.OPERARIO.value: "group-card--operario",
        ProgramLevel.AUXILIAR.value: "group-card--auxiliar",
        ProgramLevel.TECNICO.value: "group-card--tecnico",
        ProgramLevel.TECNOLOGO.value: "group-card--tecnologo",
    }

    return mapping.get(
        value,
        "group-card--default",
    )


def _build_card_row(group, consecutive):
    """
    Construye la información que necesita groups/index.html.

    Los campos de catálogo reciben:
        *_label
    mientras los valores originales permanecen disponibles cuando
    el HTML necesite el valor canónico para enlaces/formularios.
    """
    row = {}

    for field in CARD_FIELDS:
        value = getattr(group, field, None)

        row[field] = format_value(value)

    row["consecutive"] = consecutive

    row["program_level_label"] = _safe_catalog_label(
        ProgramLevel,
        group.program_level,
    )

    row["modality_label"] = _safe_catalog_label(
        GroupModality,
        group.modality,
    )

    row["sofia_group_status_label"] = _safe_catalog_label(
        GroupStatus,
        group.sofia_group_status,
    )

    row["municipality_label"] = _safe_catalog_label(
        GroupMunicipality,
        group.municipality,
    )

    row["program_level_class"] = _program_level_css(
        group.program_level,
    )

    return row


# =============================================================================
# Índice
# =============================================================================

@groups_bp.route("/")
@login_required
def index():
    """
    Listado principal de grupos.

    Alcance:
    - FOLLOW_UP_INSTRUCTOR:
        únicamente grupos donde figura como instructor de seguimiento.
    - FOLLOW_UP_INSTRUCTOR_lider:
        únicamente grupos donde figura como instructor de seguimiento.
        Su visión administrativa/global de estadísticas pertenece al dashboard,
        no se mezcla con el listado operativo de grupos.
    - CENTER_STAFF:
        visión global.
    - CERTIFIER:
        visión global de consulta.
    - SUPPORT:
        visión global.
    - APPRENTICE:
        no tiene acceso al módulo de grupos.
    """

    if _is_apprentice():
        flash(
            "No tienes permisos para acceder a los grupos.",
            "warning",
        )
        return redirect(url_for("dashboard.index"))

    groups = (
        _filtered_groups_query()
        .order_by(TrainingGroup.group_number)
        .all()
    )

    group_record_fields = getattr(
        TrainingGroup,
        "RECORD_FIELDS",
        [],
    )

    cards = []

    for index_number, group in enumerate(groups, start=1):
        cards.append(
            {
                "group": group,
                "row": _build_card_row(
                    group,
                    index_number,
                ),
                "followup": get_group_followup_summary(group),
            }
        )

    # El total debe corresponder al alcance del usuario.
    total_groups = len(groups)

    catalogs = _group_catalogs()

    return render_template(
        "groups/index.html",
        groups=groups,
        cards=cards,
        group_record_fields=group_record_fields,
        total_groups=total_groups,
        program_levels=catalogs["program_levels"],
        modalities=catalogs["modalities"],
        sofia_statuses=catalogs["sofia_statuses"],
        municipalities=catalogs["municipalities"],
        can_create=_has_group_management(),
        can_edit=_has_group_management(),
        can_delete=_has_group_management(),
        can_bulk_delete=_has_group_management(),
        can_delete_all=_can_manage_all_groups(),
        can_import=_has_group_management(),
        can_export=not _is_apprentice(),
        is_global_view=_has_global_group_view(),
        is_assigned_view=(
            _is_followup_instructor()
            or _is_lead_followup_instructor()
        ),
    )


# =============================================================================
# Consulta filtrada
# =============================================================================

def _filtered_groups_query():
    """
    Construye la consulta aplicando primero el alcance por rol
    y posteriormente los filtros solicitados por el usuario.

    Esto evita que un instructor pueda recuperar grupos ajenos
    mediante parámetros GET.
    """
    query = _scope_visible_groups_query()

    # -------------------------------------------------------------------------
    # Filtros
    # -------------------------------------------------------------------------

    search = (
        request.args.get("search", "")
        .strip()
    )

    municipality = (
        request.args.get("municipality", "")
        .strip()
    )

    program_level = (
        request.args.get("program_level", "")
        .strip()
    )

    sofia_group_status = (
        request.args.get("sofia_group_status", "")
        .strip()
    )

    if search:
        search = search[:300]

        pattern = f"%{search}%"

        query = query.filter(
            or_(
                TrainingGroup.group_number.ilike(pattern),
                TrainingGroup.program_name.ilike(pattern),
                TrainingGroup.lead_instructor.ilike(pattern),
                TrainingGroup.followup_instructor.ilike(pattern),
            )
        )

    if municipality:
        query = query.filter(
            TrainingGroup.municipality == municipality
        )

    if program_level:
        query = query.filter(
            TrainingGroup.program_level == program_level
        )

    if sofia_group_status:
        query = query.filter(
            TrainingGroup.sofia_group_status == sofia_group_status
        )

    return query


# =============================================================================
# Detalle
# =============================================================================

@groups_bp.route("/<int:id>")
@login_required
def detail(id):
    group = TrainingGroup.query.get_or_404(id)

    if not _can_view_group(group):
        flash(
            "No tienes permisos para consultar este grupo.",
            "warning",
        )
        return redirect(url_for("groups.index"))

    ensure_group_submissions(group)

    try:
        db.session.commit()
    except Exception:
        db.session.rollback()
        current_app.logger.exception(
            "Error preparando evidencias del grupo %s.",
            group.group_number,
        )

    record_fields = getattr(
        TrainingGroup,
        "RECORD_FIELDS",
        [],
    )

    row = {}

    for key, _label in record_fields:
        if key == "consecutive":
            continue

        row[key] = format_value(
            getattr(
                group,
                key,
                None,
            )
        )

    # Fechas derivadas para presentación; respetan fechas explícitas importadas.
    row = build_derived_group_dates(row)

    # Etiquetas oficiales para el detalle.
    row["program_level_label"] = _safe_catalog_label(
        ProgramLevel,
        group.program_level,
    )

    row["modality_label"] = _safe_catalog_label(
        GroupModality,
        group.modality,
    )

    row["sofia_group_status_label"] = _safe_catalog_label(
        GroupStatus,
        group.sofia_group_status,
    )

    row["municipality_label"] = _safe_catalog_label(
        GroupMunicipality,
        group.municipality,
    )

    apprentices = (
        Apprentice.query
        .filter_by(group_id=group.id)
        .order_by(
            Apprentice.last_names,
            Apprentice.first_names,
        )
        .all()
    )

    # Etiquetas de presentación normalizadas para la tabla del detalle del grupo.
    for apprentice in apprentices:
        apprentice.program_level_label = _safe_catalog_label(
            ProgramLevel, apprentice.program_level
        )
        apprentice.ep_modality_label = _safe_catalog_label(
            EpModality, apprentice.ep_modality
        )
        try:
            from catalogs.apprentice import SofiaStatus
            apprentice.sofia_status_label = _safe_catalog_label(
                SofiaStatus, apprentice.sofia_status
            )
        except Exception:
            apprentice.sofia_status_label = apprentice.sofia_status or "Sin estado"

    # -------------------------------------------------------------------------
    # Estadísticas
    # -------------------------------------------------------------------------

    try:
        stats = compute_group_stats(group)

        row.update(
            {
                "contrato_aprendizaje": int(stats.contrato_aprendizaje or 0),
                "contrato_vinculo_formativo": int(stats.contrato_vinculo_formativo or 0),
                "vinculo_laboral": int(stats.vinculo_laboral or 0),
                "proyecto_productivo": int(stats.proyecto_productivo or 0),
                "monitoria": int(stats.monitoria or 0),
                "practicas_economia_popular": int(stats.practicas_economia_popular or 0),
                "apprentices_statistics": int(
                    stats.total or 0
                ),
                "apprentices_enabled": int(
                    stats.habilitados or 0
                ),
                "learning_contract": int(
                    stats.con_alternativa or 0
                ),
                "apprentices_without_alternative": int(
                    stats.sin_alternativa or 0
                ),
                "apprentices_certified": int(
                    stats.certificados or 0
                ),
            }
        )

        # ---------------------------------------------------------------------
        # Modalidades de etapa productiva
        # ---------------------------------------------------------------------
        # El backend entrega directamente label + count.
        # El template no interpreta EpModality.
        # ---------------------------------------------------------------------

        ep_modality_stats = [
        {
            "label": get_label(
                EpModality,
                EpModality.CONTRATO_APRENDIZAJE.value,
            ),
            "count": int(
                stats.contrato_aprendizaje or 0
            ),
        },
        {
            "label": get_label(
                EpModality,
                EpModality.CONTRATO_VINCULO_FORMATIVO.value,
            ),
            "count": int(
                stats.contrato_vinculo_formativo or 0
            ),
        },
        {
            "label": get_label(
                EpModality,
                EpModality.VINCULO_LABORAL.value,
            ),
            "count": int(
                stats.vinculo_laboral or 0
            ),
        },
        {
            "label": get_label(
                EpModality,
                EpModality.PROYECTO_PRODUCTIVO.value,
            ),
            "count": int(
                stats.proyecto_productivo or 0
            ),
        },
        {
            "label": get_label(
                EpModality,
                EpModality.MONITORIA.value,
            ),
            "count": int(
                stats.monitoria or 0
            ),
        },
        {
            "label": get_label(
                EpModality,
                EpModality.PRACTICAS_ECONOMIA_POPULAR.value,
            ),
            "count": int(
                stats.practicas_economia_popular or 0
            ),
        },
    ]

        stats_real_time = True
        stats_last_updated = (
            _local_now_label()
        )

    except Exception:
        current_app.logger.exception(
            "Error calculando estadísticas del grupo %s.",
            group.group_number,
        )

        stats_real_time = False
        stats_last_updated = None

        ep_modality_stats = []

    followup = get_group_followup_summary(group)

    return render_template(
        "groups/detail.html",
        group=group,
        row=row,
        apprentices=apprentices,
        group_record_fields=record_fields,
        stats_real_time=stats_real_time,
        stats_last_updated=stats_last_updated,
        ep_modality_stats=ep_modality_stats,
        followup=followup,
        followup_rows=followup.get("rows", []),
        can_edit=_can_manage_group(group),
        can_delete=_can_manage_group(group),
    )


# =============================================================================
# Estadísticas del grupo
# =============================================================================

def compute_group_stats(record):
    """
    Calcula estadísticas dinámicas de un grupo.

    Las estadísticas se obtienen directamente desde Apprentice.
    Las modalidades de etapa productiva se comparan utilizando
    exclusivamente los valores canónicos de EpModality.

    No modifica la base de datos.
    """

    if getattr(record, "id", None):
        group_filter = Apprentice.group_id == record.id
    else:
        group_filter = (
            func.trim(Apprentice.group_number)
            == func.trim(record.group_number)
        )

    # -------------------------------------------------------------------------
    # Total
    # -------------------------------------------------------------------------

    total = int(
        db.session.query(
            func.count(Apprentice.id).label("total")
        )
        .filter(group_filter)
        .one()
        .total
        or 0
    )

    # -------------------------------------------------------------------------
    # Indicadores generales
    # -------------------------------------------------------------------------

    # Estadísticas derivadas calculadas a continuación.
    # -------------------------------------------------------------------------
    # Habilitación, alternativa y certificación
    # -------------------------------------------------------------------------

    certificados = 0
    if hasattr(Apprentice, "sofia_status"):
        from catalogs.apprentice import SofiaStatus
        certificados = int(
            db.session.query(func.count(Apprentice.id))
            .filter(
                group_filter,
                Apprentice.sofia_status == SofiaStatus.CERTIFICADO.value,
            )
            .scalar()
            or 0
        )

    con_alternativa = 0
    if hasattr(Apprentice, "ep_modality"):
        con_alternativa = int(
            db.session.query(func.count(Apprentice.id))
            .filter(
                group_filter,
                Apprentice.ep_modality.is_not(None),
                func.trim(Apprentice.ep_modality) != "",
                Apprentice.sofia_status != SofiaStatus.CERTIFICADO.value,
            )
            .scalar()
            or 0
        )

    habilitados = con_alternativa
    sin_alternativa = max(0, total - certificados - con_alternativa)

    # -------------------------------------------------------------------------
    # Modalidades oficiales de Etapa Productiva.
    # Los aprendices certificados no pertenecen a estas estadísticas
    # derivadas de la etapa productiva.
    # -------------------------------------------------------------------------
    #
    # IMPORTANTE:
    # Se utilizan los valores oficiales de EpModality.
    #
    # No se utilizan LIKE, textos de presentación ni nombres alternativos.
    # La presentación de las etiquetas corresponde a display.py.
    #
    # EpModality oficial:
    #
    # CONTRATO_APRENDIZAJE
    # CONTRATO_VINCULO_FORMATIVO
    # VINCULO_LABORAL
    # PROYECTO_PRODUCTIVO
    # MONITORIA
    # PRACTICAS_ECONOMIA_POPULAR
    #
    # -------------------------------------------------------------------------

    contrato_aprendizaje = 0
    contrato_vinculo_formativo = 0
    vinculo_laboral = 0
    proyecto_productivo = 0
    monitoria = 0
    practicas_economia_popular = 0

    if hasattr(Apprentice, "ep_modality"):

        modalidad = func.coalesce(
            Apprentice.ep_modality,
            "",
        )

        result = (
            db.session.query(
                func.coalesce(
                    func.sum(
                        case(
                            (
                                modalidad
                                == EpModality.CONTRATO_APRENDIZAJE.value,
                                1,
                            ),
                            else_=0,
                        )
                    ),
                    0,
                ).label("contrato_aprendizaje"),

                func.coalesce(
                    func.sum(
                        case(
                            (
                                modalidad
                                == EpModality.CONTRATO_VINCULO_FORMATIVO.value,
                                1,
                            ),
                            else_=0,
                        )
                    ),
                    0,
                ).label("contrato_vinculo_formativo"),

                func.coalesce(
                    func.sum(
                        case(
                            (
                                modalidad
                                == EpModality.VINCULO_LABORAL.value,
                                1,
                            ),
                            else_=0,
                        )
                    ),
                    0,
                ).label("vinculo_laboral"),

                func.coalesce(
                    func.sum(
                        case(
                            (
                                modalidad
                                == EpModality.PROYECTO_PRODUCTIVO.value,
                                1,
                            ),
                            else_=0,
                        )
                    ),
                    0,
                ).label("proyecto_productivo"),

                func.coalesce(
                    func.sum(
                        case(
                            (
                                modalidad
                                == EpModality.MONITORIA.value,
                                1,
                            ),
                            else_=0,
                        )
                    ),
                    0,
                ).label("monitoria"),

                func.coalesce(
                    func.sum(
                        case(
                            (
                                modalidad
                                == EpModality.PRACTICAS_ECONOMIA_POPULAR.value,
                                1,
                            ),
                            else_=0,
                        )
                    ),
                    0,
                ).label("practicas_economia_popular"),
            )
            .filter(
                group_filter,
                Apprentice.sofia_status != SofiaStatus.CERTIFICADO.value,
            )
            .one()
        )

        contrato_aprendizaje = int(
            result.contrato_aprendizaje or 0
        )

        contrato_vinculo_formativo = int(
            result.contrato_vinculo_formativo or 0
        )

        vinculo_laboral = int(
            result.vinculo_laboral or 0
        )

        proyecto_productivo = int(
            result.proyecto_productivo or 0
        )

        monitoria = int(
            result.monitoria or 0
        )

        practicas_economia_popular = int(
            result.practicas_economia_popular or 0
        )

    # -------------------------------------------------------------------------
    # Resultado
    # -------------------------------------------------------------------------

    return SimpleNamespace(
        total=total,

        habilitados=habilitados,
        con_alternativa=con_alternativa,
        sin_alternativa=sin_alternativa,
        certificados=certificados,

        contrato_aprendizaje=contrato_aprendizaje,
        contrato_vinculo_formativo=contrato_vinculo_formativo,
        vinculo_laboral=vinculo_laboral,
        proyecto_productivo=proyecto_productivo,
        monitoria=monitoria,
        practicas_economia_popular=practicas_economia_popular,
    )

# =============================================================================
# Crear grupo
# =============================================================================

@groups_bp.route(
    "/create",
    methods=["GET", "POST"],
)
@login_required
@permission_required("groups.manage")
def create():

    if not _require_group_management():
        return redirect(url_for("groups.index"))

    catalogs = _group_catalogs()

    if request.method == "POST":

        data = _group_form_data()

        if (
            not data["group_number"]
            or not data["program_name"]
        ):
            flash(
                "Número de ficha y nombre del programa son obligatorios.",
                "warning",
            )

            return render_template(
                "groups/create.html",
                form=request.form,
                editing=False,
                group=None,
                **catalogs,
            )

        group = TrainingGroup()

        for field, value in data.items():
            if hasattr(group, field):
                setattr(
                    group,
                    field,
                    value or None,
                )

        if hasattr(group, "created_by"):
            group.created_by = current_user.id

        # Un instructor normal solo puede crear una ficha dentro de su
        # alcance. Si no se indica seguimiento, se asigna al propio instructor.
        if _is_followup_instructor():
            assigned = (getattr(group, "followup_instructor", None) or "").strip()
            identity = _scope_group_identity()
            if not assigned:
                group.followup_instructor = identity
            elif not identity or assigned.casefold() != identity.casefold():
                flash(
                    "Un instructor solo puede crear fichas asignadas a su propio seguimiento.",
                    "danger",
                )
                return render_template(
                    "groups/create.html",
                    form=request.form,
                    editing=False,
                    group=None,
                    **catalogs,
                )

        try:
            db.session.add(group)
            db.session.flush()

            ensure_template_activities_for_group(group)

            db.session.commit()

            flash(
                "Ficha creada correctamente.",
                "success",
            )

            return redirect(
                url_for("groups.index")
            )

        except Exception:
            db.session.rollback()

            current_app.logger.exception(
                "Error creando la ficha."
            )

            flash(
                "Ocurrió un error al crear la ficha.",
                "danger",
            )

            return render_template(
                "groups/create.html",
                form=request.form,
                editing=False,
                group=None,
                **catalogs,
            )

    return render_template(
        "groups/create.html",
        editing=False,
        group=None,
        **catalogs,
    )


# =============================================================================
# Editar grupo
# =============================================================================

@groups_bp.route(
    "/<int:id>/edit",
    methods=["GET", "POST"],
)
@login_required
@permission_required("groups.manage")
def edit(id):

    group = TrainingGroup.query.get_or_404(id)

    if not _can_manage_group(group):
        flash(
            "No tienes permisos para editar este grupo.",
            "warning",
        )

        return redirect(
            url_for("groups.index")
        )

    catalogs = _group_catalogs()

    if request.method == "POST":

        data = _group_form_data()

        if (
            not data["group_number"]
            or not data["program_name"]
        ):
            flash(
                "Número de ficha y nombre del programa son obligatorios.",
                "warning",
            )

            return render_template(
                "groups/create.html",
                editing=True,
                group=group,
                form=request.form,
                **catalogs,
            )

        for field, value in data.items():
            if hasattr(group, field):
                setattr(
                    group,
                    field,
                    value or None,
                )

        if hasattr(group, "updated_by"):
            group.updated_by = current_user.id

        try:
            db.session.commit()

            flash(
                "Ficha actualizada correctamente.",
                "success",
            )

            return redirect(
                url_for("groups.index")
            )

        except Exception:
            db.session.rollback()

            current_app.logger.exception(
                "Error actualizando la ficha %s.",
                group.group_number,
            )

            flash(
                "No fue posible actualizar la ficha.",
                "danger",
            )

            return render_template(
                "groups/create.html",
                editing=True,
                group=group,
                form=request.form,
                **catalogs,
            )

    return render_template(
        "groups/create.html",
        editing=True,
        group=group,
        **catalogs,
    )


# =============================================================================
# Eliminar grupo
# =============================================================================

@groups_bp.route(
    "/<int:id>/delete",
    methods=["POST"],
)
@login_required
@permission_required("groups.manage")
def delete(id):

    group = TrainingGroup.query.get_or_404(id)

    if not _can_manage_group(group):
        flash(
            "No tienes permisos para eliminar este grupo.",
            "warning",
        )

        return redirect(
            url_for("groups.index")
        )

    associated = (
        Apprentice.query
        .filter_by(group_id=group.id)
        .count()
    )

    try:

        Apprentice.query.filter_by(
            group_id=group.id
        ).update(
            {
                "group_id": None,
            },
            synchronize_session=False,
        )

        db.session.delete(group)
        db.session.commit()

        current_app.logger.info(
            "Ficha %s eliminada. Aprendices afectados: %s.",
            group.group_number,
            associated,
        )

        flash(
            f"Ficha eliminada correctamente. "
            f"Aprendices desasociados: {associated}.",
            "success",
        )

    except Exception:

        db.session.rollback()

        current_app.logger.exception(
            "Error eliminando la ficha %s.",
            id,
        )

        flash(
            "No fue posible eliminar la ficha.",
            "danger",
        )

    return redirect(
        url_for("groups.index")
    )


# =============================================================================
# Eliminación múltiple
# =============================================================================

@groups_bp.route(
    "/bulk-delete",
    methods=["POST"],
)
@login_required
@permission_required("groups.manage")
def bulk_delete():

    if not _require_group_management():
        return redirect(url_for("groups.index"))

    try:
        ids = [
            int(value)
            for value in request.form.getlist("selected_ids")
            if str(value).strip()
        ]
    except Exception:

        flash(
            "Selección inválida.",
            "warning",
        )

        return redirect(
            url_for("groups.index")
        )

    if not ids:

        flash(
            "Selecciona al menos una ficha.",
            "warning",
        )

        return redirect(
            url_for("groups.index")
        )

    try:

        groups = (
            TrainingGroup.query
            .filter(
                TrainingGroup.id.in_(ids)
            )
            .all()
        )

        # Nunca permitir que un instructor borre grupos ajenos.
        if not _is_support():
            groups = [
                group
                for group in groups
                if _can_manage_group(group)
            ]

        if not groups:

            flash(
                "No tienes permisos sobre las fichas seleccionadas.",
                "warning",
            )

            return redirect(
                url_for("groups.index")
            )

        deleted = len(groups)
        affected_apprentices = 0

        for group in groups:

            affected_apprentices += (
                Apprentice.query
                .filter_by(group_id=group.id)
                .count()
            )

            Apprentice.query.filter_by(
                group_id=group.id
            ).update(
                {
                    "group_id": None,
                },
                synchronize_session=False,
            )

            db.session.delete(group)

        db.session.commit()

        current_app.logger.info(
            "Eliminación múltiple de fichas (%s). "
            "Aprendices afectados: %s.",
            deleted,
            affected_apprentices,
        )

        flash(
            f"Se eliminaron {deleted} fichas. "
            f"Aprendices desasociados: {affected_apprentices}.",
            "success",
        )

    except Exception:

        db.session.rollback()

        current_app.logger.exception(
            "Error eliminando fichas."
        )

        flash(
            "No fue posible eliminar las fichas seleccionadas.",
            "danger",
        )

    return redirect(
        url_for("groups.index")
    )


# =============================================================================
# Eliminar todas las fichas
# =============================================================================

@groups_bp.route(
    "/delete-all",
    methods=["POST"],
)
@login_required
@permission_required("groups.manage")
def delete_all():

    if not _require_global_management():
        return redirect(
            url_for("groups.index")
        )

    try:

        total = TrainingGroup.query.count()

        affected_apprentices = (
            Apprentice.query
            .filter(
                Apprentice.group_id.isnot(None)
            )
            .count()
        )

        Apprentice.query.filter(
            Apprentice.group_id.isnot(None)
        ).update(
            {
                "group_id": None,
            },
            synchronize_session=False,
        )

        for group in TrainingGroup.query.all():
            db.session.delete(group)

        db.session.commit()

        current_app.logger.warning(
            "Todas las fichas fueron eliminadas por soporte. "
            "Total: %s.",
            total,
        )

        flash(
            f"Se eliminaron {total} fichas. "
            f"Aprendices desasociados: {affected_apprentices}.",
            "success",
        )

    except Exception:

        db.session.rollback()

        current_app.logger.exception(
            "Error eliminando todas las fichas."
        )

        flash(
            "No fue posible eliminar todas las fichas.",
            "danger",
        )

    return redirect(
        url_for("groups.index")
    )


# =============================================================================
# Recalcular estadísticas persistidas
# =============================================================================

@groups_bp.route(
    "/<int:id>/recalculate-stats",
    methods=["POST"],
)
@login_required
@permission_required("groups.manage")
def recalculate_stats(id):
    """Sincroniza los campos estadísticos cacheados con el cálculo actual."""

    group = TrainingGroup.query.get_or_404(id)

    if not _can_manage_group(group):
        flash("No tienes permisos para recalcular las estadísticas de este grupo.", "warning")
        return redirect(url_for("groups.detail", id=group.id))

    try:
        stats = compute_group_stats(group)
        group.apprentices_statistics = str(stats.total or 0)
        # Los antiguos campos "En lectiva" y "En productiva" ya no son
        # indicadores funcionales de GIA. Mantenerlos en cero evita dejar
        # referencias a atributos eliminados de compute_group_stats().
        group.apprentices_training = "0"
        group.apprentices_enabled = str(stats.habilitados or 0)
        group.apprentices_practice = "0"
        group.apprentices_certified = str(stats.certificados or 0)
        group.learning_contract = str(stats.con_alternativa or 0)
        group.apprentices_without_alternative = str(
            stats.sin_alternativa or 0
        )
        group.productive_modalities = ", ".join(
            label + f": {count}"
            for label, count in (
                (get_label(EpModality, EpModality.CONTRATO_APRENDIZAJE.value), stats.contrato_aprendizaje),
                (get_label(EpModality, EpModality.CONTRATO_VINCULO_FORMATIVO.value), stats.contrato_vinculo_formativo),
                (get_label(EpModality, EpModality.VINCULO_LABORAL.value), stats.vinculo_laboral),
                (get_label(EpModality, EpModality.PROYECTO_PRODUCTIVO.value), stats.proyecto_productivo),
                (get_label(EpModality, EpModality.MONITORIA.value), stats.monitoria),
                (get_label(EpModality, EpModality.PRACTICAS_ECONOMIA_POPULAR.value), stats.practicas_economia_popular),
            )
        )
        db.session.commit()
        flash("Las estadísticas del grupo fueron recalculadas correctamente.", "success")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error recalculando estadísticas del grupo %s", group.group_number)
        flash("No fue posible recalcular las estadísticas del grupo.", "danger")

    return redirect(url_for("groups.detail", id=group.id))


# =============================================================================
# Importar
# =============================================================================

@groups_bp.route(
    "/import",
    methods=["GET", "POST"],
)
@login_required
@permission_required("groups.manage")
def import_groups():

    if not _require_group_management():
        return redirect(
            url_for("groups.index")
        )

    if request.method == "POST":

        file = request.files.get("file")

        if not file:

            flash(
                "No se seleccionó ningún archivo.",
                "warning",
            )

            return redirect(
                url_for("groups.import_groups")
            )

        if not getattr(file, "filename", None):

            flash(
                "El archivo seleccionado no tiene nombre.",
                "warning",
            )

            return redirect(
                url_for("groups.import_groups")
            )

        secure_filename(
            file.filename
        )

        try:

            result = import_reference_workbook(
                file,
                owner_id=current_user.id,
                mode="both",
                group_scope=(
                    None
                    if not _is_followup_instructor()
                    else lambda existing, incoming: (
                        _scope_group_belongs_to_current_instructor(existing)
                        if existing is not None
                        else (
                            (incoming.get("followup_instructor") or "").strip().casefold()
                            == _scope_group_identity().casefold()
                        )
                    )
                ),
            )

            if (
                not result.has_apprentice_sheet
                and not result.has_group_sheet
            ):

                flash(
                    "El archivo no contiene las hojas oficiales.",
                    "warning",
                )

            else:

                flash(
                    f"Importación completada: "
                    f"{result.group_count} fichas y "
                    f"{result.apprentice_count} aprendices.",
                    "success" if not result.errors else "warning",
                )

                for message in result.errors[:5]:

                    flash(
                        message,
                        "warning",
                    )

        except Exception:

            current_app.logger.exception(
                "Error importando archivo."
            )

            flash(
                "Error al procesar el archivo.",
                "danger",
            )

        return redirect(
            url_for("groups.index")
        )

    catalogs = _group_catalogs()

    return render_template(
        "groups/import.html",
        **catalogs,
    )


# =============================================================================
# Exportar
# =============================================================================

@groups_bp.route("/export")
@login_required
def export_groups():

    if _is_apprentice():
        flash(
            "No tienes permisos para exportar grupos.",
            "warning",
        )

        return redirect(
            url_for("groups.index")
        )

    try:

        groups = (
            _filtered_groups_query()
            .order_by(
                TrainingGroup.group_number
            )
            .all()
        )

        apprentices = (
            Apprentice.query.filter(
                Apprentice.group_id.in_(
                    [group.id for group in groups]
                )
            ).all()
            if groups
            else []
        )

        output = export_reference_workbook(
            apprentices,
            groups,
        )

        if isinstance(output, bytes):
            output = BytesIO(output)

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="grupos_y_aprendices.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except Exception:

        current_app.logger.exception(
            "Error exportando grupos."
        )

        flash(
            "No fue posible generar el archivo.",
            "danger",
        )

        return redirect(
            url_for("groups.index")
        )
