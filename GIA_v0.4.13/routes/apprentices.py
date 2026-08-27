# routes/apprentices.py
"""
Rutas del módulo de aprendices de GIA.

Responsabilidades:
- Listar aprendices respetando el alcance del usuario.
- Mostrar detalle del aprendiz.
- Crear y editar aprendices.
- Eliminar aprendices individual o masivamente.
- Importar y exportar información mediante los servicios existentes.
- Preparar información para evidencias y seguimiento.
- Presentar valores de catálogos mediante catalogs.display.get_label().

Reglas de alcance:

- Aprendiz:
    No utiliza este módulo administrativo directamente.
    Su acceso operativo está controlado por app.py y evidences.py.

- Instructor de seguimiento:
    Solo puede consultar y gestionar aprendices pertenecientes a
    grupos cuyo followup_instructor corresponde al usuario autenticado.

- Instructor de seguimiento líder:
    Tiene visión administrativa/global.

- Administrativo del centro:
    Tiene visión global de consulta.

- Certificador:
    Tiene visión global de consulta.

- Soporte:
    Tiene visión global y capacidad administrativa/técnica.

La lógica de evidencias y seguimiento permanece delegada a sus respectivos
servicios.
"""

from __future__ import annotations

from collections import defaultdict
from io import BytesIO

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
from sqlalchemy import func, or_
from sqlalchemy.exc import IntegrityError
from werkzeug.utils import secure_filename

from catalogs.apprentice import EpModality, SofiaStatus
from catalogs.common_catalogs import DocumentType, ProgramLevel
from catalogs.display import get_choices_sorted, get_label

from extensions import db

from models import (
    Apprentice,
    EvidenceSubmission,
    TrainingGroup,
    User,
)

from services.auth_helpers import permission_required
from services.evidence_service import (
    ensure_submissions_for_apprentice,
    summarize_submissions,
)

from services.excel_export import (
    export_reference_workbook,
)

from services.excel_import import (
    APPRENTICE_MODEL_FIELDS as APPRENTICE_FIELDS,
    import_reference_workbook,
    upsert_student_user,
)

from services.followup_service import (
    get_followup_dashboard,
    get_apprentice_followup,
)
from services.date_rules import build_derived_apprentice_dates, calculate_group_validity_from_training_end

from services.access_scope import (
    current_role as _scope_current_role,
    visible_groups_query as _scope_visible_groups_query,
    visible_group_ids as _scope_visible_group_ids,
    visible_apprentices_query as _scope_visible_apprentices_query,
    instructor_identity as _scope_group_identity,
    group_belongs_to_current_instructor as _scope_group_belongs_to_current_instructor,
)

from services.utils import parse_form
from services.permissions import (
    ROLE_APPRENTICE,
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_CENTER_STAFF,
    ROLE_CERTIFIER,
    ROLE_SUPPORT,
    GLOBAL_ROLES,
    APPRENTICE_MANAGEMENT_ROLES,
)


# =============================================================================
# BLUEPRINT
# =============================================================================

apprentices_bp = Blueprint(
    "apprentices",
    __name__,
    url_prefix="/apprentices",
)

# =============================================================================
# Catálogos para formularios
# =============================================================================

def _catalog_choices(catalog):
    return [
        {
            "value": item.value,
            "label": get_label(catalog, item.value),
        }
        for item in catalog
    ]

def _form_catalogs():
    """
    Prepara los catálogos oficiales para las plantillas.

    Las plantillas reciben únicamente:
        value -> valor persistido
        label -> texto de presentación

    La traducción de los valores se mantiene exclusivamente
    en catalogs/display.py.
    """
    document_types = _catalog_choices(DocumentType)
    program_levels = _catalog_choices(ProgramLevel)
    ep_modalities = _catalog_choices(EpModality)

    return {
        "document_types": document_types,
        "program_levels": program_levels,
        "ep_modalities": ep_modalities,
    }
    
    
# =============================================================================
# ROLES CANÓNICOS / ALCANCE DELEGADO
# =============================================================================

MANAGE_ROLES = APPRENTICE_MANAGEMENT_ROLES


def _current_role() -> str | None:
    return _scope_current_role()


def _is_global_role() -> bool:
    return _current_role() in GLOBAL_ROLES


def _is_followup_instructor() -> bool:
    return _current_role() == ROLE_FOLLOWUP_INSTRUCTOR


def _is_lead_followup_instructor() -> bool:
    return _current_role() == ROLE_LEAD_FOLLOWUP_INSTRUCTOR


def _can_manage_apprentices() -> bool:
    return _current_role() in MANAGE_ROLES


def _followup_group_query():
    return _scope_visible_groups_query()


def _visible_groups_query():
    return _scope_visible_groups_query()


def _visible_group_ids() -> list[int]:
    return _scope_visible_group_ids()


# =============================================================================
# ALCANCE POR APRENDICES
# =============================================================================

def _visible_apprentices_query():
    return _scope_visible_apprentices_query()


def _get_visible_apprentice_or_404(apprentice_id: int):
    """
    Obtiene un aprendiz únicamente si pertenece al alcance
    del usuario autenticado.
    """

    apprentice = (
        _visible_apprentices_query()
        .filter(Apprentice.id == apprentice_id)
        .first()
    )

    if apprentice is None:
        from flask import abort

        abort(404)

    return apprentice


# =============================================================================
# ETIQUETAS DE CATÁLOGOS
# =============================================================================

def _catalog_choices(catalog):
    """
    Devuelve opciones de catálogo ordenadas por label.

    La plantilla recibe:
        value
        label

    y nunca necesita conocer los textos internos del catálogo.
    """

    try:
        return [
            {
                "value": value.value,
                "label": get_label(catalog, value),
            }
            for value, _label in get_choices_sorted(catalog)
        ]
    except Exception:
        current_app.logger.exception(
            "Error preparando opciones del catálogo %s.",
            getattr(catalog, "__name__", str(catalog)),
        )
        return []


def _presentation_catalogs():
    """
    Catálogos utilizados por apprentices/index.html.

    Los valores son canónicos y las etiquetas provienen exclusivamente
    de catalogs.display.
    """

    return {
        "document_types": _catalog_choices(DocumentType),
        "program_levels": _catalog_choices(ProgramLevel),
        "ep_modalities": _catalog_choices(EpModality),
        "sofia_statuses": _catalog_choices(SofiaStatus),
    }

def _detail_catalog_labels(apprentice):
    """
    Prepara las etiquetas de presentación de los catálogos para
    templates/apprentices/detail.html.

    La plantilla nunca debe interpretar directamente los valores
    canónicos almacenados en la base de datos.
    """
    return {
        "document_type_label": (
            get_label(DocumentType, apprentice.document_type)
            if getattr(apprentice, "document_type", None)
            else None
        ),
        "program_level_label": (
            get_label(ProgramLevel, apprentice.program_level)
            if getattr(apprentice, "program_level", None)
            else None
        ),
        "ep_modality_label": (
            get_label(EpModality, apprentice.ep_modality)
            if getattr(apprentice, "ep_modality", None)
            else None
        ),
        "sofia_status_label": (
            get_label(SofiaStatus, apprentice.sofia_status)
            if getattr(apprentice, "sofia_status", None)
            else None
        ),
    }

# =============================================================================
# LISTADO
# =============================================================================

@apprentices_bp.route("/")
@login_required
def index():
    """
    Listado principal de aprendices.

    La consulta ya está limitada por el rol antes de aplicar
    los filtros enviados por GET.
    """

    query = _filtered_apprentices_query()

    apprentices = (
        query
        .order_by(
            Apprentice.last_names,
            Apprentice.first_names,
        )
        .all()
    )

    try:
        total_apprentices = (
            _visible_apprentices_query().count()
        )
    except Exception:
        current_app.logger.exception(
            "Error obteniendo total de aprendices visibles."
        )
        total_apprentices = len(apprentices)

    try:
        group_numbers = [
            value
            for (value,) in (
                _visible_apprentices_query()
                .with_entities(Apprentice.group_number)
                .filter(
                    Apprentice.group_number.isnot(None)
                )
                .distinct()
                .order_by(Apprentice.group_number)
                .all()
            )
            if value
        ]
    except Exception:
        current_app.logger.exception(
            "Error obteniendo fichas visibles."
        )
        group_numbers = []

    catalogs = _presentation_catalogs()

    return render_template(
        "apprentices/index.html",
        apprentices=apprentices,
        total_apprentices=total_apprentices,
        group_numbers=group_numbers,
        document_types=catalogs["document_types"],
        program_levels=catalogs["program_levels"],
        ep_modalities=catalogs["ep_modalities"],
        sofia_statuses=catalogs["sofia_statuses"],
        can_manage=_can_manage_apprentices(),
        is_global=_is_global_role(),
        is_followup_instructor=_is_followup_instructor(),
        is_lead_followup_instructor=_is_lead_followup_instructor(),
    )


# =============================================================================
# CONSULTA FILTRADA
# =============================================================================

def _filtered_apprentices_query():
    """
    Construye la consulta de aprendices respetando primero
    el alcance del usuario y después los filtros.
    """

    query = _visible_apprentices_query()

    search = (
        request.args.get("search", "")
        .strip()
    )

    group_number = (
        request.args.get("group_number", "")
        .strip()
    )

    ep_modality = (
        request.args.get("ep_modality", "")
        .strip()
    )

    status = (
        request.args.get("status", "")
        .strip()
    )

    program_level = (
        request.args.get("program_level", "")
        .strip()
    )

    if search:
        search = search[:300]
        pattern = f"%{search}%"

        query = query.filter(
            or_(
                Apprentice.first_names.ilike(pattern),
                Apprentice.last_names.ilike(pattern),
                Apprentice.document_number.ilike(pattern),
                Apprentice.email.ilike(pattern),
                Apprentice.program_name.ilike(pattern),
                Apprentice.group_number.ilike(pattern),
            )
        )

    if group_number:
        query = query.filter(
            Apprentice.group_number == group_number
        )

    if ep_modality:
        query = query.filter(
            Apprentice.ep_modality == ep_modality
        )

    if status:
        query = query.filter(
            Apprentice.sofia_status == status
        )

    if program_level:
        query = query.filter(
            Apprentice.program_level == program_level
        )

    return query


# =============================================================================
# DETALLE DEL APRENDIZ
# =============================================================================

@apprentices_bp.route("/<int:id>")
@login_required
def detail(id):
    """
    Muestra el detalle de un aprendiz respetando el alcance del
    usuario autenticado.

    La ruta prepara toda la información que necesita
    templates/apprentices/detail.html:

    - record
    - etiquetas de catálogos
    - permisos de edición/eliminación
    - evidencias
    - resumen de evidencias
    - evidencias agrupadas
    - información de seguimiento
    """

    # =========================================================================
    # 1. ALCANCE
    # =========================================================================

    apprentice = _get_visible_apprentice_or_404(id)

    role = _current_role()

    # =========================================================================
    # 2. IDENTIDAD / NOMBRE
    # =========================================================================

    full_name = (
        getattr(apprentice, "full_name", None)
        or " ".join(
            part
            for part in (
                getattr(apprentice, "first_names", None),
                getattr(apprentice, "last_names", None),
            )
            if part
        )
    ).strip()

    if not full_name:
        full_name = f"Aprendiz {apprentice.id}"

    # =========================================================================
    # 3. ETIQUETAS DE CATÁLOGOS
    # =========================================================================

    catalog_labels = _detail_catalog_labels(apprentice)

    # Fechas derivadas para presentación: no sobrescriben datos importados.
    derived_dates = build_derived_apprentice_dates({
        "practice_start_date": getattr(apprentice, "practice_start_date", None),
        "practice_end_date": getattr(apprentice, "practice_end_date", None),
    })
    for _key, _value in derived_dates.items():
        if _key.startswith("followup_moment") and _value:
            setattr(apprentice, _key, _value)

    display_group_validity = getattr(apprentice, "group_validity", None)
    if not display_group_validity and getattr(apprentice, "group", None):
        display_group_validity = getattr(apprentice.group, "group_validity", None)
        if not display_group_validity:
            display_group_validity = calculate_group_validity_from_training_end(
                getattr(apprentice.group, "training_end_date", None)
            )

    # =========================================================================
    # 4. PERMISOS
    # =========================================================================
    #
    # La autorización real continúa estando en las rutas.
    # Estas variables solamente permiten que el template sepa qué
    # acciones puede mostrar.
    #

    can_edit = (
        _can_manage_apprentices()
        and "apprentices.edit" in current_app.view_functions
    )

    can_delete = (
        _can_manage_apprentices()
        and "apprentices.delete" in current_app.view_functions
    )

    can_manage = _can_manage_apprentices()

    is_global = _is_global_role()
    is_followup_instructor = _is_followup_instructor()
    is_lead_followup_instructor = _is_lead_followup_instructor()

    # =========================================================================
    # 5. URL DE EDICIÓN
    # =========================================================================

    edit_url = None

    if can_edit:
        edit_url = url_for(
            "apprentices.edit",
            id=apprentice.id,
        )

    # =========================================================================
    # 6. URL DE ELIMINACIÓN
    # =========================================================================

    delete_url = None

    if can_delete:
        delete_url = url_for(
            "apprentices.delete",
            id=apprentice.id,
        )

    # =========================================================================
    # 7. EVIDENCIAS
    # =========================================================================

    try:
        ensure_submissions_for_apprentice(apprentice)
        db.session.commit()

        evidence_submissions = (
            EvidenceSubmission.query
            .filter_by(
                apprentice_id=apprentice.id
            )
            .all()
        )

        evidence_summary = summarize_submissions(
            evidence_submissions
        )

    except Exception:
        db.session.rollback()

        current_app.logger.exception(
            "Error preparando evidencias del aprendiz %s.",
            apprentice.id,
        )

        evidence_submissions = []
        evidence_summary = {}

    # =========================================================================
    # 8. AGRUPACIÓN DE EVIDENCIAS
    # =========================================================================

    grouped = defaultdict(list)

    for submission in evidence_submissions:

        try:
            category_name = (
                getattr(getattr(submission.activity, "category", None), "name", None)
                or "Sin categoría"
            ).strip()

        except Exception:
            category_name = "Sin categoría"

        grouped[category_name].append(
            submission
        )

    evidence_groups = {}

    for category_name, submissions in grouped.items():

        total_files = 0
        approved = 0
        pending = 0

        for submission in submissions:

            evidences = (
                getattr(
                    submission,
                    "evidences",
                    [],
                )
                or []
            )

            total_files += len(evidences)

            for evidence in evidences:

                status = (
                    getattr(
                        evidence,
                        "status",
                        "",
                    )
                    or ""
                ).strip().lower()

                if status in {
                    "approved",
                    "aprobado",
                }:
                    approved += 1

                elif status in {
                    "pending",
                    "pendiente",
                }:
                    pending += 1

        evidence_groups[category_name] = {
            "submissions": submissions,
            "total_files": total_files,
            "approved": approved,
            "pending": pending,
        }

    # =========================================================================
    # 9. SEGUIMIENTO
    # =========================================================================

    try:
        followup_dashboard = get_followup_dashboard(apprentice)
        followup_details = get_apprentice_followup(apprentice)

    except Exception:
        current_app.logger.exception(
            "Error obteniendo seguimiento del aprendiz %s.",
            apprentice.id,
        )

        followup_dashboard = {}
        followup_details = []

    # =========================================================================
    # 10. URL DE RETORNO
    # =========================================================================

    back_url = url_for(
        "apprentices.index"
    )

    # =========================================================================
    # 11. RENDER
    # =========================================================================

    return render_template(
        "apprentices/detail.html",

        # Registro principal
        record=apprentice,

        # Identificación
        page_title=f"Detalle - {full_name}",
        detail_type="apprentice",

        # Navegación
        back_url=back_url,
        edit_url=edit_url,
        delete_url=delete_url,

        # Catálogos: presentación
        document_type_label=catalog_labels["document_type_label"],
        program_level_label=catalog_labels["program_level_label"],
        ep_modality_label=catalog_labels["ep_modality_label"],
        sofia_status_label=catalog_labels["sofia_status_label"],

        # Permisos / contexto de rol
        can_edit=can_edit,
        can_delete=can_delete,
        can_manage=can_manage,
        is_global=is_global,
        is_followup_instructor=is_followup_instructor,
        is_lead_followup_instructor=is_lead_followup_instructor,
        current_role=role,

        # Evidencias
        evidence_submissions=evidence_submissions,
        evidence_summary=evidence_summary,
        evidence_groups=evidence_groups,

        # Seguimiento
        followup_dashboard=followup_dashboard,
        followup_details=followup_details,
        display_group_validity=display_group_validity,
    )

    # -------------------------------------------------------------------------
    # Evidencias
    # -------------------------------------------------------------------------

    try:
        ensure_submissions_for_apprentice(apprentice)
        db.session.commit()

        evidence_submissions = (
            EvidenceSubmission.query
            .filter_by(
                apprentice_id=apprentice.id
            )
            .all()
        )

        evidence_summary = summarize_submissions(
            evidence_submissions
        )

    except Exception:
        db.session.rollback()

        current_app.logger.exception(
            "Error preparando evidencias del aprendiz %s.",
            apprentice.id,
        )

        evidence_submissions = []
        evidence_summary = {}

    # -------------------------------------------------------------------------
    # Agrupar evidencias
    # -------------------------------------------------------------------------

    grouped = defaultdict(list)

    for submission in evidence_submissions:
        try:
            category_name = (
                getattr(getattr(submission.activity, "category", None), "name", None)
                or "Sin categoría"
            ).strip()
        except Exception:
            category_name = "Sin categoría"

        grouped[category_name].append(
            submission
        )

    evidence_groups = {}

    for category_name, submissions in grouped.items():

        total_files = 0
        approved = 0
        pending = 0

        for submission in submissions:

            evidences = (
                getattr(
                    submission,
                    "evidences",
                    [],
                )
                or []
            )

            total_files += len(evidences)

            for evidence in evidences:

                status = (
                    getattr(
                        evidence,
                        "status",
                        "",
                    )
                    or ""
                ).strip().lower()

                if status in {
                    "approved",
                    "aprobado",
                }:
                    approved += 1

                elif status in {
                    "pending",
                    "pendiente",
                }:
                    pending += 1

        evidence_groups[category_name] = {
            "submissions": submissions,
            "total_files": total_files,
            "approved": approved,
            "pending": pending,
        }

    # -------------------------------------------------------------------------
    # Seguimiento
    # -------------------------------------------------------------------------

    try:
        followup_dashboard = get_followup_dashboard(apprentice)
        followup_details = get_apprentice_followup(apprentice)
    except Exception:
        current_app.logger.exception(
            "Error obteniendo seguimiento del aprendiz %s.",
            apprentice.id,
        )
        followup_dashboard = {}

    # -------------------------------------------------------------------------
    # Datos auxiliares
    # -------------------------------------------------------------------------

    full_name = (
        getattr(
            apprentice,
            "full_name",
            None,
        )
        or f"{apprentice.first_names} {apprentice.last_names}"
    ).strip()

    if not full_name:
        full_name = f"Aprendiz {apprentice.id}"

    edit_url = None

    if (
        _can_manage_apprentices()
        and "apprentices.edit" in current_app.view_functions
    ):
        edit_url = url_for(
            "apprentices.edit",
            id=apprentice.id,
        )

    back_url = url_for(
        "apprentices.index"
    )

    return render_template(
        "apprentices/detail.html",
        record=apprentice,
        page_title=f"Detalle - {full_name}",
        detail_type="apprentice",
        back_url=back_url,
        edit_url=edit_url,
        evidence_submissions=evidence_submissions,
        evidence_summary=evidence_summary,
        evidence_groups=evidence_groups,
        followup_dashboard=followup_dashboard,
    )


# =============================================================================
# CREAR APRENDIZ
# =============================================================================

@apprentices_bp.route(
    "/create",
    methods=["GET", "POST"],
)
@login_required
@permission_required("apprentices.manage")
def create():
    """
    Crear aprendiz.

    Permitido para:
    - Instructor de seguimiento.
    - Instructor de seguimiento líder.
    - Soporte.
    """

    catalogs = _form_catalogs()
    

    if not _can_manage_apprentices():
        flash(
            "No tienes permisos para crear aprendices.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    if request.method == "POST":

        try:
            form_data = parse_form(
                request.form,
                APPRENTICE_FIELDS,
            )

        except Exception:
            form_data = {
                "group_number": request.form.get(
                    "group_number",
                    "",
                ).strip() or None,

                "first_names": request.form.get(
                    "first_names",
                    "",
                ).strip() or None,

                "last_names": request.form.get(
                    "last_names",
                    "",
                ).strip() or None,

                "document_type": request.form.get(
                    "document_type",
                    "",
                ).strip() or None,

                "document_number": request.form.get(
                    "document_number",
                    "",
                ).strip() or None,

                "email": request.form.get(
                    "email",
                    "",
                ).strip() or None,

                "phone": request.form.get(
                    "phone",
                    "",
                ).strip() or None,

                "municipality_origin": request.form.get(
                    "municipality_origin",
                    "",
                ).strip() or None,

                "program_name": request.form.get(
                    "program_name",
                    "",
                ).strip() or None,

                "program_level": request.form.get(
                    "program_level",
                    "",
                ).strip() or None,

                "lead_instructor": request.form.get(
                    "lead_instructor",
                    "",
                ).strip() or None,

                "followup_instructor": request.form.get(
                    "followup_instructor",
                    "",
                ).strip() or None,

                "ep_modality": request.form.get(
                    "ep_modality",
                    "",
                ).strip() or None,

                "practice_start_date": request.form.get(
                    "practice_start_date",
                    "",
                ).strip() or None,

                "practice_end_date": request.form.get(
                    "practice_end_date",
                    "",
                ).strip() or None,
            }

        if not (
            form_data.get("first_names")
            or form_data.get("last_names")
        ):
            flash(
                "El nombre del aprendiz es obligatorio.",
                "warning",
            )

            return render_template(
                "apprentices/create.html",
                form=request.form,
            )

        if not form_data.get("group_number"):
            flash(
                "La ficha es obligatoria.",
                "warning",
            )

            return render_template(
                "apprentices/create.html",
                form=request.form,
            )

        group = (
            TrainingGroup.query
            .filter_by(
                group_number=form_data["group_number"]
            )
            .first()
        )

        if not group:
            flash(
                f"No existe la ficha {form_data['group_number']}.",
                "warning",
            )

            return render_template(
                "apprentices/create.html",
                form=request.form,
                editing=False,
                apprentice=None,
                **catalogs,
            )

        # Instructor normal solamente puede asignar aprendices
        # a uno de sus grupos.
        if _is_followup_instructor():

            visible_group_ids = set(
                _visible_group_ids()
            )

            if group.id not in visible_group_ids:
                flash(
                    "No puedes asignar un aprendiz a una ficha que no tienes asignada.",
                    "danger",
                )

                return render_template(
                    "apprentices/create.html",
                    form=request.form,
                    editing=False,
                    apprentice=None,
                    **catalogs,
                )

        form_data["group_id"] = group.id

        apprentice = Apprentice()

        for field, value in form_data.items():

            if hasattr(apprentice, field):
                setattr(
                    apprentice,
                    field,
                    value,
                )

        if hasattr(apprentice, "created_by"):
            apprentice.created_by = current_user.id

        try:
            db.session.add(apprentice)
            db.session.flush()

            upsert_student_user(
                apprentice
            )

            ensure_submissions_for_apprentice(
                apprentice
            )

            db.session.commit()

            flash(
                "Aprendiz creado correctamente.",
                "success",
            )

            return redirect(
                url_for("apprentices.index")
            )

        except IntegrityError:
            db.session.rollback()

            current_app.logger.exception(
                "IntegrityError creando aprendiz."
            )

            flash(
                "No fue posible crear el aprendiz. Verifica los datos únicos.",
                "danger",
            )

        except Exception:
            db.session.rollback()

            current_app.logger.exception(
                "Error creando aprendiz."
            )

            flash(
                "Ocurrió un error al crear el aprendiz.",
                "danger",
            )

    return render_template(
        "apprentices/create.html" ,
        editing=False,
        apprentice=None,
        **catalogs,
    )

    


# =============================================================================
# EDITAR APRENDIZ
# =============================================================================

@apprentices_bp.route(
    "/<int:id>/edit",
    methods=["GET", "POST"],
)
@login_required
@permission_required("apprentices.manage")
def edit(id):

    apprentice = Apprentice.query.get_or_404(id)
    
    catalogs = _form_catalogs()

    """
    Editar aprendiz respetando el alcance del usuario.
    """

    if not _can_manage_apprentices():
        flash(
            "No tienes permisos para editar aprendices.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    apprentice = _get_visible_apprentice_or_404(
        id
    )

    if request.method == "POST":

        try:
            form_data = parse_form(
                request.form,
                APPRENTICE_FIELDS,
            )

        except Exception:
            form_data = {
                "first_names": request.form.get(
                    "first_names",
                    "",
                ).strip() or None,

                "last_names": request.form.get(
                    "last_names",
                    "",
                ).strip() or None,

                "document_type": request.form.get(
                    "document_type",
                    "",
                ).strip() or None,

                "document_number": request.form.get(
                    "document_number",
                    "",
                ).strip() or None,

                "email": request.form.get(
                    "email",
                    "",
                ).strip() or None,

                "phone": request.form.get(
                    "phone",
                    "",
                ).strip() or None,

                "municipality_origin": request.form.get(
                    "municipality_origin",
                    "",
                ).strip() or None,

                "program_name": request.form.get(
                    "program_name",
                    "",
                ).strip() or None,

                "program_level": request.form.get(
                    "program_level",
                    "",
                ).strip() or None,

                "lead_instructor": request.form.get(
                    "lead_instructor",
                    "",
                ).strip() or None,

                "followup_instructor": request.form.get(
                    "followup_instructor",
                    "",
                ).strip() or None,

                "group_number": request.form.get(
                    "group_number",
                    "",
                ).strip() or None,

                "ep_modality": request.form.get(
                    "ep_modality",
                    "",
                ).strip() or None,
            }

        group = None

        if form_data.get("group_number"):

            group = (
                TrainingGroup.query
                .filter_by(
                    group_number=form_data["group_number"]
                )
                .first()
            )

            if not group:
                flash(
                    f"No existe la ficha {form_data['group_number']}.",
                    "warning",
                )

                return render_template(
                    "apprentices/create.html",
                    editing=True,
                    apprentice=apprentice,
                    form=request.form,
                    **catalogs,
                )

            if _is_followup_instructor():

                visible_group_ids = set(
                    _visible_group_ids()
                )

                if group.id not in visible_group_ids:
                    flash(
                        "No puedes mover el aprendiz a una ficha que no tienes asignada.",
                        "danger",
                    )

                    return render_template(
                        "apprentices/create.html",
                        editing=True,
                        apprentice=apprentice,
                        form=request.form,
                    )

            form_data["group_id"] = group.id

        for field, value in form_data.items():

            if hasattr(apprentice, field):
                setattr(
                    apprentice,
                    field,
                    value,
                )

        try:
            db.session.commit()

            flash(
                "Aprendiz actualizado correctamente.",
                "success",
            )

            return redirect(
                url_for(
                    "apprentices.detail",
                    id=apprentice.id,
                )
            )

        except IntegrityError:
            db.session.rollback()

            current_app.logger.exception(
                "IntegrityError actualizando aprendiz."
            )

            flash(
                "No fue posible actualizar el aprendiz. Verifica los datos.",
                "danger",
            )

        except Exception:
            db.session.rollback()

            current_app.logger.exception(
                "Error actualizando aprendiz."
            )

            flash(
                "Ocurrió un error al actualizar el aprendiz.",
                "danger",
            )

    return render_template(
        "apprentices/create.html",
        editing=True,
        apprentice=apprentice,
        **catalogs,
    )


# =============================================================================
# ELIMINAR APRENDIZ
# =============================================================================

@apprentices_bp.route(
    "/<int:id>/delete",
    methods=["POST"],
)
@login_required
@permission_required("apprentices.manage")
def delete(id):
    """
    Eliminar un aprendiz respetando el alcance del usuario.
    """

    if not _can_manage_apprentices():
        flash(
            "No tienes permisos para eliminar aprendices.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    apprentice = _get_visible_apprentice_or_404(
        id
    )

    try:

        student_user_id = (
            apprentice.student_user_id
        )

        document_number = (
            apprentice.document_number
        )

        if student_user_id:

            student_user = User.query.get(
                student_user_id
            )

            if (
                student_user
                and student_user.role == ROLE_APPRENTICE
            ):

                references = (
                    Apprentice.query
                    .filter(
                        Apprentice.student_user_id
                        == student_user.id,
                        Apprentice.id
                        != apprentice.id,
                    )
                    .count()
                )

                if references == 0:
                    db.session.delete(
                        student_user
                    )

        db.session.delete(
            apprentice
        )

        db.session.commit()

        current_app.logger.info(
            "Aprendiz eliminado: id=%s documento=%s usuario=%s",
            apprentice.id,
            document_number,
            current_user.id,
        )

        flash(
            "Aprendiz eliminado correctamente.",
            "success",
        )

    except IntegrityError:
        db.session.rollback()

        current_app.logger.exception(
            "IntegrityError eliminando aprendiz."
        )

        flash(
            "No fue posible eliminar el aprendiz por restricciones de la base de datos.",
            "danger",
        )

    except Exception:
        db.session.rollback()

        current_app.logger.exception(
            "Error eliminando aprendiz."
        )

        flash(
            "Ocurrió un error al eliminar el aprendiz.",
            "danger",
        )

    return redirect(
        url_for("apprentices.index")
    )


# =============================================================================
# ELIMINACIÓN MASIVA
# =============================================================================

@apprentices_bp.route(
    "/bulk-delete",
    methods=["POST"],
)
@login_required
@permission_required("apprentices.manage")
def bulk_delete():
    """
    Eliminación masiva respetando el alcance del usuario.
    """

    if not _can_manage_apprentices():
        flash(
            "No tienes permisos para eliminar aprendices.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    try:
        ids = [
            int(value)
            for value in request.form.getlist(
                "selected_ids"
            )
            if value.strip()
        ]

    except Exception:
        flash(
            "Selección inválida.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    if not ids:
        flash(
            "Selecciona al menos un aprendiz.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    try:

        apprentices = (
            _visible_apprentices_query()
            .filter(
                Apprentice.id.in_(ids)
            )
            .all()
        )

        if not apprentices:
            flash(
                "Ninguno de los aprendices seleccionados pertenece a tu alcance.",
                "warning",
            )
            return redirect(
                url_for("apprentices.index")
            )

        for apprentice in apprentices:

            if apprentice.student_user_id:

                student_user = User.query.get(
                    apprentice.student_user_id
                )

                if (
                    student_user
                    and student_user.role
                    == ROLE_APPRENTICE
                ):

                    references = (
                        Apprentice.query
                        .filter(
                            Apprentice.student_user_id
                            == student_user.id,
                            Apprentice.id
                            != apprentice.id,
                        )
                        .count()
                    )

                    if references == 0:
                        db.session.delete(
                            student_user
                        )

            db.session.delete(
                apprentice
            )

        db.session.commit()

        current_app.logger.info(
            "Aprendices eliminados en lote: total=%s usuario=%s",
            len(apprentices),
            current_user.id,
        )

        flash(
            f"Se eliminaron {len(apprentices)} aprendices.",
            "success",
        )

    except Exception:
        db.session.rollback()

        current_app.logger.exception(
            "Error eliminando aprendices en lote."
        )

        flash(
            "No fue posible eliminar los aprendices seleccionados.",
            "danger",
        )

    return redirect(
        url_for("apprentices.index")
    )


# =============================================================================
# ELIMINAR TODOS
# =============================================================================

@apprentices_bp.route(
    "/delete-all",
    methods=["POST"],
)
@login_required
@permission_required("apprentices.manage")
def delete_all():
    """
    Elimina todos los aprendices visibles.

    Se conserva principalmente para Soporte.
    Un instructor de seguimiento nunca puede ejecutar
    una eliminación global.
    """

    if _current_role() != ROLE_SUPPORT:
        flash(
            "No tienes permisos para eliminar todos los aprendices.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    try:

        apprentices = (
            Apprentice.query
            .all()
        )

        total = len(apprentices)

        for apprentice in apprentices:

            if apprentice.student_user_id:

                student_user = User.query.get(
                    apprentice.student_user_id
                )

                if (
                    student_user
                    and student_user.role
                    == ROLE_APPRENTICE
                ):

                    references = (
                        Apprentice.query
                        .filter(
                            Apprentice.student_user_id
                            == student_user.id,
                            Apprentice.id
                            != apprentice.id,
                        )
                        .count()
                    )

                    if references == 0:
                        db.session.delete(
                            student_user
                        )

            db.session.delete(
                apprentice
            )

        db.session.commit()

        current_app.logger.warning(
            "Todos los aprendices eliminados: total=%s usuario=%s",
            total,
            current_user.id,
        )

        flash(
            f"Se eliminaron {total} aprendices.",
            "success",
        )

    except Exception:
        db.session.rollback()

        current_app.logger.exception(
            "Error eliminando todos los aprendices."
        )

        flash(
            "No fue posible eliminar todos los aprendices.",
            "danger",
        )

    return redirect(
        url_for("apprentices.index")
    )


# =============================================================================
# IMPORTAR EXCEL
# =============================================================================

@apprentices_bp.route(
    "/import",
    methods=["GET", "POST"],
)
@login_required
@permission_required("apprentices.manage")
def import_excel():
    """
    Importación de los libros oficiales de GIA.

    La operación queda disponible para los roles de gestión.
    """

    if not _can_manage_apprentices():
        flash(
            "No tienes permisos para importar aprendices.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    if request.method == "POST":

        file = request.files.get(
            "file"
        )

        if not file:
            flash(
                "No se seleccionó ningún archivo.",
                "warning",
            )
            return redirect(
                url_for("apprentices.import_excel")
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
                    if _current_role() != ROLE_FOLLOWUP_INSTRUCTOR
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

            current_app.logger.error(
                "ROUTE DEBUG result=%r type=%s errors=%r errors_type=%s",
                result,
                type(result),
                result.errors,
                type(result.errors),
            )

            if (
                result.apprentice_count == 0
                and result.group_count == 0
                and result.skipped_apprentices == 0
            ):
                flash(
                    "El archivo no contiene datos importables.",
                    "warning",
                )

            else:

                flash(
                    f"Importación completada: "
                    f"{result.apprentice_count} aprendices y "
                    f"{result.group_count} grupos. "
                    f"Omitidos: {result.skipped_apprentices}.",
                    (
                        "success"
                        if not result.errors
                        else "warning"
                    ),
                )

                for message in result.errors[:5]:
                    flash(
                        message,
                        "warning",
                    )

        except Exception:

            current_app.logger.exception(
                "Error importando aprendices."
            )

            flash(
                "Ocurrió un error al procesar el archivo.",
                "danger",
            )

        return redirect(
            url_for("apprentices.index")
        )

    return render_template(
        "apprentices/import.html"
    )


# =============================================================================
# EXPORTAR EXCEL
# =============================================================================

@apprentices_bp.route(
    "/export"
)
@login_required
def export_all():
    """
    Exporta únicamente los aprendices que el usuario puede consultar.
    """

    if not _is_global_role() and not _is_followup_instructor():
        flash(
            "No tienes permisos para exportar aprendices.",
            "warning",
        )
        return redirect(
            url_for("apprentices.index")
        )

    try:

        apprentices = (
            _filtered_apprentices_query()
            .order_by(
                Apprentice.last_names,
                Apprentice.first_names,
            )
            .all()
        )

        group_numbers = {
            apprentice.group_number
            for apprentice in apprentices
            if apprentice.group_number
        }

        groups = []

        if group_numbers:

            groups = (
                TrainingGroup.query
                .filter(
                    TrainingGroup.group_number.in_(
                        group_numbers
                    )
                )
                .all()
            )

        output = export_reference_workbook(
            apprentices,
            groups,
        )

        if isinstance(output, bytes):

            output = BytesIO(
                output
            )

            output.seek(0)

        elif hasattr(output, "seek"):

            try:
                output.seek(0)
            except Exception:
                pass

        else:

            output = BytesIO(
                output
            )

            output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="referencias_aprendices.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except Exception:

        flash(
            "No fue posible generar el archivo de exportación.",
            "danger",
        )

        return redirect(
            url_for("apprentices.index")
        )
