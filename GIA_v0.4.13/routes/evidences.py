# =============================================================================
# routes/evidences.py
# =============================================================================

"""
Rutas del módulo de Evidencias.

Este módulo coordina las peticiones HTTP del módulo de evidencias.

La lógica de dominio pertenece a:

- models/evidence.py
- services/evidence_service.py
- services/followup_service.py
- services/pdf_service.py

Responsabilidades principales:

- consultar evidencias;
- visualizar una entrega;
- cargar/reemplazar una entrega;
- registrar observaciones;
- solicitar correcciones;
- aprobar evidencias;
- descargar archivos;
- visualizar archivos;

Estados soportados:

- no_entregado
- pendiente_revision
- requiere_correccion
- aprobado

Roles considerados:

- APPRENTICE
- FOLLOW_UP_INSTRUCTOR
- FOLLOW_UP_INSTRUCTOR_lider
- CERTIFIER
- SUPPORT

"""

from __future__ import annotations

import os
from io import BytesIO
from datetime import datetime, timezone

from flask import (
    Blueprint,
    abort,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)

from flask_login import current_user, login_required

from sqlalchemy.orm import joinedload

from extensions import db

from models import (
    Apprentice,
    EvidenceCategory,
    EvidenceActivity,
    EvidenceSubmission,
    EvidenceComment,
    EVIDENCE_STATUS_NOT_SUBMITTED,
    EVIDENCE_STATUS_PENDING_REVIEW,
    EVIDENCE_STATUS_REQUIRES_CORRECTION,
    EVIDENCE_STATUS_APPROVED,
    EVIDENCE_STATUS_LABELS,
    EVIDENCE_STATUS_COLORS,
)


from services.access_scope import (
    can_view_submission as _scope_can_view_submission,
    visible_submissions_query as _scope_visible_submissions_query,
)

from services.evidence_service import (
    get_active_evidence_categories,
    ensure_submissions_for_apprentice,
    summarize_submissions,
    build_evidence_groups,
    global_evidence_stats,
)

from services.pdf_service import (
    get_pdf_page_count,
    render_pdf_page,
)

from core.storage import (
    save_file,
    remove_file,
)

from services.notification_service import (
    notify_apprentice_correction,
    notify_apprentice_approval,
    notify_reviewer_resubmission,
    notify_evidence_comment,
)

from services.permissions import (
    ROLE_APPRENTICE,
    EVIDENCE_MANAGEMENT_ROLES,
    EVIDENCE_APPROVAL_ROLES,
    EVIDENCE_UPLOAD_ROLES,
    has_permission,
)


# =============================================================================
# Blueprint
# =============================================================================

evidences_bp = Blueprint(
    "evidences",
    __name__,
    url_prefix="/evidencias",
)


# =============================================================================
# Constantes de permisos
# =============================================================================

# =============================================================================
# Helpers privados
# =============================================================================

def _now() -> datetime:
    """
    Devuelve la fecha/hora actual en UTC.

    EvidenceSubmission utiliza DateTime(timezone=True), por lo que
    se mantiene un datetime consciente de zona horaria.
    """

    return datetime.now(timezone.utc)


def _visible_submissions():
    apprentice = None
    if current_user.role == ROLE_APPRENTICE:
        apprentice = Apprentice.query.filter_by(
            student_user_id=current_user.id,
        ).first()
        if apprentice is not None:
            created = ensure_submissions_for_apprentice(apprentice)
            if created:
                db.session.commit()
    return _scope_visible_submissions_query(), apprentice


def _load_submission(
    submission_id: int,
) -> EvidenceSubmission:
    """
    Obtiene una entrega junto con sus relaciones principales.
    """

    return (
        EvidenceSubmission.query
        .options(
            joinedload(
                EvidenceSubmission.activity,
            ),
            joinedload(
                EvidenceSubmission.apprentice,
            ),
        )
        .get_or_404(submission_id)
    )


def _check_submission_access(submission: EvidenceSubmission) -> None:
    if not _scope_can_view_submission(submission):
        abort(403)


def _check_management_access() -> None:
    """
    Verifica que el usuario pueda gestionar evidencias.
    """

    if not has_permission("evidences.manage"):
        abort(403)


def _check_approval_access() -> None:
    """
    Verifica que el usuario pueda aprobar evidencias.
    """

    if not has_permission("evidences.approve"):
        abort(403)


def _check_upload_access() -> None:
    """
    Verifica que el usuario pueda cargar evidencias.
    """

    if not has_permission("evidences.upload"):
        abort(403)


def _redirect_to_detail(
    submission: EvidenceSubmission,
):
    """
    Redirección centralizada al detalle de una entrega.
    """

    return redirect(
        url_for(
            "evidences.detail",
            submission_id=submission.id,
        )
    )


def _get_uploaded_file_size(uploaded_file) -> int | None:
    """
    Intenta obtener el tamaño del archivo recibido.

    Si no puede determinarse, devuelve None.
    """

    try:
        stream = uploaded_file.stream

        current_position = stream.tell()

        stream.seek(
            0,
            os.SEEK_END,
        )

        size = stream.tell()

        stream.seek(
            current_position,
        )

        if size < 0:
            return None

        return int(size)

    except Exception:
        return None


# =============================================================================
# Listado principal
# =============================================================================

@evidences_bp.route("/")
@login_required
def index():
    """
    Panel principal de evidencias.

    Permite:

    - filtrar por ficha;
    - filtrar por categoría;
    - filtrar por estado;
    - mostrar estadísticas;
    - mostrar agrupaciones;
    - mostrar información de seguimiento.
    """

    query, selected_apprentice = _visible_submissions()

    group_number = (
        request.args.get(
            "group_number",
            "",
        )
        .strip()
    )

    category_code = (
        request.args.get("category", "").strip()
    )

    status = (
        request.args.get(
            "status",
            "",
        )
        .strip()
    )

    # -------------------------------------------------------------------------
    # Filtros
    # -------------------------------------------------------------------------

    if group_number:

        query = query.filter(
            Apprentice.group_number == group_number,
        )

    if category_code:
        query = (
            query
            .join(EvidenceActivity)
            .filter(EvidenceActivity.category.has(EvidenceCategory.code == category_code))
        )

    if status:

        valid_statuses = {
            EVIDENCE_STATUS_NOT_SUBMITTED,
            EVIDENCE_STATUS_PENDING_REVIEW,
            EVIDENCE_STATUS_REQUIRES_CORRECTION,
            EVIDENCE_STATUS_APPROVED,
        }

        if status not in valid_statuses:
            flash(
                "El estado seleccionado no es válido.",
                "warning",
            )

            status = ""

        else:
            query = query.filter(
                EvidenceSubmission.status == status,
            )

    # -------------------------------------------------------------------------
    # Evidencias
    # -------------------------------------------------------------------------

    submissions = (
        query
        .options(
            joinedload(
                EvidenceSubmission.activity,
            ),
            joinedload(
                EvidenceSubmission.apprentice,
            ),
        )
        .join(EvidenceActivity)
        .join(EvidenceCategory)
        .order_by(
            EvidenceCategory.sort_order,
            EvidenceActivity.sort_order,
            EvidenceActivity.id,
            EvidenceSubmission.created_at,
        )
        .all()
    )

    # -------------------------------------------------------------------------
    # Catálogo de aprendices
    # -------------------------------------------------------------------------

    apprentices = (
        Apprentice.query
        .order_by(
            Apprentice.last_names,
            Apprentice.first_names,
        )
        .all()
    )

    groups = sorted(
        {
            apprentice.group_number
            for apprentice in apprentices
            if apprentice.group_number
        }
    )

    # -------------------------------------------------------------------------
    # Estadísticas
    # -------------------------------------------------------------------------

    evidence_summary = summarize_submissions(
        submissions,
    )

    evidence_groups = build_evidence_groups(
        submissions,
    )

    global_stats = global_evidence_stats(
        query,
    )

    # -------------------------------------------------------------------------
    # Vista
    # -------------------------------------------------------------------------

    return render_template(
        "evidences/index.html",

        submissions=submissions,

        selected_apprentice=selected_apprentice,

        apprentices=apprentices,

        groups=groups,

        evidence_categories=get_active_evidence_categories(),

        status_labels=EVIDENCE_STATUS_LABELS,
        status_colors=EVIDENCE_STATUS_COLORS,
        EVIDENCE_STATUS_NOT_SUBMITTED=EVIDENCE_STATUS_NOT_SUBMITTED,
        EVIDENCE_STATUS_PENDING_REVIEW=EVIDENCE_STATUS_PENDING_REVIEW,
        EVIDENCE_STATUS_REQUIRES_CORRECTION=EVIDENCE_STATUS_REQUIRES_CORRECTION,
        EVIDENCE_STATUS_APPROVED=EVIDENCE_STATUS_APPROVED,

        evidence_summary=evidence_summary,

        evidence_groups=evidence_groups,

        global_stats=global_stats,

        selected_group=group_number,

        selected_category=category_code,

        selected_status=status,
    )


# =============================================================================
# Detalle
# =============================================================================

@evidences_bp.route("/<int:submission_id>")
@login_required
def detail(submission_id: int):
    """
    Visualiza el detalle de una entrega.
    """

    submission = _load_submission(
        submission_id,
    )

    _check_submission_access(
        submission,
    )

    role = str(getattr(current_user, "role", "") or "")
    is_instructor_reviewer = role in {
        "FOLLOW_UP_INSTRUCTOR",
        "LEAD_FOLLOW_UP_INSTRUCTOR",
    }
    is_approval_reviewer = has_permission("evidences.approve")
    can_comment = role == ROLE_APPRENTICE or is_instructor_reviewer
    can_request_correction = (
        is_instructor_reviewer
        and submission.has_file
        and submission.status == EVIDENCE_STATUS_PENDING_REVIEW
    )
    can_approve_evidence = (
        is_approval_reviewer
        and submission.has_file
        and submission.status == EVIDENCE_STATUS_PENDING_REVIEW
    )
    pdf_page_count = 0
    if submission.has_file and submission.is_pdf and submission.file_path and os.path.exists(submission.file_path):
        try:
            pdf_page_count = get_pdf_page_count(submission.file_path)
        except Exception:
            current_app.logger.exception("No se pudo obtener el número de páginas del PDF original.")

    # La UI debe mostrar las acciones de revisión a los revisores aunque una
    # acción esté temporalmente bloqueada por el estado. Esto evita una UI
    # ambigua donde el instructor no sabe por qué no aparecen los controles.
    show_review_controls = (
        is_instructor_reviewer
        or is_approval_reviewer
    )
    review_action_block_reason = None
    if submission.status == EVIDENCE_STATUS_REQUIRES_CORRECTION:
        review_action_block_reason = (
            "La evidencia requiere una nueva entrega antes de poder revisarla nuevamente."
        )
    elif submission.status == EVIDENCE_STATUS_APPROVED:
        review_action_block_reason = "La evidencia ya fue aprobada."
    elif not submission.has_file:
        review_action_block_reason = "Debe existir una entrega antes de iniciar la revisión."

    return render_template(
        "evidences/detail.html",
        submission=submission,
        can_comment=can_comment,
        can_request_correction=can_request_correction,
        can_approve_evidence=can_approve_evidence,
        pdf_page_count=pdf_page_count,
        show_review_controls=show_review_controls,
        review_action_block_reason=review_action_block_reason,
        is_instructor_reviewer=is_instructor_reviewer,
        is_approval_reviewer=is_approval_reviewer,
        EVIDENCE_STATUS_NOT_SUBMITTED=EVIDENCE_STATUS_NOT_SUBMITTED,
        EVIDENCE_STATUS_PENDING_REVIEW=EVIDENCE_STATUS_PENDING_REVIEW,
        EVIDENCE_STATUS_REQUIRES_CORRECTION=EVIDENCE_STATUS_REQUIRES_CORRECTION,
        EVIDENCE_STATUS_APPROVED=EVIDENCE_STATUS_APPROVED,
    )


# =============================================================================
# Cargar / reemplazar evidencia
# =============================================================================

@evidences_bp.route(
    "/<int:submission_id>/upload",
    methods=["POST"],
)
@login_required
def upload(submission_id: int):
    """
    Carga una evidencia.

    Flujo normal:

        no_entregado
            ↓
        pendiente_revision
            ↓
        aprobado

    O:

        pendiente_revision
            ↓
        requiere_correccion
            ↓
        pendiente_revision

    Los roles administrativos autorizados pueden realizar cargas
    administrativas cuando corresponda.
    """

    submission = _load_submission(
        submission_id,
    )

    _check_submission_access(
        submission,
    )

    _check_upload_access()

    # -------------------------------------------------------------------------
    # Validación del estado para aprendiz
    # -------------------------------------------------------------------------

    if has_permission("evidences.upload") and current_user.role == ROLE_APPRENTICE:

        if not submission.can_be_resubmitted:

            flash(
                (
                    "Esta evidencia no puede ser cargada "
                    "en su estado actual."
                ),
                "warning",
            )

            return _redirect_to_detail(
                submission,
            )

    # -------------------------------------------------------------------------
    # Archivo recibido
    # -------------------------------------------------------------------------

    uploaded_file = request.files.get(
        "file",
    )

    if (
        uploaded_file is None
        or not uploaded_file.filename
        or not uploaded_file.filename.strip()
    ):

        flash(
            "Debe seleccionar un archivo.",
            "warning",
        )

        return _redirect_to_detail(
            submission,
        )

    # -------------------------------------------------------------------------
    # Guardar archivo
    # -------------------------------------------------------------------------

    old_file_path = submission.file_path

    try:

        activity = submission.activity
        policy_extensions = activity.allowed_extensions_list if activity else ()
        stored_path, stored_name = save_file(
            uploaded_file,
            subdir=f"evidences/{submission.apprentice_id}",
            allowed_extensions=policy_extensions or None,
            max_size_mb=activity.max_file_size_mb if activity else None,
        )

    except ValueError as exc:

        flash(
            str(exc),
            "warning",
        )

        return _redirect_to_detail(
            submission,
        )

    except Exception:

        current_app.logger.exception(
            "Error almacenando evidencia.",
        )

        flash(
            "No fue posible guardar el archivo.",
            "danger",
        )

        return _redirect_to_detail(
            submission,
        )

    # -------------------------------------------------------------------------
    # Registrar la entrega mediante el modelo
    # -------------------------------------------------------------------------

    try:

        mime_type = getattr(
            uploaded_file,
            "mimetype",
            None,
        )

        file_size_bytes = _get_uploaded_file_size(
            uploaded_file,
        )

        submission.submit(
            file_name=stored_name,
            file_path=stored_path,
            mime_type=mime_type,
            file_size_bytes=file_size_bytes,
            uploaded_at=_now(),
        )

        db.session.commit()

        # Reentrega: avisar al instructor asignado para nueva revisión.
        if submission.is_pending_review:
            notify_reviewer_resubmission(submission)
            db.session.commit()

    except Exception:

        db.session.rollback()

        try:

            remove_file(
                stored_path,
            )

        except Exception:

            current_app.logger.exception(
                "No fue posible eliminar el archivo "
                "tras fallar el registro de la evidencia.",
            )

        current_app.logger.exception(
            "Error registrando entrega de evidencia.",
        )

        flash(
            "No fue posible registrar la evidencia.",
            "danger",
        )

        return _redirect_to_detail(
            submission,
        )

    # -------------------------------------------------------------------------
    # Eliminar archivo anterior
    # -------------------------------------------------------------------------

    if (
        old_file_path
        and old_file_path != stored_path
    ):

        try:

            remove_file(
                old_file_path,
            )

        except Exception:

            current_app.logger.exception(
                "No fue posible eliminar "
                "el archivo anterior de la evidencia.",
            )

    # -------------------------------------------------------------------------
    # Respuesta
    # -------------------------------------------------------------------------

    flash(
        (
            "La evidencia fue cargada correctamente "
            "y quedó pendiente de revisión."
        ),
        "success",
    )

    return _redirect_to_detail(
        submission,
    )


# =============================================================================
# Comentarios / solicitud opcional de corrección
# =============================================================================

@evidences_bp.route(
    "/<int:submission_id>/comments",
    methods=["POST"],
)
@login_required
def add_comment(submission_id: int):
    """Agrega un comentario visible a la conversación de la evidencia.

    El instructor puede marcar opcionalmente el comentario como solicitud de
    corrección; esa marca cambia el estado a ``requiere_correccion`` y activa
    la notificación correspondiente. Los comentarios normales no cambian el
    estado.
    """
    submission = _load_submission(submission_id)
    _check_submission_access(submission)

    # La conversación es una interacción educativa: aprendiz o instructor
    # de seguimiento. Soporte/certificador gestionan el dominio, pero no
    # participan como autores de la conversación de la evidencia.
    is_apprentice = current_user.role == ROLE_APPRENTICE
    is_instructor = current_user.role in {
        "FOLLOW_UP_INSTRUCTOR",
        "LEAD_FOLLOW_UP_INSTRUCTOR",
    }
    if not (is_apprentice or is_instructor):
        abort(403)

    comment_text = request.form.get("comment", "").strip()
    correction_requested = request.form.get("request_correction") == "1"

    if not comment_text:
        flash("El comentario no puede estar vacío.", "warning")
        return _redirect_to_detail(submission)

    if correction_requested:
        if not is_instructor:
            abort(403)
        if not submission.has_file:
            flash("No se puede solicitar corrección sin una entrega.", "warning")
            return _redirect_to_detail(submission)
        if submission.status != EVIDENCE_STATUS_PENDING_REVIEW:
            flash("Solo una evidencia pendiente de revisión puede requerir corrección.", "warning")
            return _redirect_to_detail(submission)

    try:
        submission.add_observation(
            comment_text,
            author_id=current_user.id,
            is_correction_request=correction_requested,
        )

        if correction_requested:
            submission.request_revision(reviewed_by_id=current_user.id)

        db.session.commit()

        if correction_requested:
            notify_apprentice_correction(submission)
        else:
            notify_evidence_comment(submission, author=current_user)
        db.session.commit()

    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error registrando comentario de evidencia.")
        flash("No fue posible registrar el comentario.", "danger")
        return _redirect_to_detail(submission)

    if correction_requested:
        flash("Comentario registrado y se solicitó corrección de la evidencia.", "success")
    else:
        flash("Comentario registrado correctamente.", "success")

    return _redirect_to_detail(submission)


# =============================================================================
# Observaciones / solicitud de corrección
# =============================================================================

@evidences_bp.route(
    "/<int:submission_id>/observe",
    methods=["POST"],
)
@login_required
def observe(submission_id: int):
    """
    Registra una observación y solicita correcciones.

    La entrega pasa a:

        requiere_correccion
    """

    submission = _load_submission(
        submission_id,
    )

    # Autorización y alcance son controles independientes.
    # Nunca permitir que una URL directa permita observar una evidencia ajena.
    _check_submission_access(
        submission,
    )
    _check_management_access()

    if not submission.has_file:

        flash(
            "No se pueden solicitar correcciones "
            "sobre una evidencia que no ha sido entregada.",
            "warning",
        )

        return _redirect_to_detail(
            submission,
        )

    observation = (
        request.form.get(
            "observation",
            "",
        )
        .strip()
    )

    if not observation:

        flash(
            "Debe escribir una observación.",
            "warning",
        )

        return _redirect_to_detail(
            submission,
        )

    # Compatibilidad hacia atrás: observe siempre representa un comentario
    # marcado como solicitud de corrección.
    try:
        submission.add_observation(
            observation,
            author_id=current_user.id,
            is_correction_request=True,
        )
        submission.request_revision(reviewed_by_id=current_user.id)
        db.session.commit()

        notify_apprentice_correction(submission)
        db.session.commit()

    except Exception:

        db.session.rollback()

        current_app.logger.exception(
            "Error registrando solicitud de corrección.",
        )

        flash(
            "No fue posible registrar la solicitud de corrección.",
            "danger",
        )

        return _redirect_to_detail(
            submission,
        )

    flash(
        (
            "La observación fue registrada "
            "y se solicitaron correcciones."
        ),
        "success",
    )

    return _redirect_to_detail(
        submission,
    )


# =============================================================================
# Aprobar evidencia
# =============================================================================

@evidences_bp.route(
    "/<int:submission_id>/approve",
    methods=["POST"],
)
@login_required
def approve(submission_id: int):
    """
    Aprueba una evidencia.

    Pueden aprobar:

    - Instructor de seguimiento
    - Instructor de seguimiento líder
    - Certificador
    - Soporte
    """

    submission = _load_submission(
        submission_id,
    )

    # Un permiso global de aprobación no elimina el alcance del registro.
    _check_submission_access(
        submission,
    )
    _check_approval_access()

    # -------------------------------------------------------------------------
    # Validar archivo
    # -------------------------------------------------------------------------

    if not submission.has_file:

        flash(
            (
                "No se puede aprobar una evidencia "
                "que no ha sido entregada."
            ),
            "warning",
        )

        return _redirect_to_detail(
            submission,
        )

    # -------------------------------------------------------------------------
    # Validar estado
    # -------------------------------------------------------------------------

    if submission.status not in {
        EVIDENCE_STATUS_PENDING_REVIEW,
        EVIDENCE_STATUS_REQUIRES_CORRECTION,
    }:

        flash(
            (
                "La evidencia no se encuentra en un estado "
                "que permita su aprobación."
            ),
            "warning",
        )

        return _redirect_to_detail(
            submission,
        )

    # -------------------------------------------------------------------------
    # Aprobar mediante el modelo
    # -------------------------------------------------------------------------

    try:

        submission.approve(
            approved_by_id=current_user.id,
            approved_at=_now(),
        )

        db.session.commit()

        notify_apprentice_approval(submission)
        db.session.commit()

    except Exception:

        db.session.rollback()

        current_app.logger.exception(
            "Error aprobando evidencia.",
        )

        flash(
            "No fue posible aprobar la evidencia.",
            "danger",
        )

        return _redirect_to_detail(
            submission,
        )

    flash(
        "La evidencia fue aprobada correctamente.",
        "success",
    )

    return _redirect_to_detail(
        submission,
    )


# =============================================================================
# Descargar evidencia original
# =============================================================================

@evidences_bp.route(
    "/<int:submission_id>/download",
)
@login_required
def download(submission_id: int):
    """
    Descarga el archivo original de la evidencia.

    El archivo firmado se mantiene separado.
    """

    submission = _load_submission(
        submission_id,
    )

    _check_submission_access(
        submission,
    )

    if (
        not submission.file_path
        or not os.path.exists(
            submission.file_path,
        )
    ):

        flash(
            "La evidencia no posee un archivo.",
            "warning",
        )

        return _redirect_to_detail(
            submission,
        )

    return send_file(
        submission.file_path,
        as_attachment=True,
        download_name=(
            submission.file_name
            or "evidencia"
        ),
    )


# =============================================================================
# Legacy signed-document endpoint intentionally disabled.
# =============================================================================

@evidences_bp.route(
    "/<int:submission_id>/download-signed",
)
@login_required
def download_signed(submission_id: int):
    # Firma digital fue retirada del producto por decisión funcional.
    abort(404)


# =============================================================================
# Visor PDF integrado: render de página
# =============================================================================

@evidences_bp.route("/<int:submission_id>/pdf-page")
@login_required
def pdf_page(submission_id: int):
    """Renderiza una página PDF como PNG para el visor propio de GIA."""
    submission = _load_submission(submission_id)
    _check_submission_access(submission)

    raw_page = request.args.get("page", "1")
    raw_zoom = request.args.get("zoom", "1.25")

    try:
        page_number = int(raw_page)
        zoom = float(raw_zoom)
    except (TypeError, ValueError):
        abort(400)

    if zoom < 0.75 or zoom > 2.5:
        abort(400)

    file_path = submission.file_path
    if not file_path or not os.path.exists(file_path) or not submission.is_pdf:
        abort(404)

    try:
        png_bytes, _, _ = render_pdf_page(file_path, page_number, zoom=zoom)
    except IndexError:
        abort(404)
    except Exception:
        current_app.logger.exception("No se pudo renderizar la página PDF.")
        abort(500)

    return send_file(
        BytesIO(png_bytes),
        mimetype="image/png",
        as_attachment=False,
        download_name=f"evidencia_{submission.id}_{page_number}.png",
        max_age=0,
    )


# =============================================================================
# Vista previa
# =============================================================================

@evidences_bp.route(
    "/<int:submission_id>/preview",
)
@login_required
def preview(submission_id: int):
    """
    Visualiza el archivo original de una evidencia.
    """

    submission = _load_submission(
        submission_id,
    )

    _check_submission_access(
        submission,
    )

    if (
        not submission.file_path
        or not os.path.exists(
            submission.file_path,
        )
    ):

        flash(
            "La evidencia no posee un archivo.",
            "warning",
        )

        return _redirect_to_detail(
            submission,
        )

    return send_file(
        submission.file_path,
        as_attachment=False,
        download_name=(
            submission.file_name
            or "evidencia"
        ),
    )


# =============================================================================
# Vista previa del PDF firmado
# =============================================================================

@evidences_bp.route(
    "/<int:submission_id>/preview-signed",
)
@login_required
def preview_signed(submission_id: int):
    abort(404)


# =============================================================================
# Firma digital retirada
# =============================================================================

@evidences_bp.route("/<int:submission_id>/sign", methods=["POST"])
@login_required
def sign_submission(submission_id: int):
    abort(404)


# =============================================================================
# Exportaciones
# =============================================================================

__all__ = [
    "evidences_bp",
]