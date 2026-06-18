# routes/apprentices.py
from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file, current_app
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
from io import BytesIO
from sqlalchemy import or_
from sqlalchemy.exc import IntegrityError
from models import Apprentice, TrainingGroup, EvidenceSubmission, User
from extensions import db
from services.excel_import import import_reference_workbook
from services.excel_export import export_reference_workbook, export_workbook
from services.utils import parse_form
from services.excel_import import APPRENTICE_MODEL_FIELDS as APPRENTICE_FIELDS
from services.evidence_service import ensure_submissions_for_apprentice, summarize_submissions

apprentices_bp = Blueprint("apprentices", __name__, url_prefix="/apprentices")


@apprentices_bp.route("/")
@login_required
def index():
    """
    Lista de aprendices (index).
    - Aplica filtros via _filtered_apprentices_query (sin paginar por defecto, como antes).
    - Devuelve listas únicas para poblar selects: group_numbers, ep_modalities, sofia_statuses.
    """
    # Obtener aprendices filtrados (igual que antes)
    apprentices = _filtered_apprentices_query().order_by(Apprentice.last_names, Apprentice.first_names).all()

    # Total para toolbar
    try:
        total_apprentices = Apprentice.query.count()
    except Exception:
        total_apprentices = len(apprentices)

    # Obtener opciones únicas para selects (valores existentes en la DB)
    try:
        group_numbers_q = (
            db.session.query(Apprentice.group_number)
            .filter(Apprentice.group_number.isnot(None))
            .distinct()
            .order_by(Apprentice.group_number)
            .all()
        )
        group_numbers = [g[0] for g in group_numbers_q if g[0]]
    except Exception:
        current_app.logger.debug("Error fetching group_numbers", exc_info=True)
        group_numbers = []

    try:
        ep_modalities_q = (
            db.session.query(Apprentice.ep_modality)
            .filter(Apprentice.ep_modality.isnot(None))
            .distinct()
            .order_by(Apprentice.ep_modality)
            .all()
        )
        ep_modalities = [m[0] for m in ep_modalities_q if m[0]]
    except Exception:
        current_app.logger.debug("Error fetching ep_modalities", exc_info=True)
        ep_modalities = []

    try:
        sofia_statuses_q = (
            db.session.query(Apprentice.sofia_status)
            .filter(Apprentice.sofia_status.isnot(None))
            .distinct()
            .order_by(Apprentice.sofia_status)
            .all()
        )
        sofia_statuses = [s[0] for s in sofia_statuses_q if s[0]]
    except Exception:
        current_app.logger.debug("Error fetching sofia_statuses", exc_info=True)
        sofia_statuses = []

    return render_template(
        "apprentices/index.html",
        apprentices=apprentices,
        total_apprentices=total_apprentices,
        group_numbers=group_numbers,
        ep_modalities=ep_modalities,
        sofia_statuses=sofia_statuses,
    )



def _filtered_apprentices_query():
    query = Apprentice.query

    search = (request.args.get("search") or "").strip()
    group_number = (request.args.get("group_number") or "").strip()
    ep_modality = (request.args.get("ep_modality") or "").strip()
    status = (request.args.get("status") or "").strip()  # si prefieres 'sofia_status' cambia aquí y en la plantilla

    # Búsqueda general (varios campos)
    if search:
        if len(search) > 300:
            search = search[:300]
        pattern = f"%{search}%"
        query = query.filter(or_(
            Apprentice.first_names.ilike(pattern),
            Apprentice.last_names.ilike(pattern),
            Apprentice.document_number.ilike(pattern),
            Apprentice.email.ilike(pattern),
        ))

    # group_number: como vendrá de select, usar igualdad exacta para evitar coincidencias parciales
    if group_number:
        query = query.filter(Apprentice.group_number == group_number)

    # ep_modality: igualdad exacta (select)
    if ep_modality:
        query = query.filter(Apprentice.ep_modality == ep_modality)

    # status: igualdad exacta (select)
    if status:
        query = query.filter(Apprentice.sofia_status == status)

    return query


@apprentices_bp.route("/<int:id>")
@login_required
def detail(id):
    """
    Vista detalle del aprendiz.

    Pasa a la plantilla las variables que ésta espera:
    - record: objeto Apprentice
    - page_title: título de la página
    - detail_type: 'apprentice' para que la plantilla seleccione la rama correcta
    - back_url: URL para volver a la lista
    - edit_url: URL de edición si el endpoint existe

    Además prepara las evidencias agrupadas por submission.activity.evidence_type
    y calcula estadísticas (total archivos, aprobados, pendientes) por tipo.
    """
    apprentice = Apprentice.query.get_or_404(id)

    # Preparar evidencias: asegurar que existan submissions y resumir
    try:
        ensure_submissions_for_apprentice(apprentice)
        db.session.commit()
        # obtener todas las submissions relacionadas
        evidence_submissions = EvidenceSubmission.query.filter_by(apprentice_id=apprentice.id).all()
        evidence_summary = summarize_submissions(evidence_submissions)
    except Exception:
        current_app.logger.debug("No se pudo preparar evidencias del aprendiz", exc_info=True)
        evidence_submissions = []
        evidence_summary = {}

    # Agrupar submissions por activity.evidence_type y calcular stats por grupo
    from collections import defaultdict

    grouped = defaultdict(list)
    try:
        for s in evidence_submissions:
            # seguridad: activity puede ser None o no tener evidence_type
            etype = None
            try:
                etype = (getattr(s.activity, "evidence_type", None) or "Sin tipo").strip()
            except Exception:
                etype = "Sin tipo"
            grouped[etype].append(s)
    except Exception:
        current_app.logger.debug("Error agrupando evidence_submissions", exc_info=True)
        grouped = defaultdict(list)

    # Construir estructura final con estadísticas por tipo
    evidence_groups = {}
    try:
        for etype, submissions in grouped.items():
            total_files = 0
            approved = 0
            pending = 0
            # recorrer submissions y sus evidences (si existen)
            for s in submissions:
                evidences = getattr(s, "evidences", []) or []
                total_files += len(evidences)
                for ev in evidences:
                    st = (getattr(ev, "status", "") or "").strip().lower()
                    if st == "aprobado" or st == "approved":
                        approved += 1
                    elif st == "pendiente" or st == "pending":
                        pending += 1
            evidence_groups[etype] = {
                "submissions": submissions,
                "total_files": total_files,
                "approved": approved,
                "pending": pending,
            }
    except Exception:
        current_app.logger.debug("Error calculando estadísticas de evidence_groups", exc_info=True)
        evidence_groups = {}

    # Logging para depuración: muestra en consola los campos clave
    try:
        current_app.logger.debug(
            "Apprentice detail: id=%s first=%s last=%s email=%s group=%s submissions=%s groups=%s",
            id,
            getattr(apprentice, "first_names", None),
            getattr(apprentice, "last_names", None),
            getattr(apprentice, "email", None),
            getattr(apprentice, "group_number", None),
            len(evidence_submissions),
            list(evidence_groups.keys()),
        )
    except Exception:
        current_app.logger.debug("Apprentice detail: could not log apprentice fields", exc_info=True)

    # Construir full_name de forma segura
    full_name = getattr(apprentice, "full_name", None)
    if not full_name:
        fn = getattr(apprentice, "first_names", "") or ""
        ln = getattr(apprentice, "last_names", "") or ""
        full_name = (fn + " " + ln).strip() or f"Aprendiz {apprentice.id}"

    # Determinar edit_url sólo si el endpoint existe
    edit_url = None
    if "apprentices.edit" in current_app.view_functions:
        try:
            edit_url = url_for("apprentices.edit", id=apprentice.id)
        except Exception:
            edit_url = None

    back_url = None
    if "apprentices.index" in current_app.view_functions:
        try:
            back_url = url_for("apprentices.index")
        except Exception:
            back_url = None

    return render_template(
        "apprentices/detail.html",
        record=apprentice,
        page_title=f"Detalle - {full_name}",
        detail_type="apprentice",
        back_url=back_url,
        edit_url=edit_url,
        # variables relacionadas con evidencias
        evidence_submissions=evidence_submissions,  # para compatibilidad si lo necesitas
        evidence_summary=evidence_summary,
        evidence_groups=evidence_groups,  # estructura agrupada y con stats para la plantilla
    )


@apprentices_bp.route("/create", methods=["GET", "POST"])
@login_required
def create():
    """
    Create a new apprentice.
    - GET: render a simple form (templates/apprentices/create.html)
    - POST: validate input, create Apprentice, commit and redirect to index
    """
    if request.method == "POST":
        # Try to parse form using parse_form util if available and APPRENTICE_FIELDS defined
        try:
            form_data = parse_form(request.form, APPRENTICE_FIELDS)
        except Exception:
            # Fallback: read common fields manually
            form_data = {
                "group_number": request.form.get("group_number", "").strip() or None,
                "first_names": request.form.get("first_names", "").strip() or None,
                "last_names": request.form.get("last_names", "").strip() or None,
                "document_type": request.form.get("document_type", "").strip() or None,
                "document_number": request.form.get("document_number", "").strip() or None,
                "email": request.form.get("email", "").strip() or None,
                "phone": request.form.get("phone", "").strip() or None,
                "municipality_origin": request.form.get("municipality_origin", "").strip() or None,
                "program_name": request.form.get("program_name", "").strip() or None,
                "lead_instructor": request.form.get("lead_instructor", "").strip() or None,
                "followup_instructor": request.form.get("followup_instructor", "").strip() or None,
                "ep_modality": request.form.get("ep_modality", "").strip() or None,
                "practice_start_date": request.form.get("practice_start_date", "").strip() or None,
                "practice_end_date": request.form.get("practice_end_date", "").strip() or None,
            }

        # Basic validation: require at least a name
        if not (form_data.get("first_names") or form_data.get("last_names")):
            flash("El nombre del aprendiz es obligatorio.", "warning")
            return render_template("apprentices/create.html", form=request.form)

        # If the model has a non-nullable group_number column, require it here.
        require_group = "group_number" in (APPRENTICE_FIELDS or []) or hasattr(Apprentice, "group_number")
        if require_group and not form_data.get("group_number"):
            flash("El número de ficha (group_number) es obligatorio.", "warning")
            return render_template("apprentices/create.html", form=request.form)

        # Optional: verify that the referenced group exists (if business rule requires)
        group_number = form_data.get("group_number")
        if group_number:
            try:
                group = TrainingGroup.query.filter_by(group_number=group_number).first()
                if group is None:
                    flash(f"No existe la ficha con numero {group_number}. Verifica el numero.", "warning")
                    return render_template("apprentices/create.html", form=request.form)
                form_data["group_id"] = group.id
            except Exception:
                current_app.logger.debug("No se pudo verificar existencia de TrainingGroup", exc_info=True)

        # Create Apprentice instance mapping only known attributes
        apprentice = Apprentice()
        for key, value in form_data.items():
            if hasattr(apprentice, key):
                try:
                    setattr(apprentice, key, value)
                except Exception:
                    current_app.logger.debug("Failed to set attribute %s on Apprentice", key, exc_info=True)

        # Set owner/creator if model has such field
        if hasattr(apprentice, "created_by"):
            try:
                apprentice.created_by = getattr(current_user, "id", None)
            except Exception:
                current_app.logger.debug("Could not set created_by on apprentice", exc_info=True)

        try:
            db.session.add(apprentice)
            db.session.flush()
            from services.excel_import import upsert_student_user
            upsert_student_user(apprentice)
            ensure_submissions_for_apprentice(apprentice)
            db.session.commit()
            flash("Aprendiz creado correctamente.", "success")
            return redirect(url_for("apprentices.index"))
        except IntegrityError:
            db.session.rollback()
            current_app.logger.exception("IntegrityError creando aprendiz")
            flash("Error creando aprendiz: faltan datos obligatorios o conflicto en la base de datos.", "danger")
            return render_template("apprentices/create.html", form=request.form)
        except Exception:
            db.session.rollback()
            current_app.logger.exception("Error creando aprendiz")
            flash("Ocurrió un error al crear el aprendiz. Revisa los datos e intenta de nuevo.", "danger")
            return render_template("apprentices/create.html", form=request.form)

    # GET
    return render_template("apprentices/create.html")


@apprentices_bp.route("/<int:id>/edit", methods=["GET", "POST"])
@login_required
def edit(id):
    """
    Editar aprendiz: GET muestra formulario prellenado; POST actualiza campos permitidos.
    """
    apprentice = Apprentice.query.get_or_404(id)

    if request.method == "POST":
        try:
            form_data = parse_form(request.form, APPRENTICE_FIELDS)
        except Exception:
            form_data = {
                "first_names": request.form.get("first_names", "").strip() or None,
                "last_names": request.form.get("last_names", "").strip() or None,
                "email": request.form.get("email", "").strip() or None,
                "phone": request.form.get("phone", "").strip() or None,
                "program_name": request.form.get("program_name", "").strip() or None,
                "lead_instructor": request.form.get("lead_instructor", "").strip() or None,
                "followup_instructor": request.form.get("followup_instructor", "").strip() or None,
            }

        group_number = form_data.get("group_number")
        if group_number:
            group = TrainingGroup.query.filter_by(group_number=group_number).first()
            if not group:
                flash(f"No existe la ficha con numero {group_number}.", "warning")
                return render_template("apprentices/create.html", form=request.form, editing=True, apprentice=apprentice)
            form_data["group_id"] = group.id

        # Update only attributes that exist on the model
        for key, value in form_data.items():
            if hasattr(apprentice, key):
                try:
                    setattr(apprentice, key, value)
                except Exception:
                    current_app.logger.debug("Failed to set attribute %s on Apprentice during edit", key, exc_info=True)

        try:
            db.session.commit()
            flash("Aprendiz actualizado correctamente.", "success")
            return redirect(url_for("apprentices.detail", id=apprentice.id))
        except IntegrityError:
            db.session.rollback()
            current_app.logger.exception("IntegrityError actualizando aprendiz")
            flash("Error al actualizar aprendiz: datos inválidos.", "danger")
            return render_template("apprentices/create.html", form=request.form, editing=True, apprentice=apprentice)
        except Exception:
            db.session.rollback()
            current_app.logger.exception("Error actualizando aprendiz")
            flash("Ocurrió un error al actualizar el aprendiz.", "danger")
            return render_template("apprentices/create.html", form=request.form, editing=True, apprentice=apprentice)

    # GET -> render form prefilled
    return render_template("apprentices/create.html", editing=True, apprentice=apprentice)


@apprentices_bp.route("/<int:id>/delete", methods=["POST"])
@login_required
def delete(id):
    """
    Elimina un aprendiz. Maneja errores de integridad y hace rollback si falla.
    """
    apprentice = Apprentice.query.get_or_404(id)
    try:
        if apprentice.student_user_id:
            student_user = User.query.get(apprentice.student_user_id)
            if student_user and student_user.role == "aprendiz":
                other_apprentice_references = Apprentice.query.filter(
                    Apprentice.student_user_id == student_user.id,
                    Apprentice.id != apprentice.id,
                ).count()
                if other_apprentice_references == 0:
                    db.session.delete(student_user)

        document_number = apprentice.document_number
        db.session.delete(apprentice)
        db.session.commit()
        current_app.logger.info("Aprendiz eliminado: id=%s documento=%s usuario=%s", id, document_number, current_user.id)
        flash("Aprendiz eliminado correctamente.", "success")
    except IntegrityError:
        db.session.rollback()
        current_app.logger.exception("IntegrityError eliminando aprendiz")
        flash("No se pudo eliminar el aprendiz por restricciones en la base de datos.", "danger")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error eliminando aprendiz")
        flash("Ocurrió un error al eliminar el aprendiz.", "danger")
    return redirect(url_for("apprentices.index"))


@apprentices_bp.route("/bulk-delete", methods=["POST"])
@login_required
def bulk_delete():
    if getattr(current_user, "role", None) != "super_admin":
        flash("No tienes permisos para eliminar aprendices en lote.", "warning")
        return redirect(url_for("apprentices.index"))

    try:
        ids = [int(item) for item in request.form.getlist("selected_ids") if str(item).strip()]
    except Exception:
        flash("Seleccion invalida.", "warning")
        return redirect(url_for("apprentices.index"))

    if not ids:
        flash("Selecciona al menos un aprendiz para eliminar.", "warning")
        return redirect(url_for("apprentices.index"))

    try:
        items = Apprentice.query.filter(Apprentice.id.in_(ids)).all()
        deleted = len(items)
        for item in items:
            db.session.delete(item)
        db.session.commit()
        current_app.logger.info("Aprendices eliminados en lote: total=%s usuario=%s", deleted, current_user.id)
        flash(f"Eliminados {deleted} aprendices.", "success")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error eliminando aprendices en lote")
        flash("No se pudieron eliminar los aprendices seleccionados.", "danger")
    return redirect(url_for("apprentices.index"))


@apprentices_bp.route("/delete-all", methods=["POST"])
@login_required
def delete_all():
    if getattr(current_user, "role", None) != "super_admin":
        flash("No tienes permisos para eliminar todos los aprendices.", "warning")
        return redirect(url_for("apprentices.index"))
    try:
        items = Apprentice.query.all()
        total = len(items)
        for item in items:
            db.session.delete(item)
        db.session.commit()
        current_app.logger.info("Todos los aprendices eliminados: total=%s usuario=%s", total, current_user.id)
        flash(f"Eliminados todos los aprendices ({total}).", "success")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error eliminando todos los aprendices")
        flash("No se pudieron eliminar todos los aprendices.", "danger")
    return redirect(url_for("apprentices.index"))


@apprentices_bp.route("/import", methods=["GET", "POST"])
@login_required
def import_excel():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            flash("No se subió ningún archivo", "warning")
            return redirect(url_for("apprentices.import_excel"))
        filename = secure_filename(file.filename)
        try:
            result = import_reference_workbook(file, owner_id=current_user.id, mode="both")
            if not result.has_apprentice_sheet and not result.has_group_sheet:
                flash("El archivo no contiene las hojas oficiales Record Fichas y Aprendices.", "warning")
            else:
                flash(
                    f"Importacion completada: {result.apprentice_count} aprendices y {result.group_count} fichas. "
                    f"Omitidos: {result.skipped_apprentices}.",
                    "success" if not result.errors else "warning",
                )
                for message in result.errors[:5]:
                    flash(message, "warning")
        except Exception:
            current_app.logger.exception("Error importando aprendices desde Excel")
            flash("Error al procesar el archivo. Verifica el formato y vuelve a intentarlo.", "danger")
        return redirect(url_for("apprentices.index"))
    return render_template("apprentices/import.html")


@apprentices_bp.route("/export")
@login_required
def export_all():
    try:
        apprentices = _filtered_apprentices_query().order_by(Apprentice.last_names, Apprentice.first_names).all()
        group_numbers = {item.group_number for item in apprentices if item.group_number}
        groups = TrainingGroup.query.filter(TrainingGroup.group_number.in_(group_numbers)).all() if group_numbers else []
        output = export_reference_workbook(apprentices, groups)

        # Ensure output is file-like
        if isinstance(output, bytes):
            output = BytesIO(output)
            output.seek(0)
        elif hasattr(output, "seek"):
            try:
                output.seek(0)
            except Exception:
                pass

        return send_file(
            output,
            as_attachment=True,
            download_name="referencias_aprendices.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception:
        current_app.logger.exception("Error exportando aprendices")
        flash("Error al generar el archivo de exportación. Intenta de nuevo.", "danger")
        return redirect(url_for("apprentices.index"))
