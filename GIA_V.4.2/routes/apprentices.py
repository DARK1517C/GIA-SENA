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
    apprentices = _filtered_apprentices_query().order_by(Apprentice.last_names, Apprentice.first_names).all()
    return render_template(
        "apprentices/index.html",
        apprentices=apprentices,
        total_apprentices=Apprentice.query.count(),
    )


def _filtered_apprentices_query():
    query = Apprentice.query
    search = request.args.get("search", "").strip()
    group_number = request.args.get("group_number", "").strip()
    ep_modality = request.args.get("ep_modality", "").strip()
    status = request.args.get("status", "").strip()

    if search:
        pattern = f"%{search}%"
        query = query.filter(or_(
            Apprentice.first_names.ilike(pattern),
            Apprentice.last_names.ilike(pattern),
            Apprentice.document_number.ilike(pattern),
            Apprentice.email.ilike(pattern),
        ))
    if group_number:
        query = query.filter(Apprentice.group_number.ilike(f"%{group_number}%"))
    if ep_modality:
        query = query.filter(Apprentice.ep_modality.ilike(f"%{ep_modality}%"))
    if status:
        query = query.filter(Apprentice.sofia_status.ilike(f"%{status}%"))
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
    Además registra en logs algunos campos clave para depuración.
    """
    apprentice = Apprentice.query.get_or_404(id)
    try:
        ensure_submissions_for_apprentice(apprentice)
        db.session.commit()
        evidence_submissions = EvidenceSubmission.query.filter_by(apprentice_id=apprentice.id).all()
        evidence_summary = summarize_submissions(evidence_submissions)
    except Exception:
        current_app.logger.debug("No se pudo preparar evidencias del aprendiz", exc_info=True)
        evidence_submissions = []
        evidence_summary = {}

    # Logging para depuración: muestra en consola los campos clave
    try:
        current_app.logger.debug(
            "Apprentice detail: id=%s first=%s last=%s email=%s group=%s",
            id,
            getattr(apprentice, "first_names", None),
            getattr(apprentice, "last_names", None),
            getattr(apprentice, "email", None),
            getattr(apprentice, "group_number", None),
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
        evidence_submissions=evidence_submissions,
        evidence_summary=evidence_summary,
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
                    flash(f"No existe la ficha con número {group_number}. Verifica el número.", "warning")
                    return render_template("apprentices/create.html", form=request.form)
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

        db.session.delete(apprentice)
        db.session.commit()
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
    if getattr(current_user, "role", None) not in ["docente", "super_admin"]:
        flash("No tienes permisos para eliminar aprendices.", "warning")
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
        flash(f"Eliminados {deleted} aprendices.", "success")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error eliminando aprendices en lote")
        flash("No se pudieron eliminar los aprendices seleccionados.", "danger")
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
            apprentice_count, group_count, has_apprentice, has_group = import_reference_workbook(
                file, owner_id=current_user.id, mode="both"
            )
            if not has_apprentice and not has_group:
                flash("El archivo no contiene hojas reconocibles para importar.", "warning")
            else:
                flash(f"Importados/actualizados {apprentice_count} aprendices y {group_count} fichas.", "success")
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
