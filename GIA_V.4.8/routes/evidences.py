import os
from datetime import datetime
from flask import Blueprint, current_app, flash, redirect, render_template, request, send_file, url_for
from flask_login import current_user, login_required
from werkzeug.utils import secure_filename

from extensions import db
from models import (
    Apprentice,
    EvidenceActivity,
    EvidenceSubmission,
    InstructorSignature,
    EVIDENCE_STATUS_APPROVED,
    EVIDENCE_STATUS_COLORS,
    EVIDENCE_STATUS_LABELS,
    EVIDENCE_STATUS_NOT_SUBMITTED,
    EVIDENCE_STATUS_PENDING,
    EVIDENCE_TYPES,
)
from services.evidence_service import ensure_submissions_for_apprentice, seed_default_evidences_for_group, summarize_submissions

# Nuevo helper de almacenamiento
from core.storage import save_file, remove_file

evidences_bp = Blueprint("evidences", __name__, url_prefix="/evidences")


def _upload_dir(*parts):
    path = os.path.join(current_app.config["UPLOAD_DIR"], "evidences", *parts)
    os.makedirs(path, exist_ok=True)
    return path


def _visible_submissions():
    if current_user.role == "aprendiz":
        apprentice = Apprentice.query.filter_by(student_user_id=current_user.id).first()
        if apprentice:
            ensure_submissions_for_apprentice(apprentice)
            db.session.commit()
            return EvidenceSubmission.query.filter_by(apprentice_id=apprentice.id), apprentice
        return EvidenceSubmission.query.filter(False), None
    return EvidenceSubmission.query.join(Apprentice), None


@evidences_bp.route("/")
@login_required
def index():
    query, selected_apprentice = _visible_submissions()
    group_number = request.args.get("group_number", "").strip()
    evidence_type = request.args.get("evidence_type", "").strip()
    status = request.args.get("status", "").strip()

    if group_number:
        query = query.filter(Apprentice.group_number == group_number)
    if evidence_type:
        query = query.join(EvidenceActivity).filter(EvidenceActivity.evidence_type == evidence_type)
    if status:
        query = query.filter(EvidenceSubmission.status == status)

    submissions = query.order_by(EvidenceSubmission.status, EvidenceSubmission.id).all()
    apprentices = Apprentice.query.order_by(Apprentice.last_names, Apprentice.first_names).all()
    groups = sorted({apprentice.group_number for apprentice in apprentices if apprentice.group_number})

    return render_template(
        "evidences/index.html",
        submissions=submissions,
        selected_apprentice=selected_apprentice,
        apprentices=apprentices,
        groups=groups,
        evidence_types=EVIDENCE_TYPES,
        status_labels=EVIDENCE_STATUS_LABELS,
        status_colors=EVIDENCE_STATUS_COLORS,
        summary=summarize_submissions(submissions),
    )


@evidences_bp.route("/<int:submission_id>")
@login_required
def detail(submission_id):
    submission = EvidenceSubmission.query.get_or_404(submission_id)
    if current_user.role == "aprendiz" and submission.apprentice.student_user_id != current_user.id:
        flash("No tienes permisos para ver esta evidencia.", "warning")
        return redirect(url_for("evidences.index"))
    return render_template(
        "evidences/detail.html",
        submission=submission,
        status_labels=EVIDENCE_STATUS_LABELS,
        status_colors=EVIDENCE_STATUS_COLORS,
        signature=InstructorSignature.query.filter_by(user_id=current_user.id).first(),
    )


@evidences_bp.route("/<int:submission_id>/upload", methods=["POST"])
@login_required
def upload(submission_id):
    submission = EvidenceSubmission.query.get_or_404(submission_id)
    if current_user.role == "aprendiz" and submission.apprentice.student_user_id != current_user.id:
        flash("No puedes cargar evidencias para otro aprendiz.", "warning")
        return redirect(url_for("evidences.index"))
    if current_user.role not in ["aprendiz", "docente", "super_admin"]:
        flash("No tienes permisos para cargar esta evidencia.", "warning")
        return redirect(url_for("evidences.detail", submission_id=submission.id))

    file = request.files.get("file")
    if not file or not file.filename:
        flash("Selecciona un archivo para cargar.", "warning")
        return redirect(url_for("evidences.detail", submission_id=submission.id))

    # Guardar archivo usando el helper centralizado
    try:
        path, stored_name = save_file(file, subdir=f"evidences/{submission.apprentice_id}")
    except ValueError as e:
        flash(str(e), "warning")
        return redirect(url_for("evidences.detail", submission_id=submission.id))
    except Exception:
        current_app.logger.exception("Error guardando archivo de evidencia")
        flash("Error guardando el archivo.", "danger")
        return redirect(url_for("evidences.detail", submission_id=submission.id))

    # Intentar eliminar archivo previo si existía
    if getattr(submission, "file_path", None):
        try:
            remove_file(submission.file_path)
        except Exception:
            current_app.logger.debug("No se pudo eliminar archivo previo", exc_info=True)

    # Actualizar registro
    submission.file_name = stored_name
    submission.file_path = path
    submission.status = EVIDENCE_STATUS_PENDING
    submission.uploaded_at = datetime.utcnow()
    db.session.commit()
    flash("Evidencia cargada y enviada a revision.", "success")
    return redirect(url_for("evidences.detail", submission_id=submission.id))


@evidences_bp.route("/<int:submission_id>/observe", methods=["POST"])
@login_required
def observe(submission_id):
    if current_user.role not in ["docente", "super_admin"]:
        flash("Solo Instructor o Administrador puede registrar observaciones.", "warning")
        return redirect(url_for("evidences.index"))
    submission = EvidenceSubmission.query.get_or_404(submission_id)
    submission.observations = request.form.get("observations", "").strip()
    db.session.commit()
    flash("Observaciones actualizadas.", "success")
    return redirect(url_for("evidences.detail", submission_id=submission.id))


@evidences_bp.route("/<int:submission_id>/approve", methods=["POST"])
@login_required
def approve(submission_id):
    if current_user.role not in ["docente", "super_admin"]:
        flash("Solo Instructor o Administrador puede aprobar evidencias.", "warning")
        return redirect(url_for("evidences.index"))

    submission = EvidenceSubmission.query.get_or_404(submission_id)
    signature_file = request.files.get("signature")
    signature = InstructorSignature.query.filter_by(user_id=current_user.id).first()

    if signature_file and signature_file.filename:
        # Guardar la firma usando helper
        try:
            path, stored_name = save_file(signature_file, subdir=f"signatures/{current_user.id}")
        except ValueError as e:
            flash(str(e), "warning")
            return redirect(url_for("evidences.detail", submission_id=submission.id))
        except Exception:
            current_app.logger.exception("Error guardando archivo de firma")
            flash("Error guardando la firma.", "danger")
            return redirect(url_for("evidences.detail", submission_id=submission.id))

        # Eliminar firma previa si existía
        if signature and getattr(signature, "file_path", None):
            try:
                remove_file(signature.file_path)
            except Exception:
                current_app.logger.debug("No se pudo eliminar firma previa", exc_info=True)

        if signature is None:
            signature = InstructorSignature(user_id=current_user.id, file_name=stored_name, file_path=path)
            db.session.add(signature)
        else:
            signature.file_name = stored_name
            signature.file_path = path

    if signature is None:
        flash("Carga una firma para aprobar esta evidencia.", "warning")
        return redirect(url_for("evidences.detail", submission_id=submission.id))

    submission.status = EVIDENCE_STATUS_APPROVED
    submission.approved_at = datetime.utcnow()
    submission.approved_by_id = current_user.id
    submission.signature_file_name = signature.file_name
    submission.signature_file_path = signature.file_path
    db.session.commit()
    flash("Evidencia aprobada y firmada.", "success")
    return redirect(url_for("evidences.detail", submission_id=submission.id))


@evidences_bp.route("/<int:submission_id>/download")
@login_required
def download(submission_id):
    submission = EvidenceSubmission.query.get_or_404(submission_id)
    if current_user.role == "aprendiz" and submission.apprentice.student_user_id != current_user.id:
        flash("No tienes permisos para descargar este archivo.", "warning")
        return redirect(url_for("evidences.index"))
    if not submission.file_path or not os.path.exists(submission.file_path):
        flash("La evidencia no tiene archivo cargado.", "warning")
        return redirect(url_for("evidences.detail", submission_id=submission.id))
    return send_file(submission.file_path, as_attachment=True, download_name=submission.file_name)


@evidences_bp.route("/<int:submission_id>/preview")
@login_required
def preview(submission_id):
    submission = EvidenceSubmission.query.get_or_404(submission_id)
    if current_user.role == "aprendiz" and submission.apprentice.student_user_id != current_user.id:
        flash("No tienes permisos para ver este archivo.", "warning")
        return redirect(url_for("evidences.index"))
    if not submission.file_path or not os.path.exists(submission.file_path):
        flash("La evidencia no tiene archivo cargado.", "warning")
        return redirect(url_for("evidences.detail", submission_id=submission.id))
    return send_file(submission.file_path, as_attachment=False, download_name=submission.file_name)
