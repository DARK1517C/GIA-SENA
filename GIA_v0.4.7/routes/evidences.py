# routes/evidences.py
import os
import io
import tempfile
from datetime import datetime
from flask import (
    Blueprint,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
    jsonify,
    abort,
)
from flask_login import current_user, login_required
from werkzeug.utils import secure_filename
from sqlalchemy.orm import joinedload

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
    # NOTE: SignatureAudit intentionally NOT imported here to avoid import-time failures.
)
from services.evidence_service import (
    ensure_submissions_for_apprentice,
    seed_default_evidences_for_group,
    summarize_submissions,
    build_evidence_groups,
)

# Helper de almacenamiento centralizado
from core.storage import save_file, remove_file

# Definición del blueprint
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

    submissions = (
        query
        .options(joinedload(EvidenceSubmission.activity), joinedload(EvidenceSubmission.apprentice))
        .join(EvidenceActivity)
        .order_by(EvidenceActivity.id, EvidenceSubmission.created_at)
        .all()
    )

    apprentices = Apprentice.query.order_by(Apprentice.last_names, Apprentice.first_names).all()
    groups = sorted({apprentice.group_number for apprentice in apprentices if apprentice.group_number})

    evidence_summary = summarize_submissions(submissions)
    evidence_groups = build_evidence_groups(submissions)

    return render_template(
        "evidences/index.html",
        submissions=submissions,
        selected_apprentice=selected_apprentice,
        apprentices=apprentices,
        groups=groups,
        evidence_types=EVIDENCE_TYPES,
        status_labels=EVIDENCE_STATUS_LABELS,
        status_colors=EVIDENCE_STATUS_COLORS,
        evidence_summary=evidence_summary,
        evidence_groups=evidence_groups,
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
    # Servir inline para que PDF.js o iframe puedan mostrarlo
    return send_file(submission.file_path, as_attachment=False, download_name=submission.file_name)


@evidences_bp.route("/sign", methods=["POST"])
@login_required
def sign():
    """
    Endpoint para recibir la imagen de firma (dibujo o imagen subida) y aplicarla
    sobre la página indicada del PDF. Guarda una copia firmada y registra un audit trail.
    Parámetros esperados (multipart/form-data):
      - submission_id
      - page (int)
      - pos_x (float) porcentaje 0-100
      - pos_y (float) porcentaje 0-100
      - scale (float) factor de escala relativo
      - signature_image (file) imagen PNG/JPEG con la firma
    """
    # Validaciones básicas
    submission_id = request.form.get("submission_id")
    if not submission_id:
        return "submission_id requerido", 400

    try:
        submission = EvidenceSubmission.query.get_or_404(int(submission_id))
    except Exception:
        return "Submission no encontrado", 404

    # Permisos: solo aprendiz (propio), docente o super_admin pueden firmar según política
    if current_user.role == "aprendiz" and submission.apprentice.student_user_id != current_user.id:
        return "No tienes permisos para firmar esta evidencia", 403

    # Recibir metadatos
    try:
        page = int(request.form.get("page", 1))
    except Exception:
        page = 1
    try:
        pos_x = float(request.form.get("pos_x", 50.0))
        pos_y = float(request.form.get("pos_y", 50.0))
        scale = float(request.form.get("scale", 0.6))
    except Exception:
        pos_x, pos_y, scale = 50.0, 50.0, 0.6

    sig_file = request.files.get("signature_image")
    if not sig_file:
        return "signature_image requerido", 400

    if not submission.file_path or not os.path.exists(submission.file_path):
        return "PDF no disponible", 404

    # Import dinámico de librerías pesadas y del modelo de auditoría
    try:
        from PIL import Image
        from reportlab.pdfgen import canvas
        from reportlab.lib.utils import ImageReader
        from PyPDF2 import PdfReader, PdfWriter
    except Exception:
        current_app.logger.exception("Dependencias para firmar no disponibles")
        return "Dependencias para procesar la firma no están instaladas en el servidor.", 500

    try:
        from models import SignatureAudit  # opcional: puede no existir en tu esquema
    except Exception:
        SignatureAudit = None
        current_app.logger.debug("Modelo SignatureAudit no encontrado; la auditoría de firma no se guardará.", exc_info=True)

    # Procesar la imagen de firma y superponerla en la página indicada
    try:
        # Abrir la imagen de firma con Pillow
        sig_img = Image.open(sig_file.stream).convert("RGBA")

        # Leer PDF original
        reader = PdfReader(submission.file_path)
        if page < 1 or page > len(reader.pages):
            return "Página inválida", 400

        # Obtener tamaño de la página en puntos (PyPDF2 usa mediaBox)
        target_page = reader.pages[page - 1]
        media = target_page.mediabox
        page_width = float(media.width)
        page_height = float(media.height)

        # Calcular tamaño de la firma en puntos
        sig_target_w = sig_img.width * scale
        sig_target_h = sig_img.height * scale

        # Convertir posición porcentual a coordenadas en puntos (origin bottom-left)
        x_pt = (pos_x / 100.0) * page_width
        y_pt = (pos_y / 100.0) * page_height

        # Crear un PDF temporal con la imagen en la posición indicada
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(page_width, page_height))
        img_reader = ImageReader(sig_img)

        # Ajuste si la imagen es demasiado grande
        if sig_target_w > page_width * 0.9:
            factor = (page_width * 0.9) / sig_target_w
            sig_target_w *= factor
            sig_target_h *= factor

        c.drawImage(img_reader, x_pt, y_pt, width=sig_target_w, height=sig_target_h, mask='auto')
        c.save()
        packet.seek(0)

        # Leer overlay y combinar páginas
        overlay_pdf = PdfReader(packet)
        writer = PdfWriter()

        for i, p in enumerate(reader.pages):
            base = p
            if i == page - 1:
                overlay_page = overlay_pdf.pages[0]
                try:
                    base.merge_page(overlay_page)
                except Exception:
                    # fallback: intentar merge de nuevo y registrar
                    current_app.logger.debug("merge_page falló, intentando fallback", exc_info=True)
                    base.merge_page(overlay_page)
            writer.add_page(base)

        # Guardar PDF firmado en un nuevo archivo (no sobrescribir original)
        timestamp = int(datetime.utcnow().timestamp())
        orig_dir = os.path.dirname(submission.file_path)
        orig_name = os.path.splitext(submission.file_name)[0]
        signed_filename = f"{orig_name}.signed.{timestamp}.pdf"
        signed_path = os.path.join(orig_dir, signed_filename)

        with open(signed_path, "wb") as out_f:
            writer.write(out_f)

        # Actualizar submission (guardar copia firmada; conservar original para auditoría)
        submission.file_path = signed_path
        submission.file_name = signed_filename
        submission.status = EVIDENCE_STATUS_PENDING
        submission.uploaded_at = datetime.utcnow()
        db.session.add(submission)

        # Guardar registro de auditoría si el modelo existe
        if SignatureAudit is not None:
            try:
                audit = SignatureAudit(
                    submission_id=submission.id,
                    user_id=current_user.id,
                    created_at=datetime.utcnow(),
                    ip=request.remote_addr,
                    user_agent=request.headers.get("User-Agent"),
                    page=page,
                    pos_x=pos_x,
                    pos_y=pos_y,
                    scale=scale,
                    signature_file_name=secure_filename(sig_file.filename or "signature.png"),
                )
                db.session.add(audit)
            except Exception:
                current_app.logger.exception("No se pudo crear registro SignatureAudit")

        db.session.commit()

        return jsonify({"ok": True, "file": submission.file_name})

    except Exception as e:
        current_app.logger.exception("Error aplicando firma al PDF")
        db.session.rollback()
        return str(e), 500
