import os
from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file, current_app
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename

from models import Apprentice, Bitacora
from extensions import db

# helper de almacenamiento
from core.storage import save_file, remove_file

bitacoras_bp = Blueprint("bitacoras", __name__, url_prefix="/bitacoras")


@bitacoras_bp.route("/", methods=["GET", "POST"])
@login_required
def index():
    # GET: listar bitácoras
    if request.method == "GET":
        bitacoras = Bitacora.query.order_by(Bitacora.created_at.desc()).all()
        # adapta el template a templates/bitacoras/index.html
        return render_template("bitacoras/index.html", bitacoras=bitacoras, apprentices=Apprentice.query.all())

    # POST: subir nueva bitácora (ejemplo mínimo)
    apprentice_id = request.form.get("apprentice_id")
    title = request.form.get("title", "").strip() or "Sin título"
    notes = request.form.get("notes", "").strip()
    file = request.files.get("file")
    file_name = None
    file_path = None

    if file and getattr(file, "filename", None):
        try:
            # Guardar usando helper; subdir por aprendiz para organización
            subdir = f"bitacoras/{apprentice_id or 'general'}"
            file_path, stored_name = save_file(file, subdir=subdir)
            file_name = stored_name
        except ValueError as e:
            flash(str(e), "warning")
            return redirect(url_for("bitacoras.index"))
        except Exception:
            current_app.logger.exception("Error guardando archivo de bitácora")
            flash("Error guardando el archivo.", "danger")
            return redirect(url_for("bitacoras.index"))

    bit = Bitacora(
        apprentice_id=apprentice_id,
        uploaded_by_id=current_user.id,
        title=title,
        notes=notes,
        file_name=file_name,
        file_path=file_path,
    )
    db.session.add(bit)
    db.session.commit()
    flash("Bitácora registrada", "success")
    return redirect(url_for("bitacoras.index"))


@bitacoras_bp.route("/<int:entry_id>/download")
@login_required
def download(entry_id):
    entry = Bitacora.query.get_or_404(entry_id)
    # Control de permisos si aplica (ejemplo: solo quien subió o roles específicos)
    if not entry.file_path or not os.path.exists(entry.file_path):
        flash("El archivo no está disponible.", "warning")
        return redirect(url_for("bitacoras.index"))
    return send_file(entry.file_path, as_attachment=True, download_name=entry.file_name)
