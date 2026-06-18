# routes/bitacoras.py
from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required, current_user
from models import Apprentice, Bitacora
from extensions import db

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
    if file:
        file_name = file.filename
        # guarda en uploads y registra path (ajusta según tu lógica)
        upload_dir = current_app.config.get("UPLOAD_DIR")
        file_path = os.path.join(upload_dir, secure_filename(file_name))
        file.save(file_path)
    bit = Bitacora(apprentice_id=apprentice_id, uploaded_by_id=current_user.id, title=title, notes=notes, file_name=file_name, file_path=file_path)
    db.session.add(bit)
    db.session.commit()
    flash("Bitácora registrada", "success")
    return redirect(url_for("bitacoras.index"))
