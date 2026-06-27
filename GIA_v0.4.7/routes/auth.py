from flask import Blueprint, render_template, redirect, url_for, request, flash, current_app
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash
from models import User, Apprentice
from extensions import db
from urllib.parse import urlparse, urljoin

auth_bp = Blueprint("auth", __name__, url_prefix="/auth")

def is_safe_url(target):
    """
    Evita redirecciones abiertas. Devuelve True si 'target' es una URL segura
    dentro del mismo host.
    """
    if not target:
        return False
    ref_url = urlparse(request.host_url)
    test_url = urlparse(urljoin(request.host_url, target))
    return (test_url.scheme in ("http", "https")) and (ref_url.netloc == test_url.netloc)


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)

            # Priorizar 'next' si es seguro
            next_url = request.args.get("next") or request.form.get("next")
            if next_url and is_safe_url(next_url):
                return redirect(next_url)

            # Redirección por rol: aprendiz -> evidences, otros -> dashboard
            if getattr(user, "role", None) == "aprendiz":
                return redirect(url_for("evidences.index"))
            else:
                return redirect(url_for("dashboard.index"))

        flash("Usuario o contraseña incorrectos", "danger")
    return render_template("auth/login.html")

@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("auth.login"))

@auth_bp.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    """
    Vista de perfil unificada y robusta:
    - Maneja cambio de contraseña.
    - Actualiza campos de User y Apprentice según rol.
    - Comprueba existencia de atributos en el modelo antes de asignar.
    """
    # Cargar datos relacionados (si aplica)
    apprentice = None
    if current_user.role == "aprendiz":
        apprentice = Apprentice.query.filter_by(student_user_id=current_user.id).first()

    # Construir back_url/back_label (misma lógica previa)
    role = getattr(current_user, "role", None)

    def _safe_url(endpoint, fallback=None, **values):
        try:
            if endpoint and endpoint in current_app.view_functions:
                return url_for(endpoint, **values)
        except Exception:
            current_app.logger.debug("safe_url failed for %s", endpoint, exc_info=True)
        return fallback

    if role == "aprendiz":
        back_url = _safe_url("evidences.index",
                             fallback=_safe_url("apprentices.index", fallback=url_for("auth.profile")))
        back_label = "‹ Volver a Evidencias"
    elif role in ("docente", "super_admin", "visualizador", "administrativo"):
        back_url = _safe_url("dashboard.index", fallback=url_for("auth.profile"))
        back_label = "‹ Volver a Estadisticas"
    else:
        back_url = _safe_url("apprentices.index", fallback=url_for("auth.profile"))
        back_label = "‹ Volver"

    # POST: procesar formularios
    if request.method == "POST":
        # Registrar datos recibidos para depuración
        current_app.logger.debug("Profile POST data: %s", dict(request.form))

        # 1) Cambio de contraseña
        new_password = request.form.get("new_password", "").strip()
        if new_password:
            confirm_password = request.form.get("confirm_password", "").strip()
            current_password = request.form.get("current_password", "").strip()
            if not current_user.check_password(current_password):
                flash("La contraseña actual no es correcta.", "warning")
                return redirect(url_for("auth.profile"))
            if new_password != confirm_password:
                flash("La confirmación de contraseña no coincide.", "warning")
                return redirect(url_for("auth.profile"))
            try:
                current_user.password_hash = generate_password_hash(new_password)
                db.session.add(current_user)
                db.session.commit()
                flash("Contraseña actualizada.", "success")
            except Exception:
                db.session.rollback()
                current_app.logger.exception("Error actualizando contraseña de usuario %s", current_user.id)
                flash("Ocurrió un error al actualizar la contraseña.", "danger")
            return redirect(url_for("auth.profile"))

        # 2) Edición de perfil (un único formulario)
        form_full_name = request.form.get("full_name")
        form_email = request.form.get("email")
        form_phone = request.form.get("phone")
        form_role = request.form.get("role")
        form_document_type = request.form.get("document_type")
        form_document_number = request.form.get("document_number")

        try:
            # SUPER_ADMIN: full_name y role
            if current_user.role == "super_admin":
                if form_full_name is not None:
                    current_user.full_name = form_full_name.strip() or current_user.full_name

                if form_role:
                    # validar role_labels si existe en contexto
                    try:
                        labels = current_app.jinja_env.globals.get("role_labels", None)
                    except Exception:
                        labels = None
                    if labels and isinstance(labels, dict) and form_role not in labels:
                        flash("Rol inválido.", "warning")
                        return redirect(url_for("auth.profile"))
                    if form_role and form_role != current_user.role:
                        current_user.role = form_role

            # DOCENTE / ADMINISTRATIVO / SUPER_ADMIN: email y phone (si User tiene esos atributos)
            if current_user.role in ("docente", "administrativo", "super_admin"):
                if form_email is not None and hasattr(current_user, "email"):
                    current_user.email = form_email.strip() or current_user.email

                if form_phone is not None:
                    # si User tiene phone lo actualizamos; si no, intentar en Apprentice
                    if hasattr(current_user, "phone"):
                        setattr(current_user, "phone", form_phone.strip() or getattr(current_user, "phone", None))
                    else:
                        if apprentice:
                            apprentice.phone = form_phone.strip() or apprentice.phone

            # APRENDIZ: actualizar Apprentice y sincronizar User cuando proceda
            if current_user.role == "aprendiz":
                # email en User si existe
                if form_email is not None and hasattr(current_user, "email"):
                    current_user.email = form_email.strip() or current_user.email

                # teléfono preferentemente en Apprentice
                if form_phone is not None and apprentice:
                    apprentice.phone = form_phone.strip() or apprentice.phone

                if apprentice:
                    # document_number: validar unicidad
                    if form_document_number:
                        new_doc = form_document_number.strip()
                        if new_doc and new_doc != (apprentice.document_number or ""):
                            other = Apprentice.query.filter_by(document_number=new_doc).first()
                            if other and other.id != apprentice.id:
                                flash("El número de documento ya está en uso por otro aprendiz.", "warning")
                                return redirect(url_for("auth.profile"))
                            apprentice.document_number = new_doc

                    if form_document_type is not None:
                        apprentice.document_type = form_document_type.strip() or apprentice.document_type

                    # sincronizar username con document_number si procede
                    if apprentice.document_number:
                        if current_user.username != apprentice.document_number:
                            existing_user = User.query.filter_by(username=apprentice.document_number).first()
                            if existing_user and existing_user.id != current_user.id:
                                flash("No se puede sincronizar el número de documento con el usuario: ya existe otro usuario con ese username.", "warning")
                                return redirect(url_for("auth.profile"))
                            current_user.username = apprentice.document_number

            # Asegurar que los objetos están en la sesión antes de commit
            db.session.add(current_user)
            if apprentice:
                db.session.add(apprentice)

            db.session.commit()
            flash("Perfil actualizado correctamente.", "success")
            return redirect(url_for("auth.profile"))

        except Exception:
            db.session.rollback()
            current_app.logger.exception("Error actualizando perfil para usuario %s", current_user.id)
            flash("Ocurrió un error al actualizar el perfil. Intenta de nuevo.", "danger")
            return redirect(url_for("auth.profile"))

    # GET: renderizar plantilla
    return render_template("auth/profile.html", apprentice=apprentice, back_url=back_url, back_label=back_label)