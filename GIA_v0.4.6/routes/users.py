# routes/users.py
from flask import Blueprint, render_template, request, redirect, url_for, flash, current_app
from flask_login import login_required, current_user
from werkzeug.security import generate_password_hash
from sqlalchemy import or_
from sqlalchemy.exc import IntegrityError

from models import Apprentice, User
from extensions import db

users_bp = Blueprint("users", __name__, url_prefix="/users")


def _is_super_admin():
    return getattr(current_user, "role", None) == "super_admin"


@users_bp.route("/", methods=["GET"])
@login_required
def index():
    """
    Listado de usuarios. Solo accesible para roles con permiso (super_admin).
    """
    if not _is_super_admin():
        flash("No tienes permisos para ver la lista de usuarios.", "warning")
        return redirect(url_for("dashboard.index"))

    search = request.args.get("search", "").strip()
    role_filter = request.args.get("role", "").strip()
    query = User.query

    if search:
        pattern = f"%{search}%"
        query = query.filter(or_(
            User.username.ilike(pattern),
            User.full_name.ilike(pattern),
            User.email.ilike(pattern),
        ))

    if role_filter:
        query = query.filter(User.role == role_filter)

    users = query.order_by(User.id).all()

    # Etiquetas legibles para roles (puedes ajustar los textos)
    role_labels = {
        "super_admin": "Super Admin",
        "docente": "Instructor",
        "visualizador": "Administrativo",
        "aprendiz": "Aprendiz",
    }

    return render_template("users/index.html", users=users, role_labels=role_labels)


@users_bp.route("/create", methods=["GET", "POST"])
@login_required
def create():
    """
    Crear un nuevo usuario. Solo super_admin.
    """
    if not _is_super_admin():
        flash("No tienes permisos para crear usuarios.", "warning")
        return redirect(url_for("users.index"))

    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        full_name = (request.form.get("full_name") or "").strip()
        email = (request.form.get("email") or "").strip()
        role = (request.form.get("role") or "visualizador").strip()
        password = (request.form.get("password") or "").strip()

        if not username or not full_name or not password:
            flash("Usuario, nombre completo y contraseña son obligatorios.", "danger")
            return render_template("users/form.html", user=None)

        if User.query.filter_by(username=username).first():
            flash("El nombre de usuario ya existe.", "danger")
            return render_template("users/form.html", user=None)

        hashed = generate_password_hash(password)
        user = User(username=username, full_name=full_name, email=email, role=role, password_hash=hashed)
        try:
            db.session.add(user)
            db.session.commit()
            flash("Usuario creado correctamente.", "success")
            return redirect(url_for("users.index"))
        except IntegrityError:
            db.session.rollback()
            current_app.logger.exception("IntegrityError creando usuario")
            flash("Error al crear el usuario. Verifica los datos.", "danger")
            return render_template("users/form.html", user=None)
        except Exception:
            db.session.rollback()
            current_app.logger.exception("Error creando usuario")
            flash("Ocurrió un error al crear el usuario.", "danger")
            return render_template("users/form.html", user=None)

    return render_template("users/form.html", user=None)


@users_bp.route("/<int:user_id>/edit", methods=["GET", "POST"])
@login_required
def edit(user_id):
    """
    Editar usuario. Super admin puede editar cualquiera; un usuario puede editar su propio perfil (limitado).
    """
    user = User.query.get_or_404(user_id)

    # Permisos: super_admin o el propio usuario
    if not (_is_super_admin() or (current_user.id == user.id)):
        flash("No tienes permisos para editar este usuario.", "warning")
        return redirect(url_for("users.index"))

    if request.method == "POST":
        full_name = (request.form.get("full_name") or "").strip()
        email = (request.form.get("email") or "").strip()
        role = (request.form.get("role") or user.role).strip()
        password = (request.form.get("password") or "").strip()

        if full_name:
            user.full_name = full_name
        user.email = email

        # Solo super_admin puede cambiar roles
        if _is_super_admin():
            user.role = role

        if password:
            user.password_hash = generate_password_hash(password)

        try:
            db.session.commit()
            flash("Usuario actualizado.", "success")
            return redirect(url_for("users.index"))
        except IntegrityError:
            db.session.rollback()
            current_app.logger.exception("IntegrityError actualizando usuario")
            flash("Error al actualizar usuario: datos inválidos.", "danger")
            return render_template("users/form.html", user=user)
        except Exception:
            db.session.rollback()
            current_app.logger.exception("Error actualizando usuario")
            flash("Ocurrió un error al actualizar el usuario.", "danger")
            return render_template("users/form.html", user=user)

    return render_template("users/form.html", user=user)


@users_bp.route("/<int:user_id>/delete", methods=["POST"])
@login_required
def delete(user_id):
    """
    Eliminar usuario. Solo super_admin. No permite eliminar la propia cuenta.
    """
    if not _is_super_admin():
        flash("No tienes permisos para eliminar usuarios.", "warning")
        return redirect(url_for("users.index"))

    if current_user.id == user_id:
        flash("No puedes eliminar tu propia cuenta.", "danger")
        return redirect(url_for("users.index"))

    user = User.query.get_or_404(user_id)
    try:
        db.session.delete(user)
        db.session.commit()
        flash("Usuario eliminado.", "success")
    except IntegrityError:
        db.session.rollback()
        current_app.logger.exception("IntegrityError eliminando usuario")
        flash("No se pudo eliminar el usuario por restricciones en la base de datos.", "danger")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error eliminando usuario")
        flash("Ocurrió un error al eliminar el usuario.", "danger")
    return redirect(url_for("users.index"))


@users_bp.route("/bulk-delete", methods=["POST"])
@login_required
def user_bulk_delete():
    """
    Elimina varios usuarios enviados desde un formulario.
    Espera inputs 'selected_ids' (múltiples checkboxes) o un campo 'selected_ids' con ids separados por comas.
    """
    if not _is_super_admin():
        flash("No tienes permisos para eliminar usuarios.", "warning")
        return redirect(url_for("users.index"))

    # Recibir ids desde form: puede venir como lista o como cadena "1,2,3"
    ids = request.form.getlist("selected_ids") or []
    if not ids:
        ids_raw = request.form.get("selected_ids") or ""
        if ids_raw:
            ids = [s.strip() for s in ids_raw.split(",") if s.strip()]

    try:
        ids = [int(i) for i in ids]
    except Exception:
        flash("IDs inválidos para eliminación masiva.", "warning")
        return redirect(url_for("users.index"))

    if not ids:
        flash("No se seleccionaron usuarios para eliminar.", "warning")
        return redirect(url_for("users.index"))

    # Evitar eliminar la propia cuenta si está incluida
    if current_user.id in ids:
        flash("No puedes eliminar tu propia cuenta en la operación masiva.", "danger")
        return redirect(url_for("users.index"))

    try:
        users_to_delete = User.query.filter(User.id.in_(ids)).all()
        if not users_to_delete:
            flash("No se encontraron usuarios para eliminar.", "warning")
            return redirect(url_for("users.index"))

        for u in users_to_delete:
            db.session.delete(u)
        db.session.commit()
        flash(f"Eliminados {len(users_to_delete)} usuarios.", "success")
    except IntegrityError:
        db.session.rollback()
        current_app.logger.exception("IntegrityError eliminando usuarios en bulk")
        flash("No se pudieron eliminar algunos usuarios por restricciones en la base de datos.", "danger")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error eliminando usuarios en bulk")
        flash("Ocurrió un error al eliminar usuarios. Intenta de nuevo.", "danger")

    return redirect(url_for("users.index"))


# Opcional: endpoint para ver perfil público/edición propia
@users_bp.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    """
    Edición rápida del perfil del usuario autenticado.
    """
    user = current_user

    if request.method == "POST":
        full_name = (request.form.get("full_name") or "").strip()
        email = (request.form.get("email") or "").strip()
        password = (request.form.get("password") or "").strip()

        if full_name:
            user.full_name = full_name
        user.email = email
        if password:
            user.password_hash = generate_password_hash(password)

        try:
            db.session.commit()
            flash("Perfil actualizado.", "success")
            return redirect(url_for("users.profile"))
        except Exception:
            db.session.rollback()
            current_app.logger.exception("Error actualizando perfil")
            flash("Ocurrió un error al actualizar el perfil.", "danger")
            return render_template("auth/profile.html", current_user=user)

    return render_template("auth/profile.html", current_user=user)
