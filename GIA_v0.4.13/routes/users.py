# =============================================================================
# IMPORTACIONES
# =============================================================================

from flask import (
    Blueprint,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    current_app,
)

from flask_login import login_required, current_user

from werkzeug.security import generate_password_hash

from sqlalchemy import or_
from sqlalchemy.exc import IntegrityError

from extensions import db
from models import Apprentice, User
from services.auth_helpers import permission_required
from services.permissions import (
    ROLE_LABELS,
    ROLES,
    has_permission,
)
from catalogs.user import UserDocumentType
from catalogs.apprentice import EpModality, SofiaStatus
from catalogs.common_catalogs import ProgramLevel, DocumentType
from catalogs.display import get_label


# =============================================================================
# BLUEPRINT
# =============================================================================

users_bp = Blueprint(
    "users",
    __name__,
    url_prefix="/users",
)


# =============================================================================
# ROLES
# =============================================================================

AVAILABLE_ROLES = list(ROLES)


# =============================================================================
# HELPERS
# =============================================================================

def _profile_catalog_labels(apprentice_record):
    """Devuelve etiquetas de catálogo para mostrar en el perfil.

    Este helper debe vivir a nivel de módulo porque la ruta de perfil y sus
    distintos caminos de renderizado lo necesitan tanto en GET como en POST.
    """
    if apprentice_record is None:
        return {}
    return {
        "program_level_label": (
            get_label(ProgramLevel, apprentice_record.program_level)
            if apprentice_record.program_level else None
        ),
        "ep_modality_label": (
            get_label(EpModality, apprentice_record.ep_modality)
            if apprentice_record.ep_modality else None
        ),
        "sofia_status_label": (
            get_label(SofiaStatus, apprentice_record.sofia_status)
            if apprentice_record.sofia_status else None
        ),
        "document_type_label": (
            get_label(DocumentType, apprentice_record.document_type)
            if apprentice_record.document_type else None
        ),
    }


def _check_admin():
    """
    Verifica si el usuario actual puede administrar usuarios.
    """

    return has_permission("users.manage")


# =============================================================================
# LISTADO DE USUARIOS
# =============================================================================

@users_bp.route("/", methods=["GET"])
@login_required
@permission_required("users.manage")
def index():

    if not _check_admin():
        flash("No tienes permisos para acceder a esta sección.", "warning")
        return redirect(url_for("dashboard.index"))

    search = (request.args.get("search") or "").strip()
    role_filter = (request.args.get("role") or "").strip()

    query = User.query

    if search:
        pattern = f"%{search}%"

        query = query.filter(
            or_(
                User.document_number.ilike(pattern),
                User.first_names.ilike(pattern),
                User.last_names.ilike(pattern),
                User.email.ilike(pattern),
            )
        )

    if role_filter:
        query = query.filter(User.role == role_filter)

    users = query.order_by(User.first_names, User.last_names).all()

    return render_template(
        "users/index.html",
        users=users,
        role_labels=ROLE_LABELS,
        available_roles=AVAILABLE_ROLES,
    )


# =============================================================================
# CREAR USUARIO
# =============================================================================

@users_bp.route("/create", methods=["GET", "POST"])
@login_required
@permission_required("users.manage")
def create():

    if not _check_admin():
        flash("No tienes permisos para crear usuarios.", "warning")
        return redirect(url_for("users.index"))

    if request.method == "POST":
        document_type = (request.form.get("document_type") or "NATIONAL_ID").strip()
        document_number = (request.form.get("document_number") or "").strip()
        full_name = (request.form.get("full_name") or "").strip()
        email = (request.form.get("email") or "").strip()
        role = (request.form.get("role") or "CENTER_STAFF").strip()
        password = (request.form.get("password") or "").strip()

        parts = full_name.split(maxsplit=1)
        first_names = parts[0] if parts else ""
        last_names = parts[1] if len(parts) > 1 else ""

        if not document_number or not full_name or not password:
            flash("Documento, nombre completo y contraseña son obligatorios.", "danger")
            return redirect(url_for("users.index"))

        if len(password) < 8:
            flash("La contraseña debe tener al menos 8 caracteres.", "danger")
            return redirect(url_for("users.index"))

        if User.query.filter_by(document_number=document_number).first():
            flash("El número de documento ya existe.", "danger")
            return redirect(url_for("users.index"))

        if email and User.query.filter_by(email=email.lower()).first():
            flash("El correo electrónico ya existe.", "danger")
            return redirect(url_for("users.index"))

        try:
            user = User(
                document_type=document_type,
                document_number=document_number,
                first_names=first_names,
                last_names=last_names,
                email=email or None,
                role=role,
                password_hash=generate_password_hash(password),
            )
            db.session.add(user)
            db.session.commit()
            flash("Usuario creado correctamente.", "success")
        except IntegrityError:
            db.session.rollback()
            current_app.logger.exception("IntegrityError creando usuario")
            flash("Error al crear el usuario. Verifica los datos únicos.", "danger")
        except Exception:
            db.session.rollback()
            current_app.logger.exception("Error creando usuario")
            flash("Ocurrió un error al crear el usuario.", "danger")

    return redirect(url_for("users.index"))

# =============================================================================
# ELIMINAR USUARIO
# =============================================================================

@users_bp.route("/<int:user_id>/delete", methods=["POST"])
@login_required
@permission_required("users.manage")
def delete(user_id):

    if not _check_admin():
        flash(
            "No tienes permisos para eliminar usuarios.",
            "warning",
        )
        return redirect(url_for("users.index"))

    if current_user.id == user_id:
        flash(
            "No puedes eliminar tu propia cuenta.",
            "danger",
        )
        return redirect(url_for("users.index"))

    user = User.query.get_or_404(user_id)

    try:

        login_identifier = user.login_identifier
        role = user.role

        # Si existe un aprendiz asociado, desvincularlo primero
        Apprentice.query.filter_by(
            student_user_id=user.id
        ).update(
            {"student_user_id": None},
            synchronize_session=False,
        )

        db.session.delete(user)
        db.session.commit()

        current_app.logger.info(
            "Usuario eliminado: id=%s identificador=%s role=%s eliminado_por=%s",
            user.id,
            login_identifier,
            role,
            current_user.id,
        )

        flash(
            "Usuario eliminado correctamente.",
            "success",
        )

    except IntegrityError:

        db.session.rollback()

        current_app.logger.exception(
            "IntegrityError eliminando usuario %s",
            user_id,
        )

        flash(
            "No fue posible eliminar el usuario por restricciones de la base de datos.",
            "danger",
        )

    except Exception:

        db.session.rollback()

        current_app.logger.exception(
            "Error eliminando usuario %s",
            user_id,
        )

        flash(
            "Ocurrió un error al eliminar el usuario.",
            "danger",
        )

    return redirect(url_for("users.index"))

# =============================================================================
# ELIMINACIÓN MASIVA
# =============================================================================

@users_bp.route("/bulk-delete", methods=["POST"])
@login_required
@permission_required("users.manage")
def user_bulk_delete():

    if not _check_admin():
        flash("No tienes permisos para eliminar usuarios.", "warning")
        return redirect(url_for("users.index"))

    ids = request.form.getlist("selected_ids") or []

    if not ids:
        ids_raw = (request.form.get("selected_ids") or "").strip()

        if ids_raw:
            ids = [
                value.strip()
                for value in ids_raw.split(",")
                if value.strip()
            ]

    try:
        ids = [int(value) for value in ids]

    except Exception:
        flash("IDs inválidos.", "warning")
        return redirect(url_for("users.index"))

    if not ids:
        flash(
            "No se seleccionaron usuarios para eliminar.",
            "warning",
        )
        return redirect(url_for("users.index"))

    if current_user.id in ids:
        flash(
            "No puedes eliminar tu propia cuenta.",
            "danger",
        )
        return redirect(url_for("users.index"))

    try:

        users = User.query.filter(
            User.id.in_(ids)
        ).all()

        if not users:

            flash(
                "No se encontraron usuarios.",
                "warning",
            )
            return redirect(url_for("users.index"))

        for user in users:

            Apprentice.query.filter_by(
                student_user_id=user.id
            ).update(
                {"student_user_id": None},
                synchronize_session=False,
            )

            db.session.delete(user)

        db.session.commit()

        current_app.logger.info(
            "Usuarios eliminados en lote: total=%s eliminado_por=%s",
            len(users),
            current_user.id,
        )

        flash(
            f"Se eliminaron {len(users)} usuarios.",
            "success",
        )

    except IntegrityError:

        db.session.rollback()

        current_app.logger.exception(
            "IntegrityError eliminando usuarios en lote"
        )

        flash(
            "No fue posible eliminar algunos usuarios por restricciones de la base de datos.",
            "danger",
        )

    except Exception:

        db.session.rollback()

        current_app.logger.exception(
            "Error eliminando usuarios en lote"
        )

        flash(
            "Ocurrió un error al eliminar usuarios.",
            "danger",
        )

    return redirect(url_for("users.index"))


# =============================================================================
# PERFIL DEL USUARIO
# =============================================================================

@users_bp.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    """Consulta y actualiza de forma segura el perfil del usuario autenticado."""

    user = current_user
    apprentice = (
        Apprentice.query.filter_by(student_user_id=user.id).first()
        if getattr(user, "role", None) == "APPRENTICE"
        else None
    )

    if request.method == "POST":
        # El formulario tiene dos operaciones distintas: perfil y contraseña.
        wants_password_change = any(
            (request.form.get(field) or "").strip()
            for field in ("current_password", "new_password", "confirm_password")
        )

        try:
            if wants_password_change:
                current_password = request.form.get("current_password") or ""
                new_password = request.form.get("new_password") or ""
                confirm_password = request.form.get("confirm_password") or ""

                if not current_password or not new_password or not confirm_password:
                    flash("Para cambiar la contraseña debes diligenciar los tres campos.", "warning")
                    raise ValueError("complete_password_form")

                if not user.check_password(current_password):
                    flash("La contraseña actual no es correcta.", "danger")
                    raise ValueError("invalid_current_password")

                if len(new_password) < 8:
                    flash("La nueva contraseña debe tener al menos 8 caracteres.", "warning")
                    raise ValueError("weak_password")

                if new_password != confirm_password:
                    flash("La confirmación de la nueva contraseña no coincide.", "warning")
                    raise ValueError("password_mismatch")

                user.set_password(new_password)
                db.session.commit()
                flash("Contraseña actualizada correctamente.", "success")
                return redirect(url_for("users.profile"))

            full_name = (request.form.get("full_name") or "").strip()
            email = (request.form.get("email") or "").strip().lower()
            phone = (request.form.get("phone") or "").strip()

            if full_name:
                parts = full_name.split(maxsplit=1)
                user.first_names = parts[0]
                user.last_names = parts[1] if len(parts) > 1 else ""
            else:
                flash("El nombre completo es obligatorio.", "warning")
                raise ValueError("missing_name")

            if email:
                user.email = email
            else:
                user.email = None

            if hasattr(user, "phone"):
                user.phone = phone or None

            # Para aprendices, sincronizar email/teléfono con el registro académico.
            # El rol nunca se modifica desde el propio perfil.
            if apprentice is not None:
                if email:
                    apprentice.email = email
                apprentice.phone = phone or None

            db.session.commit()
            flash("Perfil actualizado correctamente.", "success")
            return redirect(url_for("users.profile"))

        except ValueError:
            db.session.rollback()
            return render_template(
                "auth/profile.html",
                current_user=user,
                apprentice=apprentice,
                **_profile_catalog_labels(apprentice),
                back_url=url_for("dashboard.index"),
                back_label="Volver",
            )
        except Exception:
            db.session.rollback()
            current_app.logger.exception("Error actualizando el perfil del usuario %s", user.id)
            flash("Ocurrió un error al actualizar el perfil.", "danger")
            return render_template(
                "auth/profile.html",
                current_user=user,
                apprentice=apprentice,
                **_profile_catalog_labels(apprentice),
                back_url=url_for("dashboard.index"),
                back_label="Volver",
            )

    return render_template(
        "auth/profile.html",
        current_user=user,
        apprentice=apprentice,
        **_profile_catalog_labels(apprentice),
        back_url=url_for("dashboard.index"),
        back_label="Volver",
    )
