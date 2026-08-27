from urllib.parse import urljoin, urlparse

from flask import (
    Blueprint,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    url_for,
)

from flask_login import (
    current_user,
    login_required,
    login_user,
    logout_user,
)

from extensions import db
from models import User


auth_bp = Blueprint(
    "auth",
    __name__,
    url_prefix="/auth",
)


# =============================================================================
# UTILIDADES
# =============================================================================


def is_safe_url(target: str | None) -> bool:
    """
    Evita redirecciones abiertas.

    Solo permite destinos pertenecientes al mismo host
    desde el que se está ejecutando la aplicación.
    """
    if not target:
        return False

    ref_url = urlparse(request.host_url)

    test_url = urlparse(
        urljoin(
            request.host_url,
            target,
        )
    )

    return (
        test_url.scheme in ("http", "https")
        and ref_url.netloc == test_url.netloc
    )


def _find_user_by_login(identifier: str | None) -> User | None:
    """
    Busca un usuario por su identificador de acceso.

    El modelo User actual utiliza:

        - email
        - document_number

    El correo tiene prioridad.
    Si no existe coincidencia por correo,
    se intenta con el número de documento.
    """
    identifier = (identifier or "").strip()

    if not identifier:
        return None

    # ------------------------------------------------------------------
    # Buscar por correo
    # ------------------------------------------------------------------

    user = (
        User.query
        .filter(
            User.email == identifier.lower()
        )
        .first()
    )

    if user is not None:
        return user

    # ------------------------------------------------------------------
    # Buscar por documento
    # ------------------------------------------------------------------

    return (
        User.query
        .filter(
            User.document_number == identifier.upper()
        )
        .first()
    )


# =============================================================================
# LOGIN
# =============================================================================


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    """
    Autentica un usuario mediante:

        - correo electrónico
        - número de documento

    La contraseña se verifica mediante User.check_password().
    """

    if request.method == "POST":

        # ------------------------------------------------------------------
        # Identificador
        # ------------------------------------------------------------------

        identifier = (
            request.form.get("identifier")
            or request.form.get("login")
            or request.form.get("email")
            or request.form.get("document_number")
            or ""
        ).strip()

        # ------------------------------------------------------------------
        # Contraseña
        # ------------------------------------------------------------------

        password = (
            request.form.get("password")
            or ""
        )

        # ------------------------------------------------------------------
        # Validación básica
        # ------------------------------------------------------------------

        if not identifier or not password:
            flash(
                "Debe ingresar su correo o número de documento "
                "y contraseña.",
                "danger",
            )

            return render_template(
                "auth/login.html",
            )

        # ------------------------------------------------------------------
        # Buscar usuario
        # ------------------------------------------------------------------

        try:
            user = _find_user_by_login(identifier)

        except Exception:
            current_app.logger.exception(
                "Error consultando usuario durante el login. "
                "Identificador=%s",
                identifier,
            )

            flash(
                "No fue posible validar las credenciales. "
                "Inténtelo nuevamente.",
                "danger",
            )

            return render_template(
                "auth/login.html",
            )

        # ------------------------------------------------------------------
        # Verificar contraseña
        # ------------------------------------------------------------------

        if user is not None and user.check_password(password):

            # --------------------------------------------------------------
            # Verificar estado
            # --------------------------------------------------------------

            if not user.is_active:
                current_app.logger.warning(
                    "Intento de inicio de sesión con cuenta "
                    "no activa: usuario=%s id=%s estado=%s",
                    user.login_identifier,
                    user.id,
                    user.status,
                )

                flash(
                    "Su cuenta no está activa. "
                    "Comuníquese con el administrador.",
                    "warning",
                )

                return render_template(
                    "auth/login.html",
                )

            # --------------------------------------------------------------
            # Crear sesión
            # --------------------------------------------------------------

            login_user(user)

            # --------------------------------------------------------------
            # Actualizar último acceso
            # --------------------------------------------------------------

            try:
                user.touch_last_login()

                db.session.commit()

            except Exception:
                db.session.rollback()

                current_app.logger.exception(
                    "No fue posible actualizar "
                    "last_login_at para usuario id=%s",
                    user.id,
                )

            # --------------------------------------------------------------
            # Registrar acceso
            # --------------------------------------------------------------

            current_app.logger.info(
                "Inicio de sesión: usuario=%s id=%s rol=%s",
                user.login_identifier,
                user.id,
                user.role,
            )

            # --------------------------------------------------------------
            # Redirección solicitada
            # --------------------------------------------------------------

            next_url = (
                request.args.get("next")
                or request.form.get("next")
            )

            if next_url and is_safe_url(next_url):
                return redirect(next_url)

            # --------------------------------------------------------------
            # Aprendiz
            # --------------------------------------------------------------

            if user.is_apprentice:
                return redirect(
                    url_for("evidences.index")
                )

            # --------------------------------------------------------------
            # Resto de usuarios
            # --------------------------------------------------------------

            return redirect(
                url_for("dashboard.index")
            )

        # ------------------------------------------------------------------
        # Credenciales incorrectas
        # ------------------------------------------------------------------

        current_app.logger.warning(
            "Intento de inicio de sesión fallido "
            "para identificador '%s'",
            identifier,
        )

        flash(
            "Correo/documento o contraseña incorrectos.",
            "danger",
        )

    return render_template(
        "auth/login.html",
    )


# =============================================================================
# LOGOUT
# =============================================================================


@auth_bp.route("/logout", methods=["POST"])
@login_required
def logout():
    """
    Cierra la sesión actual.
    """

    user_identifier = getattr(
        current_user,
        "login_identifier",
        None,
    )

    user_id = getattr(
        current_user,
        "id",
        None,
    )

    logout_user()

    current_app.logger.info(
        "Cierre de sesión: usuario=%s id=%s",
        user_identifier,
        user_id,
    )

    flash(
        "Sesión cerrada correctamente.",
        "success",
    )

    return redirect(
        url_for("auth.login")
    )