import os
import logging

from flask import (
    Flask,
    current_app,
    url_for,
    request,
    render_template,
    send_from_directory,
    redirect,
    flash,
    abort,
    Response,
)
from werkzeug.middleware.proxy_fix import ProxyFix
from flask_login import current_user
from flask_wtf.csrf import CSRFError

from config import Config
from extensions import init_extensions
from services.text import normalize_text
from services.permissions import ROLE_LABELS
from services.utils_base import format_date_value


# =============================================================================
# APLICACIÓN
# =============================================================================

def create_app(config: dict | None = None) -> Flask:
    """
    Crea e inicializa la aplicación Flask.
    """

    app = Flask(
        __name__,
        template_folder="templates",
        static_folder="static",
    )

    # -------------------------------------------------------------------------
    # Configuración
    # -------------------------------------------------------------------------

    app.config.from_object(Config)

    if config:
        app.config.update(config)

    # -------------------------------------------------------------------------
    # Seguridad de configuración
    # -------------------------------------------------------------------------

    if not app.config.get("SECRET_KEY") and not app.config.get("TESTING"):
        raise RuntimeError(
            "SECRET_KEY es obligatoria. Configure la variable de entorno SECRET_KEY."
        )

    db_uri = app.config.get("SQLALCHEMY_DATABASE_URI")

    # Nunca registrar credenciales de base de datos. Solo se informa el motor.
    if db_uri:
        try:
            scheme = db_uri.split(":", 1)[0]
            app.logger.debug("Database backend: %s", scheme)
        except Exception:
            app.logger.debug("Database backend configurado")

    # -------------------------------------------------------------------------
    # Directorio de archivos
    # -------------------------------------------------------------------------

    os.makedirs(
        app.config["UPLOAD_DIR"],
        exist_ok=True,
    )

    # -------------------------------------------------------------------------
    # Proxy
    # -------------------------------------------------------------------------

    app.wsgi_app = ProxyFix(
        app.wsgi_app,
        x_for=1,
        x_proto=1,
        x_host=1,
        x_port=1,
    )

    # -------------------------------------------------------------------------
    # Logging
    # -------------------------------------------------------------------------

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    app.logger.setLevel(
        logging.DEBUG if app.config.get("DEBUG") else logging.INFO
    )

    # -------------------------------------------------------------------------
    # Extensiones
    # -------------------------------------------------------------------------

    init_extensions(app)

    # -------------------------------------------------------------------------
    # Filtros Jinja
    # -------------------------------------------------------------------------

    try:
        app.jinja_env.filters["normalize"] = normalize_text
        app.jinja_env.filters["date_dmY"] = format_date_value
    except Exception:
        app.logger.exception(
            "Failed to register Jinja filters"
        )

    @app.context_processor
    def inject_notification_context():
        """Expose unread count plus a small recent-notification preview to the topbar."""
        try:
            from models import Notification

            if not current_user.is_authenticated:
                return {
                    "unread_notification_count": 0,
                    "recent_notifications": [],
                }

            unread_count = (
                Notification.query
                .filter_by(user_id=current_user.id, is_read=False)
                .count()
            )
            recent_notifications = (
                Notification.query
                .filter_by(user_id=current_user.id)
                .order_by(Notification.created_at.desc())
                .limit(6)
                .all()
            )
            return {
                "unread_notification_count": unread_count,
                "recent_notifications": recent_notifications,
            }
        except Exception:
            app.logger.exception("No fue posible cargar el resumen de notificaciones.")
            return {
                "unread_notification_count": 0,
                "recent_notifications": [],
            }

    # -------------------------------------------------------------------------
    # Registro de Blueprints
    # -------------------------------------------------------------------------
    # Importación explícita y centralizada: evita que routes/__init__.py
    # produzca imports laterales y permite diagnosticar exactamente qué
    # blueprint no pudo cargarse durante el bootstrap.
    blueprint_modules = (
        ("auth", "auth_bp"),
        ("dashboard", "dashboard_bp"),
        ("apprentices", "apprentices_bp"),
        ("groups", "groups_bp"),
        ("evidences", "evidences_bp"),
        ("evidence_admin", "evidence_admin_bp"),
        ("users", "users_bp"),
        ("notifications", "notifications_bp"),
        ("certification", "certification_bp"),
    )

    for module_name, blueprint_name in blueprint_modules:
        try:
            module = __import__(
                f"routes.{module_name}",
                fromlist=[blueprint_name],
            )
            blueprint = getattr(module, blueprint_name)
        except Exception as exc:
            app.logger.critical(
                "No se pudo cargar el blueprint routes.%s (%s): %s",
                module_name,
                blueprint_name,
                exc,
                exc_info=True,
            )
            raise RuntimeError(
                f"Error de bootstrap: no se pudo cargar "
                f"routes.{module_name}.{blueprint_name}"
            ) from exc

        try:
            app.register_blueprint(blueprint)
        except Exception as exc:
            app.logger.critical(
                "No se pudo registrar el blueprint %s: %s",
                blueprint_name,
                exc,
                exc_info=True,
            )
            raise RuntimeError(
                f"Error de bootstrap: no se pudo registrar "
                f"el blueprint {blueprint_name}"
            ) from exc

        app.logger.info(
            "Blueprint registrado: %s",
            blueprint.name,
        )

    if app.config["DEBUG"]:
        app.logger.debug("Endpoints registrados:")
        for rule in app.url_map.iter_rules():
            app.logger.debug("%s -> %s", rule.endpoint, rule)

    # -------------------------------------------------------------------------
    # Restricción central para aprendices
    # -------------------------------------------------------------------------
    @app.before_request
    def restrict_aprendiz_endpoints():

        from flask_login import current_user

        if not current_user.is_authenticated:
            return

        if current_user.role != "APPRENTICE":
            return

        endpoint = request.endpoint or ""

        allowed = {
            "auth.login",
            "auth.logout",
            "auth.profile",
            "users.profile",
            "static",
        }

        if endpoint in allowed:
            return

        if endpoint.startswith("evidences."):
            return

        current_app.logger.warning(
            "Acceso denegado. Usuario=%s Endpoint=%s Ruta=%s",
            current_user.id,
            endpoint,
            request.path,
        )

        flash(
            "No tienes permisos para acceder a esa sección.",
            "warning",
        )

        if "evidences.index" in app.view_functions:
            return redirect(url_for("evidences.index"))

        abort(403)

    # -------------------------------------------------------------------------
    # Manejo centralizado de errores
    # -------------------------------------------------------------------------
    @app.errorhandler(CSRFError)
    def csrf_error(err):
        app.logger.warning(
            "Solicitud rechazada por CSRF: method=%s path=%s user=%s",
            request.method,
            request.path,
            getattr(current_user, "id", None),
        )
        return (
            render_template(
                "error.html",
                title="Solicitud no válida",
                message="La sesión de seguridad expiró o el formulario no es válido. Recarga la página e inténtalo nuevamente.",
            ),
            400,
        )

    @app.errorhandler(400)
    def bad_request(err):
        return (
            render_template(
                "error.html",
                title="Solicitud no válida",
                message="La solicitud no pudo procesarse.",
            ),
            400,
        )

    @app.errorhandler(403)
    def forbidden(err):
        app.logger.warning(
            "403 Forbidden: method=%s path=%s endpoint=%s user=%s",
            request.method,
            request.path,
            request.endpoint,
            getattr(current_user, "id", None),
        )
        return (
            render_template(
                "error.html",
                title="Acceso denegado",
                message="No tienes permisos para ver esta página. Si crees que es un error, contacta al administrador.",
            ),
            403,
        )

    @app.errorhandler(404)
    def not_found(err):
        app.logger.info(
            "404 Not Found: method=%s path=%s endpoint=%s",
            request.method,
            request.path,
            request.endpoint,
        )
        return (
            render_template(
                "error.html",
                title="Página no encontrada",
                message="La página solicitada no existe.",
            ),
            404,
        )

    @app.errorhandler(413)
    def request_too_large(err):
        app.logger.warning(
            "413 Request Entity Too Large: method=%s path=%s",
            request.method,
            request.path,
        )
        return (
            render_template(
                "error.html",
                title="Archivo demasiado grande",
                message="El archivo o la solicitud supera el tamaño máximo permitido.",
            ),
            413,
        )

    @app.errorhandler(500)
    def server_error(err):
        # Una excepción de una petición puede dejar una transacción abierta.
        # El rollback aquí evita arrastrar ese estado a la siguiente petición.
        try:
            from extensions import db
            db.session.rollback()
        except Exception:
            app.logger.exception("No fue posible hacer rollback de la sesión DB")

        app.logger.exception(
            "Internal Server Error: method=%s path=%s endpoint=%s",
            request.method,
            request.path,
            request.endpoint,
        )
        return (
            render_template(
                "error.html",
                title="Error interno",
                message="Ha ocurrido un error en el servidor. Inténtelo nuevamente.",
            ),
            500,
        )

    # Context processors: inject commonly used helpers and variables into templates
    @app.context_processor
    def inject_globals():
        def safe_url(endpoint: str, **values):
            try:
                return url_for(endpoint, **values)
            except Exception:
                return "#"

        def html_date_value(value):
            try:
                if value is None:
                    return ""
                from datetime import date, datetime

                if isinstance(value, (date, datetime)):
                    return value.strftime("%Y-%m-%d")
                return str(value)
            except Exception:
                return ""

        def has_endpoint(name: str) -> bool:
            try:
                return name in app.view_functions
            except Exception:
                return False

        def now():
            from datetime import datetime
            return datetime.now()

        def format_datetime_local(value):
            """Format an aware/naive datetime for the institutional display timezone."""
            if value is None:
                return ""
            try:
                from datetime import datetime, timezone
                from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

                if not isinstance(value, datetime):
                    return str(value)

                tz_name = app.config.get("DISPLAY_TIMEZONE", "America/Bogota")
                try:
                    display_tz = ZoneInfo(tz_name)
                except ZoneInfoNotFoundError:
                    display_tz = timezone.utc

                if value.tzinfo is None:
                    # SQLite may return timezone-aware columns as naive datetimes.
                    value = value.replace(tzinfo=timezone.utc)
                else:
                    value = value.astimezone(timezone.utc)

                local_value = value.astimezone(display_tz)
                return local_value.strftime("%d/%m/%Y %H:%M:%S")
            except Exception:
                return str(value)

        return {
            "role_labels": ROLE_LABELS,
            "html_date_value": html_date_value,
            "safe_url": safe_url,
            "has_endpoint": has_endpoint,
            "now": now,
            "format_datetime_local": format_datetime_local,
        }

    # Serve favicon to avoid 404s triggering error handlers
    @app.route("/favicon.ico")
    def favicon():
        try:
            return send_from_directory(app.static_folder, "favicon.ico")
        except Exception:
            return Response(status=204)

    return app


# If executed directly, create app and run
if __name__ == "__main__":
    app = create_app()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=app.config.get("DEBUG", False))
