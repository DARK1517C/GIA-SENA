import os
import logging
from werkzeug.middleware.proxy_fix import ProxyFix
from flask import (
    Flask,
    url_for,
    request,
    render_template,
    send_from_directory,
    redirect,
    flash,
    abort,
    Response,
)
from services.text import normalize_text

# Configuration defaults (override with environment variables)
DEFAULT_CONFIG = {
    "SECRET_KEY": os.environ.get("SECRET_KEY", "dev-secret-key"),
    "SQLALCHEMY_DATABASE_URI": os.environ.get("DATABASE_URL", "sqlite:///gia.db"),
    "SQLALCHEMY_TRACK_MODIFICATIONS": False,
    "UPLOAD_DIR": os.environ.get("UPLOAD_DIR", os.path.join(os.getcwd(), "uploads")),
    "ENV": os.environ.get("FLASK_ENV", "production"),
    "DEBUG": os.environ.get("FLASK_DEBUG", "0") == "1",
}

# Role labels used in templates
ROLE_LABELS = {
    "super_admin": "Super administrador",
    "docente": "Instructor",
    "aprendiz": "Aprendiz",
    "visualizador": "Administrativo",
}

# Example select options used by templates (extend as needed)
FIELD_SELECT_OPTIONS = {
    "ep_modality": [
        ("", "Todas las modalidades"),
        ("contrato_aprendizaje", "Contrato de aprendizaje"),
        ("pasantia", "Pasantía"),
        ("proyecto_productivo", "Proyecto productivo"),
    ],
    "modality": [
        ("", "Todas"),
        ("presencial", "Presencial"),
        ("virtual", "Virtual"),
    ],
}


def create_app(config: dict | None = None) -> Flask:
    """
    Create and configure the Flask application.
    """
    app = Flask(__name__, template_folder="templates", static_folder="static")

    # Load configuration
    cfg = DEFAULT_CONFIG.copy()
    if config:
        cfg.update(config)
    app.config.update(cfg)

    # Debug helpers
    db_uri = app.config.get("SQLALCHEMY_DATABASE_URI")
    app.logger.debug("DB URI: %s", db_uri)
    if db_uri and db_uri.startswith("sqlite"):
        sqlite_path = db_uri.replace("sqlite:///", "")
        app.logger.debug("SQLite file absolute path: %s", os.path.abspath(sqlite_path))

    # Ensure upload dir exists
    os.makedirs(app.config["UPLOAD_DIR"], exist_ok=True)

    # Proxy fix
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)

    # Basic logging
    logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(levelname)s %(message)s")
    app.logger.setLevel(logging.DEBUG if app.config.get("DEBUG") else logging.INFO)

    # Initialize extensions (if you have an extensions.py with init_extensions)
    try:
        import extensions  # type: ignore

        if hasattr(extensions, "init_extensions"):
            extensions.init_extensions(app)
        else:
            if hasattr(extensions, "db") and hasattr(extensions, "login_manager"):
                extensions.db.init_app(app)
                extensions.login_manager.login_message = "Debe iniciar sesion para acceder a esta pagina."
                extensions.login_manager.login_message_category = "warning"
                extensions.login_manager.init_app(app)
    except Exception:
        app.logger.debug("No extensions initialized at startup or initialization failed.", exc_info=True)

    # Register Jinja filter for normalization (used in templates)
    try:
        app.jinja_env.filters["normalize"] = normalize_text
    except Exception:
        app.logger.exception("Failed to register Jinja filter 'normalize'")

    # -------------------------------------------------------------------------
    # Register blueprints safely inside the app context to avoid import cycles
    # -------------------------------------------------------------------------
    with app.app_context():
        def _safe_import(module_path: str, attr: str):
            try:
                module = __import__(module_path, fromlist=[attr])
                return getattr(module, attr)
            except Exception:
                app.logger.debug("Could not import %s.%s", module_path, attr, exc_info=True)
                return None

        # Import blueprints (each wrapped to avoid breaking startup if a module fails)
        auth_bp = _safe_import("routes.auth", "auth_bp")
        apprentices_bp = _safe_import("routes.apprentices", "apprentices_bp")
        groups_bp = _safe_import("routes.groups", "groups_bp")
        dashboard_bp = _safe_import("routes.dashboard", "dashboard_bp")
        bitacoras_bp = _safe_import("routes.bitacoras", "bitacoras_bp")
        users_bp = _safe_import("routes.users", "users_bp")
        evidences_bp = _safe_import("routes.evidences", "evidences_bp")
        reports_bp = _safe_import("routes.reports", "reports_bp")

        def _register(bp):
            if bp:
                try:
                    app.register_blueprint(bp)
                    app.logger.debug("Registered blueprint: %s", getattr(bp, "name", str(bp)))
                except Exception:
                    app.logger.exception("Failed to register blueprint %s", getattr(bp, "name", str(bp)))

        # Register in a logical order
        _register(auth_bp)
        _register(dashboard_bp)
        _register(apprentices_bp)
        _register(groups_bp)
        _register(bitacoras_bp)
        _register(evidences_bp)   # ensure evidences blueprint is registered here
        _register(reports_bp)
        _register(users_bp)

        # Optional: print url_map when debugging to verify endpoints (including evidences.index)
        if app.config.get("DEBUG"):
            try:
                app.logger.debug("Registered endpoints:")
                for rule in app.url_map.iter_rules():
                    app.logger.debug("%s -> %s", rule.endpoint, rule)
            except Exception:
                app.logger.debug("Could not list url_map", exc_info=True)

    # ---------------------------
    # Restricción central para rol 'aprendiz' (lista blanca)
    # ---------------------------
    @app.before_request
    def restrict_aprendiz_endpoints():
        from flask_login import current_user

        if not current_user or not getattr(current_user, "is_authenticated", False):
            return

        if getattr(current_user, "role", None) != "aprendiz":
            return

        allowed_endpoints = {
            "auth.profile",
            "auth.update_profile",
            "auth.login",
            "auth.logout",
            "static",
        }

        endpoint = (request.endpoint or "").strip()

        # Allow evidences blueprint endpoints
        if endpoint.startswith("evidences.") or endpoint in allowed_endpoints:
            return

        app.logger.warning(
            "Acceso denegado a aprendiz: user=%s endpoint=%s path=%s",
            getattr(current_user, "id", None),
            endpoint,
            request.path,
        )

        for candidate in ("evidences.index", "apprentices.index"):
            if candidate in app.view_functions:
                try:
                    flash("No tienes permisos para acceder a esa sección.", "warning")
                    return redirect(url_for(candidate))
                except Exception:
                    app.logger.debug("Fallo al redirigir a %s", candidate, exc_info=True)
                    break

        abort(403)

    # ---------------------------
    # Error handlers
    # ---------------------------
    @app.errorhandler(403)
    def forbidden(err):
        for candidate in ("evidences.index", "apprentices.index"):
            if candidate in app.view_functions:
                try:
                    return redirect(url_for(candidate))
                except Exception:
                    app.logger.debug("No se pudo redirigir a %s", candidate, exc_info=True)

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
        return (
            render_template("error.html", title="Página no encontrada", message="La página solicitada no existe."),
            404,
        )

    @app.errorhandler(500)
    def server_error(err):
        app.logger.exception("Internal Server Error")
        return (
            render_template("error.html", title="Error interno", message="Ha ocurrido un error en el servidor."),
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

        return {
            "role_labels": ROLE_LABELS,
            "field_select_options": FIELD_SELECT_OPTIONS,
            "html_date_value": html_date_value,
            "safe_url": safe_url,
            "has_endpoint": has_endpoint,
            "now": now,
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
