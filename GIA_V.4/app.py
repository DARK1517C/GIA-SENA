# app.py
import os
import logging
from werkzeug.middleware.proxy_fix import ProxyFix
from flask import Flask, url_for, current_app, render_template, send_from_directory

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
    "docente": "Docente",
    "aprendiz": "Aprendiz",
    "visualizador": "Visualizador",
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
    - config: optional dict to override defaults (useful for tests).
    """
    app = Flask(__name__, template_folder="templates", static_folder="static")

    # Load configuration
    cfg = DEFAULT_CONFIG.copy()
    if config:
        cfg.update(config)
    app.config.update(cfg)

    # --- Debug helper: print DB URI and sqlite absolute path (useful to verify which DB file is used)
    db_uri = app.config.get("SQLALCHEMY_DATABASE_URI")
    print("DB URI:", db_uri)
    if db_uri and db_uri.startswith("sqlite"):
        # sqlite:///relative/path.db  -> extract the part after sqlite:///
        sqlite_path = db_uri.replace("sqlite:///", "")
        abs_path = os.path.abspath(sqlite_path)
        print("SQLite file absolute path:", abs_path)
    # ------------------------------------------------------------------------------

    # Ensure upload dir exists
    os.makedirs(app.config["UPLOAD_DIR"], exist_ok=True)

    # Apply proxy fix if behind a proxy (optional)
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=1)

    # Basic logging
    logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(levelname)s %(message)s")
    app.logger.setLevel(logging.DEBUG)

    # Initialize extensions (if you have an extensions.py with init_extensions)
    try:
        import extensions  # type: ignore

        if hasattr(extensions, "init_extensions"):
            extensions.init_extensions(app)
        else:
            if hasattr(extensions, "db") and hasattr(extensions, "login_manager"):
                extensions.db.init_app(app)
                extensions.login_manager.init_app(app)
    except Exception:
        # If extensions module is missing or fails, continue — routes may import db lazily
        app.logger.debug("No extensions initialized at startup or initialization failed.", exc_info=True)

    # Register blueprints from routes package. Use try/except to avoid breaking if a module is missing.
    with app.app_context():
        try:
            from routes import (  # type: ignore
                auth_bp,
                apprentices_bp,
                groups_bp,
                dashboard_bp,
                bitacoras_bp,
                users_bp,
            )
        except Exception:
            auth_bp = apprentices_bp = groups_bp = dashboard_bp = bitacoras_bp = users_bp = None
            try:
                from routes.auth import auth_bp as _auth_bp  # type: ignore
                auth_bp = _auth_bp
            except Exception:
                auth_bp = None
            try:
                from routes.apprentices import apprentices_bp as _apprentices_bp  # type: ignore
                apprentices_bp = _apprentices_bp
            except Exception:
                apprentices_bp = None
            try:
                from routes.groups import groups_bp as _groups_bp  # type: ignore
                groups_bp = _groups_bp
            except Exception:
                groups_bp = None
            try:
                from routes.dashboard import dashboard_bp as _dashboard_bp  # type: ignore
                dashboard_bp = _dashboard_bp
            except Exception:
                dashboard_bp = None
            try:
                from routes.bitacoras import bitacoras_bp as _bitacoras_bp  # type: ignore
                bitacoras_bp = _bitacoras_bp
            except Exception:
                bitacoras_bp = None
            try:
                from routes.users import users_bp as _users_bp  # type: ignore
                users_bp = _users_bp
            except Exception:
                users_bp = None

        # Helper to register if blueprint exists
        def _register(bp):
            if bp:
                try:
                    app.register_blueprint(bp)
                except Exception:
                    app.logger.exception("Failed to register blueprint %s", getattr(bp, "name", str(bp)))

        _register(auth_bp)
        _register(apprentices_bp)
        _register(groups_bp)
        _register(dashboard_bp)
        _register(bitacoras_bp)
        _register(users_bp)

    # Context processors: inject commonly used helpers and variables into templates
    @app.context_processor
    def inject_globals():
        def safe_url(endpoint: str, **values):
            """
            Return url_for(endpoint, **values) or '#' if endpoint is not available.
            Use in templates as: {{ safe_url('users.index') }}
            """
            try:
                return url_for(endpoint, **values)
            except Exception:
                return "#"

        def html_date_value(value):
            """
            Convert a Python date/datetime to YYYY-MM-DD string for <input type="date">.
            Accepts None and returns empty string.
            """
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
            """
            Return True if the given endpoint name is registered in the app.
            Example: has_endpoint('groups.edit') or has_endpoint('bitacoras.index')
            """
            try:
                return name in app.view_functions
            except Exception:
                return False

        def now():
            """
            Return current datetime (naive). Templates can call now().year etc.
            """
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
            # Return 204 No Content if favicon not present to avoid 404 handling
            from flask import Response
            return Response(status=204)

    # Error handlers (optional)
    @app.errorhandler(404)
    def not_found(err):
        return (
            render_template("error.html", title="Página no encontrada", message="La página solicitada no existe."),
            404,
        )

    @app.errorhandler(500)
    def server_error(err):
        # Log full exception to help debugging
        app.logger.exception("Internal Server Error")
        return (
            render_template("error.html", title="Error interno", message="Ha ocurrido un error en el servidor."),
            500,
        )

    return app


# If executed directly, create app and run
if __name__ == "__main__":
    app = create_app()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=app.config.get("DEBUG", False))
