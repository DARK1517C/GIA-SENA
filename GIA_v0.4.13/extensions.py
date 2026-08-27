# extensions.py

import sqlite3
from copy import deepcopy

from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_wtf.csrf import CSRFProtect
from flask_migrate import Migrate
from sqlalchemy import event, MetaData
from sqlalchemy.engine import Engine


naming_convention = {
    "ix": "ix_%(column_0_label)s",
    "uq": "uq_%(table_name)s_%(column_0_name)s",
    "ck": "ck_%(table_name)s_%(constraint_name)s",
    "fk": "fk_%(table_name)s_%(column_0_name)s_%(referred_table_name)s",
    "pk": "pk_%(table_name)s",
}

metadata = MetaData(naming_convention=naming_convention)


db = SQLAlchemy(metadata=metadata)
login_manager = LoginManager()
migrate = Migrate()
csrf = CSRFProtect()

login_manager.login_view = "auth.login"
login_manager.login_message = None


def init_extensions(app):
    """Inicializa extensiones sin compartir estado mutable entre app instances."""
    engine_opts = deepcopy(
        app.config.get("SQLALCHEMY_ENGINE_OPTIONS") or {"pool_pre_ping": True}
    )

    if app.config.get("SQLALCHEMY_DATABASE_URI", "").startswith("sqlite"):
        connect_args = dict(engine_opts.get("connect_args") or {})
        connect_args.update({
            "timeout": 120,
            "check_same_thread": False,
        })
        engine_opts["connect_args"] = connect_args

    app.config["SQLALCHEMY_ENGINE_OPTIONS"] = engine_opts

    db.init_app(app)
    login_manager.init_app(app)
    migrate.init_app(app, db)
    csrf.init_app(app)


@event.listens_for(Engine, "connect")
def configure_sqlite(connection, _record):
    """Configura SQLite sin afectar conexiones PostgreSQL u otros motores."""
    if not isinstance(connection, sqlite3.Connection):
        return

    cursor = connection.cursor()
    try:
        cursor.execute("PRAGMA journal_mode=WAL")
        cursor.execute("PRAGMA synchronous=NORMAL")
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.execute("PRAGMA temp_store=MEMORY")
        cursor.execute("PRAGMA busy_timeout=120000")
    finally:
        cursor.close()
