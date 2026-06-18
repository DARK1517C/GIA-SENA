import os
import sqlite3

from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from flask_migrate import Migrate
from sqlalchemy import event
from sqlalchemy.engine import Engine
from sqlalchemy import MetaData
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()
login_manager = LoginManager()
migrate = Migrate()

login_manager.login_view = "auth.login"
login_manager.login_message = None

def init_extensions(app):
    """
    Inicializa extensiones con la app Flask.
    Aplica ajustes para SQLite cuando corresponda.
    """

    app.config.setdefault("SQLALCHEMY_ENGINE_OPTIONS", {"pool_pre_ping": True})

    # Si la URI es sqlite, añadimos connect_args
    if app.config.get("SQLALCHEMY_DATABASE_URI", "").startswith("sqlite"):
        engine_opts = app.config["SQLALCHEMY_ENGINE_OPTIONS"]
        engine_opts.setdefault("connect_args", {})
        engine_opts["connect_args"].update({"timeout": 120, "check_same_thread": False})
        app.config["SQLALCHEMY_ENGINE_OPTIONS"] = engine_opts

    db.init_app(app)
    login_manager.init_app(app)
    # Inicializa Flask-Migrate para gestionar migraciones de esquema de forma controlada
    migrate.init_app(app, db)

# PRAGMA tuning para sqlite connections
@event.listens_for(Engine, "connect")
def configure_sqlite(connection, _record):
    if isinstance(connection, sqlite3.Connection):
        cursor = connection.cursor()
        cursor.execute("PRAGMA journal_mode=WAL")
        cursor.execute("PRAGMA synchronous=NORMAL")
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.execute("PRAGMA temp_store=MEMORY")
        cursor.execute("PRAGMA busy_timeout=120000")
        cursor.close()

naming_convention = {
    "ix": "ix_%(column_0_label)s",
    "uq": "uq_%(table_name)s_%(column_0_name)s",
    "ck": "ck_%(table_name)s_%(constraint_name)s",
    "fk": "fk_%(table_name)s_%(column_0_name)s_%(referred_table_name)s",
    "pk": "pk_%(table_name)s"
}

metadata = MetaData(naming_convention=naming_convention)
db = SQLAlchemy(metadata=metadata)