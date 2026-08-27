import os
from urllib.parse import urlsplit, urlunsplit

from dotenv import load_dotenv

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# Carga la configuración local desde .env cuando existe.
# Nunca sobrescribe variables de entorno ya definidas por el proceso.
load_dotenv(os.path.join(BASE_DIR, ".env"), override=False)


def _normalize_database_url(value: str | None) -> str | None:
    """Normaliza URLs aceptadas por proveedores y SQLAlchemy.

    Para SQLite con ruta relativa, resuelve el archivo respecto a BASE_DIR
    en lugar del directorio de trabajo del proceso. Esto evita fallos como
    ``sqlite:///instance/gia.db`` cuando Flask CLI/Alembic se ejecuta desde
    otra ubicación.
    """
    if not value:
        return None

    value = value.strip()

    if value.startswith("postgres://"):
        value = "postgresql+psycopg://" + value[len("postgres://"):]
    elif value.startswith("postgresql://"):
        value = "postgresql+psycopg://" + value[len("postgresql://"):]
    elif value.startswith("sqlite:///") and not value.startswith("sqlite:////"):
        raw_path = value[len("sqlite:///"):]
        # Windows drive-letter paths are handled as absolute paths; other
        # relative paths are anchored to the project root.
        if not os.path.isabs(raw_path):
            raw_path = os.path.join(BASE_DIR, raw_path)
        raw_path = os.path.abspath(raw_path)
        os.makedirs(os.path.dirname(raw_path), exist_ok=True)
        return "sqlite:///" + raw_path.replace("\\", "/")

    return value


class Config:
    DEBUG = os.getenv("FLASK_DEBUG", "0").lower() in {"1", "true", "yes", "on"}

    # Nunca usar una clave conocida en un entorno no-debug.
    # En desarrollo se conserva un fallback estable para facilitar el arranque
    # local; producción exige SECRET_KEY explícita.
    SECRET_KEY = os.getenv("SECRET_KEY") or (
        "gia-development-secret-key" if DEBUG else None
    )
    TESTING = False

    MAX_CONTENT_LENGTH = 16 * 1024 * 1024
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # Pooling básico compatible con SQLite y PostgreSQL.
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
    }

    UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
    INSTANCE_DIR = os.path.join(BASE_DIR, "instance")
    os.makedirs(INSTANCE_DIR, exist_ok=True)

    _database_url = _normalize_database_url(os.getenv("DATABASE_URL"))

    SQLALCHEMY_DATABASE_URI = _database_url or (
        f"sqlite:///{os.path.join(INSTANCE_DIR, 'gia.db')}"
    )

    # Cookies endurecidas para despliegues con HTTPS.
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = "Lax"
    SESSION_COOKIE_SECURE = os.getenv("SESSION_COOKIE_SECURE", "0").lower() in {"1", "true", "yes", "on"}

    # Duración de sesión explícita para evitar depender de defaults del entorno.
    PERMANENT_SESSION_LIFETIME = 60 * 60 * 8


# Zona horaria institucional para mostrar fechas/horas al usuario.
DISPLAY_TIMEZONE = os.getenv("DISPLAY_TIMEZONE", "America/Bogota")
