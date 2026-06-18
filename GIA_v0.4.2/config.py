import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

class Config:
    SECRET_KEY = os.getenv("SECRET_KEY", "gia-sena-secret")
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {"pool_pre_ping": True}
    UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")

    # DATABASE_URL precedence: DATABASE_URL -> MYSQL_URL -> sqlite file
    SQLALCHEMY_DATABASE_URI = os.getenv(
        "DATABASE_URL",
        os.getenv("MYSQL_URL", f"sqlite:///{os.path.join(BASE_DIR, 'gia.db')}")
    )

    # SQLite specific connect args will be applied in extensions when needed
