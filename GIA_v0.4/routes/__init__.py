from .auth import auth_bp
from .apprentices import apprentices_bp
from .groups import groups_bp
from .dashboard import dashboard_bp

# importaciones opcionales que pueden fallar si el módulo no existe;
# usamos try/except para evitar que un módulo faltante rompa toda la importación
try:
    from .bitacoras import bitacoras_bp
except Exception:
    bitacoras_bp = None

try:
    from .users import users_bp
except Exception:
    users_bp = None

__all__ = [
    "auth_bp",
    "apprentices_bp",
    "groups_bp",
    "dashboard_bp",
    "bitacoras_bp",
    "users_bp",
]