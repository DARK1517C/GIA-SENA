# utils/auth.py
from functools import wraps
from flask import abort
from flask_login import current_user

def role_required(allowed_roles):
    """
    Decorador para restringir acceso por rol.
    allowed_roles: str o lista/tupla de roles permitidos.
    Ejemplo: @role_required(['docente','super_admin'])
    """
    if isinstance(allowed_roles, str):
        allowed = {allowed_roles}
    else:
        allowed = set(allowed_roles)

    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if not current_user or not getattr(current_user, "is_authenticated", False):
                abort(401)  # no autenticado
            if getattr(current_user, "role", None) not in allowed:
                abort(403)  # prohibido
            return f(*args, **kwargs)
        return wrapped
    return decorator
