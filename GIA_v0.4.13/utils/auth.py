"""Compatibilidad legacy para autorización.

La política canónica vive en services.permissions. Este módulo no mantiene
roles históricos ni una matriz paralela.
"""
from functools import wraps
from flask import abort
from flask_login import current_user
from services.permissions import has_any_role


def role_required(allowed_roles):
    if isinstance(allowed_roles, str):
        allowed = (allowed_roles,)
    else:
        allowed = tuple(allowed_roles)

    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if not current_user.is_authenticated:
                abort(401)
            if not has_any_role(*allowed):
                abort(403)
            return view(*args, **kwargs)
        return wrapped
    return decorator
