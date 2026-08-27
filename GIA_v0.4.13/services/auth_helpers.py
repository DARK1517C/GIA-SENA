from functools import wraps
from flask import abort
from flask_login import current_user
from extensions import login_manager
from models import User
from services.permissions import has_permission

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def role_required(*roles):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if not current_user.is_authenticated:
                return login_manager.unauthorized()
            if current_user.role not in roles:
                abort(403)
            return view(*args, **kwargs)
        return wrapped
    return decorator


def permission_required(permission):
    """Protege una vista mediante un permiso canónico basado en rol."""
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if not current_user.is_authenticated:
                return login_manager.unauthorized()
            if not has_permission(permission):
                abort(403)
            return view(*args, **kwargs)
        return wrapped
    return decorator
