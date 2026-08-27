"""Autorización canónica de GIA.

Este módulo concentra los roles y permisos puramente basados en rol.
El alcance por registro (por ejemplo, los grupos de un instructor) sigue
siendo responsabilidad de cada módulo porque depende de relaciones de dominio.
"""

from flask_login import current_user

# Roles canónicos
ROLE_APPRENTICE = "APPRENTICE"
ROLE_FOLLOWUP_INSTRUCTOR = "FOLLOW_UP_INSTRUCTOR"
ROLE_LEAD_FOLLOWUP_INSTRUCTOR = "LEAD_FOLLOW_UP_INSTRUCTOR"
ROLE_CENTER_STAFF = "CENTER_STAFF"
ROLE_CERTIFIER = "CERTIFIER"
ROLE_SUPPORT = "SUPPORT"

ROLES = (
    ROLE_APPRENTICE,
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_CENTER_STAFF,
    ROLE_CERTIFIER,
    ROLE_SUPPORT,
)

ROLE_LABELS = {
    ROLE_APPRENTICE: "Aprendiz",
    ROLE_FOLLOWUP_INSTRUCTOR: "Instructor de seguimiento",
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR: "Instructor de seguimiento líder",
    ROLE_CENTER_STAFF: "Administrativo del centro",
    ROLE_CERTIFIER: "Certificador",
    ROLE_SUPPORT: "Soporte",
}

# Visión global de datos administrativos.
GLOBAL_ROLES = frozenset({
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_CENTER_STAFF,
    ROLE_CERTIFIER,
    ROLE_SUPPORT,
})

# Política 3.B: capacidades por rol. El alcance por registro se resuelve
# en services.access_scope. No se debe usar esta tabla para saltarse el alcance.
ROLE_CAPABILITIES = {
    ROLE_APPRENTICE: frozenset({
        "evidences.upload",
    }),
    ROLE_FOLLOWUP_INSTRUCTOR: frozenset({
        "groups.manage", "apprentices.manage", "evidences.manage",
        "evidences.approve", "evidences.upload", "evidences.sign",
        "evidences.activities.manage",
    }),
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR: frozenset({
        "groups.manage", "apprentices.manage", "evidences.manage",
        "evidences.approve", "evidences.upload", "evidences.sign",
        "evidences.activities.manage",
    }),
    ROLE_CENTER_STAFF: frozenset({
        "data.global_view",
    }),
    ROLE_CERTIFIER: frozenset({
        "data.global_view", "evidences.manage", "evidences.approve",
    }),
    ROLE_SUPPORT: frozenset({
        "users.manage", "groups.manage", "apprentices.manage",
        "evidences.manage", "evidences.approve", "evidences.upload",
        "evidences.sign", "evidences.catalog.manage",
        "evidences.activities.manage", "data.global_view",
    }),
}

# Gestión administrativa de aprendices.
APPRENTICE_MANAGEMENT_ROLES = frozenset({
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_SUPPORT,
})

# Gestión institucional de evidencias.
EVIDENCE_MANAGEMENT_ROLES = frozenset({
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_CERTIFIER,
    ROLE_SUPPORT,
})

EVIDENCE_APPROVAL_ROLES = EVIDENCE_MANAGEMENT_ROLES

EVIDENCE_UPLOAD_ROLES = frozenset({
    ROLE_APPRENTICE,
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_SUPPORT,
})

EVIDENCE_SIGNATURE_ROLES = frozenset({
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
})

# Catálogo institucional de evidencias: categorías y plantillas afectan
# globalmente la definición del dominio y, por tanto, quedan reservadas a
# Soporte. Las actividades son operación por ficha y pueden ser gestionadas
# por el instructor responsable, el líder o Soporte.
EVIDENCE_CATALOG_MANAGEMENT_ROLES = frozenset({
    ROLE_SUPPORT,
})

EVIDENCE_ACTIVITY_MANAGEMENT_ROLES = frozenset({
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_SUPPORT,
})

# Administración de usuarios: reservada a Soporte.
USER_ADMIN_ROLES = frozenset({ROLE_SUPPORT})

# Permisos de rol. Los permisos de alcance se evalúan aparte.
PERMISSIONS = {
    "users.manage": USER_ADMIN_ROLES,
    "apprentices.manage": APPRENTICE_MANAGEMENT_ROLES,
    "evidences.manage": EVIDENCE_MANAGEMENT_ROLES,
    "evidences.approve": EVIDENCE_APPROVAL_ROLES,
    "evidences.upload": EVIDENCE_UPLOAD_ROLES,
    "evidences.sign": EVIDENCE_SIGNATURE_ROLES,
    "evidences.catalog.manage": EVIDENCE_CATALOG_MANAGEMENT_ROLES,
    "evidences.activities.manage": EVIDENCE_ACTIVITY_MANAGEMENT_ROLES,
    "data.global_view": GLOBAL_ROLES,
    "groups.manage": frozenset({
        ROLE_FOLLOWUP_INSTRUCTOR, ROLE_LEAD_FOLLOWUP_INSTRUCTOR, ROLE_SUPPORT,
    }),
}


def current_role() -> str | None:
    """Devuelve el rol canónico del usuario autenticado."""
    return getattr(current_user, "role", None)


def has_role(role: str) -> bool:
    return current_role() == role


def has_any_role(*roles: str) -> bool:
    return current_role() in roles


def has_permission(permission: str) -> bool:
    """Comprueba únicamente el componente de autorización basado en rol."""
    return current_role() in PERMISSIONS.get(permission, frozenset())
