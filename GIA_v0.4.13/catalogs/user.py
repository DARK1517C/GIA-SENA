"""
catalogs/user.py

Catálogos oficiales relacionados con los usuarios del sistema GIA.

Este módulo define únicamente valores canónicos utilizados por la
aplicación. Los nombres visibles para el usuario se encuentran en
catalogs/display.py.

No deben añadirse alias ni lógica de negocio en este archivo.
"""

from .common import CatalogEnum


# ==============================================================================
# ROLES DE USUARIO
# ==============================================================================

class UserRole(CatalogEnum):
    """
    Roles oficiales del sistema.

    Los valores son canónicos y no deben modificarse una vez existan
    registros en la base de datos.
    """

    APPRENTICE = "APPRENTICE"

    FOLLOW_UP_INSTRUCTOR = "FOLLOW_UP_INSTRUCTOR"

    LEAD_FOLLOW_UP_INSTRUCTOR = "LEAD_FOLLOW_UP_INSTRUCTOR"

    CERTIFIER = "CERTIFIER"

    CENTER_STAFF = "CENTER_STAFF"

    SUPPORT = "SUPPORT"


# ==============================================================================
# ESTADO DEL USUARIO
# ==============================================================================

class UserStatus(CatalogEnum):
    """
    Estado de la cuenta del usuario.

    ACTIVE:
        Usuario habilitado para ingresar al sistema.

    INACTIVE:
        Usuario deshabilitado administrativamente.

    PENDING:
        Usuario creado pero pendiente de activar.

    SUSPENDED:
        Usuario bloqueado temporalmente.
    """

    ACTIVE = "ACTIVE"

    INACTIVE = "INACTIVE"

    PENDING = "PENDING"

    SUSPENDED = "SUSPENDED"


# ==============================================================================
# TIPO DE IDENTIFICACIÓN
# ==============================================================================

class UserDocumentType(CatalogEnum):
    """
    Tipos de documento oficiales.

    Se utilizan valores descriptivos como identificadores canónicos
    (no abreviaturas) para facilitar el mantenimiento del sistema.
    """

    NATIONAL_ID = "NATIONAL_ID"

    FOREIGNER_ID = "FOREIGNER_ID"

    IDENTITY_CARD = "IDENTITY_CARD"

    PASSPORT = "PASSPORT"

    SPECIAL_STAY_PERMIT = "SPECIAL_STAY_PERMIT"

    TEMPORARY_PROTECTION_PERMIT = "TEMPORARY_PROTECTION_PERMIT"