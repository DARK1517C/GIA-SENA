"""
catalogs/aliases.py
~~~~~~~~~~~~~~~~~~~

Normalización oficial de valores para todos los catálogos del sistema.

Este módulo convierte entradas provenientes de:

- Excel
- Formularios
- APIs
- Scripts
- Migraciones

en los valores canónicos definidos mediante CatalogEnum.

IMPORTANTE
----------
Este módulo NO muestra información al usuario.

Para ello existe:

    display.py
"""

from __future__ import annotations

import re
import unicodedata

from types import MappingProxyType

from .common_catalogs import (
    ProgramLevel,
    Gender,
    DocumentType,
    YesNo,
    RecordStatus,
)

from .apprentice import (
    SofiaStatus,
    EpModality,
    EmploymentRelationship,
    TrainingRelationship,
)

from .training_group import (
    GroupModality,
    GroupStatus,
    GroupMunicipality,
)

from .user import (
    UserRole,
    UserStatus,
    UserDocumentType,
)


# ==============================================================================
# INTERNAL HELPERS
# ==============================================================================

def _freeze(mapping: dict):
    """
    Convierte un diccionario en una estructura inmutable.
    """
    return MappingProxyType(mapping)


def normalize_alias(value: str | None) -> str | None:
    """
    Normaliza una cadena antes de buscarla
    dentro de un catálogo de alias.

    Operaciones realizadas:

    - strip()
    - mayúsculas
    - elimina tildes
    - elimina puntos
    - elimina comas
    - elimina guiones
    - elimina dobles espacios
    """

    if value is None:
        return None

    value = str(value).strip().upper()

    value = unicodedata.normalize("NFD", value)

    value = "".join(
        c
        for c in value
        if unicodedata.category(c) != "Mn"
    )

    value = re.sub(r"[.,]", "", value)

    value = value.replace("-", " ")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


# ==============================================================================
# COMMON
# ==============================================================================

DOCUMENT_TYPE_ALIASES = _freeze({

    "CC":
        DocumentType.CEDULA_CIUDADANIA,

    "CEDULA":
        DocumentType.CEDULA_CIUDADANIA,

    "CEDULA DE CIUDADANIA":
        DocumentType.CEDULA_CIUDADANIA,

    "TI":
        DocumentType.TARJETA_IDENTIDAD,

    "TARJETA DE IDENTIDAD":
        DocumentType.TARJETA_IDENTIDAD,

    "CE":
        DocumentType.CEDULA_EXTRANJERIA,

    "CEDULA DE EXTRANJERIA":
        DocumentType.CEDULA_EXTRANJERIA,

    "PPT":
        DocumentType.PERMISO_PROTECCION_TEMPORAL,

    "PERMISO PROTECCION TEMPORAL":
        DocumentType.PERMISO_PROTECCION_TEMPORAL,

    "PEP":
        DocumentType.PERMISO_ESPECIAL_PERMANENCIA,

    "PERMISO ESPECIAL PERMANENCIA":
        DocumentType.PERMISO_ESPECIAL_PERMANENCIA,

})


GENDER_ALIASES = _freeze({

    "M":
        Gender.MASCULINO,

    "MASCULINO":
        Gender.MASCULINO,

    "F":
        Gender.FEMENINO,

    "FEMENINO":
        Gender.FEMENINO,

})


YES_NO_ALIASES = _freeze({

    "SI":
        YesNo.SI,

    "S":
        YesNo.SI,

    "NO":
        YesNo.NO,

    "N":
        YesNo.NO,

})


RECORD_STATUS_ALIASES = _freeze({

    "ACTIVO":
        RecordStatus.ACTIVO,

    "INACTIVO":
        RecordStatus.INACTIVO,

})


PROGRAM_LEVEL_ALIASES = _freeze({

    "OPERARIO":
        ProgramLevel.OPERARIO,

    "AUXILIAR":
        ProgramLevel.AUXILIAR,

    "TECNICO":
        ProgramLevel.TECNICO,

    "TECNOLOGO":
        ProgramLevel.TECNOLOGO,

})


# ==============================================================================
# APPRENTICE
# ==============================================================================

SOFIA_STATUS_ALIASES = _freeze({

    "EN FORMACION":
        SofiaStatus.EN_FORMACION,

    "CERTIFICADO":
        SofiaStatus.CERTIFICADO,

    "POR CANCELAR":
        SofiaStatus.POR_CANCELAR,

    "CANCELADO":
        SofiaStatus.CANCELADO,

})


EP_MODALITY_ALIASES = _freeze({

    "CONTRATO APRENDIZAJE":
        EpModality.CONTRATO_APRENDIZAJE,

    "CONTRATO VINCULO FORMATIVO":
        EpModality.CONTRATO_VINCULO_FORMATIVO,

    "VINCULO LABORAL":
        EpModality.VINCULO_LABORAL,

    "PROYECTO PRODUCTIVO":
        EpModality.PROYECTO_PRODUCTIVO,

    "MONITORIA":
        EpModality.MONITORIA,

    "PRACTICAS ECONOMIA POPULAR":
        EpModality.PRACTICAS_ECONOMIA_POPULAR,

})

# ==============================================================================
# EMPLOYMENT RELATIONSHIP
# ==============================================================================

EMPLOYMENT_RELATIONSHIP_ALIASES = _freeze({

    "TERMINO INDEFINIDO":
        EmploymentRelationship.TERMINO_INDEFINIDO,

    "INDEFINIDO":
        EmploymentRelationship.TERMINO_INDEFINIDO,

    "TERMINO FIJO":
        EmploymentRelationship.TERMINO_FIJO,

    "FIJO":
        EmploymentRelationship.TERMINO_FIJO,

    "OBRA LABOR":
        EmploymentRelationship.OBRA_LABOR,

    "OBRA O LABOR":
        EmploymentRelationship.OBRA_LABOR,

    "TEMPORAL":
        EmploymentRelationship.TEMPORAL,

    "PRESTACION SERVICIOS":
        EmploymentRelationship.PRESTACION_SERVICIOS,

    "PRESTACION DE SERVICIOS":
        EmploymentRelationship.PRESTACION_SERVICIOS,

    "OTRO":
        EmploymentRelationship.OTRO,

})


# ==============================================================================
# TRAINING RELATIONSHIP
# ==============================================================================

TRAINING_RELATIONSHIP_ALIASES = _freeze({

    "CONVENIO SENA EMPRESA":
        TrainingRelationship.CONVENIO_SENA_EMPRESA,

    "CONVENIO":
        TrainingRelationship.CONVENIO_SENA_EMPRESA,

    "PRACTICA EMPRESARIAL":
        TrainingRelationship.PRACTICA_EMPRESARIAL,

    "PROYECTO PRODUCTIVO":
        TrainingRelationship.PROYECTO_PRODUCTIVO,

    "MONITORIA":
        TrainingRelationship.MONITORIA,

    "OTRO":
        TrainingRelationship.OTRO,

})


# ==============================================================================
# TRAINING GROUP
# ==============================================================================

GROUP_MODALITY_ALIASES = _freeze({

    "PRESENCIAL":
        GroupModality.PRESENCIAL,

    "VIRTUAL":
        GroupModality.VIRTUAL,

    "DISTANCIA":
        GroupModality.DISTANCIA,

    "A DISTANCIA":
        GroupModality.DISTANCIA,

    "DUAL":
        GroupModality.DUAL,

})


GROUP_STATUS_ALIASES = _freeze({

    "EN EJECUCION":
        GroupStatus.EN_EJECUCION,

    "TERMINADO POR FECHA":
        GroupStatus.TERMINADO_POR_FECHA,

})


# ==============================================================================
# GROUP MUNICIPALITY
# ==============================================================================

GROUP_MUNICIPALITY_ALIASES = _freeze({

    normalize_alias(municipality.value): municipality
    for municipality in GroupMunicipality

})


# ==============================================================================
# USERS
# ==============================================================================

USER_ROLE_ALIASES = _freeze({

    # --------------------------------------------------------------------------
    # Aprendiz
    # --------------------------------------------------------------------------

    "APRENDIZ":
        UserRole.APPRENTICE,

    # --------------------------------------------------------------------------
    # Instructor de seguimiento
    # --------------------------------------------------------------------------

    "INSTRUCTOR":
        UserRole.FOLLOW_UP_INSTRUCTOR,

    "DOCENTE":
        UserRole.FOLLOW_UP_INSTRUCTOR,

    "INSTRUCTOR SEGUIMIENTO":
        UserRole.FOLLOW_UP_INSTRUCTOR,

    "INSTRUCTOR DE SEGUIMIENTO":
        UserRole.FOLLOW_UP_INSTRUCTOR,

    # --------------------------------------------------------------------------
    # Instructor líder
    # --------------------------------------------------------------------------

    "INSTRUCTOR LIDER":
        UserRole.LEAD_FOLLOW_UP_INSTRUCTOR,

    "INSTRUCTOR DE SEGUIMIENTO LIDER":
        UserRole.LEAD_FOLLOW_UP_INSTRUCTOR,

    "LIDER":
        UserRole.LEAD_FOLLOW_UP_INSTRUCTOR,

    # --------------------------------------------------------------------------
    # Certificador
    # --------------------------------------------------------------------------

    "CERTIFICADOR":
        UserRole.CERTIFIER,

    # --------------------------------------------------------------------------
    # Administrativo
    # --------------------------------------------------------------------------

    "ADMINISTRATIVO":
        UserRole.CENTER_STAFF,

    "ADMINISTRATIVO DEL CENTRO":
        UserRole.CENTER_STAFF,

    "ADMINISTRATIVO CENTRO":
        UserRole.CENTER_STAFF,

    # --------------------------------------------------------------------------
    # Soporte
    # --------------------------------------------------------------------------

    "SOPORTE":
        UserRole.SUPPORT,

    "SUPER ADMINISTRADOR":
        UserRole.SUPPORT,

    "SUPERADMINISTRADOR":
        UserRole.SUPPORT,

})


USER_STATUS_ALIASES = _freeze({

    "ACTIVO":
        UserStatus.ACTIVE,

    "INACTIVO":
        UserStatus.INACTIVE,

    "PENDIENTE":
        UserStatus.PENDING,

    "PENDIENTE DE ACTIVACION":
        UserStatus.PENDING,

    "SUSPENDIDO":
        UserStatus.SUSPENDED,

})


# ==============================================================================
# USER DOCUMENT TYPE
# ==============================================================================

USER_DOCUMENT_TYPE_ALIASES = DOCUMENT_TYPE_ALIASES


# ==============================================================================
# REGISTRO OFICIAL
# ==============================================================================

CATALOG_ALIASES = _freeze({

    # --------------------------------------------------------------------------
    # Common
    # --------------------------------------------------------------------------

    ProgramLevel:
        PROGRAM_LEVEL_ALIASES,

    Gender:
        GENDER_ALIASES,

    DocumentType:
        DOCUMENT_TYPE_ALIASES,

    YesNo:
        YES_NO_ALIASES,

    RecordStatus:
        RECORD_STATUS_ALIASES,

    # --------------------------------------------------------------------------
    # Apprentice
    # --------------------------------------------------------------------------

    SofiaStatus:
        SOFIA_STATUS_ALIASES,

    EpModality:
        EP_MODALITY_ALIASES,

    EmploymentRelationship:
        EMPLOYMENT_RELATIONSHIP_ALIASES,

    TrainingRelationship:
        TRAINING_RELATIONSHIP_ALIASES,

    # --------------------------------------------------------------------------
    # Training Group
    # --------------------------------------------------------------------------

    GroupModality:
        GROUP_MODALITY_ALIASES,

    GroupStatus:
        GROUP_STATUS_ALIASES,

    GroupMunicipality:
        GROUP_MUNICIPALITY_ALIASES,

    # --------------------------------------------------------------------------
    # Users
    # --------------------------------------------------------------------------

    UserRole:
        USER_ROLE_ALIASES,

    UserStatus:
        USER_STATUS_ALIASES,

    UserDocumentType:
        USER_DOCUMENT_TYPE_ALIASES,

})

# ==============================================================================
# REVERSE ALIASES
# ==============================================================================

def _build_reverse_aliases():
    """
    Construye automáticamente el registro inverso de alias.

    Resultado:

        Enum -> [alias1, alias2, alias3]
    """

    reverse = {}

    for catalog_aliases in CATALOG_ALIASES.values():

        for alias, canonical in catalog_aliases.items():

            reverse.setdefault(canonical, []).append(alias)

    return {
        key: tuple(sorted(values))
        for key, values in reverse.items()
    }


REVERSE_CATALOG_ALIASES = _freeze(_build_reverse_aliases())


# ==============================================================================
# HELPERS
# ==============================================================================

def normalize(catalog, value):
    """
    Normaliza un valor utilizando los alias registrados.

    Si el valor ya pertenece al CatalogEnum indicado,
    se devuelve sin modificaciones.

    Parameters
    ----------
    catalog
        Clase CatalogEnum.

    value
        Valor a normalizar.

    Returns
    -------
    CatalogEnum | None
    """

    if value is None:
        return None

    if isinstance(value, catalog):
        return value

    aliases = CATALOG_ALIASES.get(catalog)

    if aliases is None:
        raise KeyError(
            f"El catálogo '{catalog.__name__}' "
            "no está registrado."
        )

    key = normalize_alias(value)

    return aliases.get(key)


# ------------------------------------------------------------------------------


def normalize_or_none(catalog, value):
    """
    Igual que normalize().

    Se incluye por claridad semántica cuando el código
    quiera expresar explícitamente que None es aceptable.
    """

    return normalize(catalog, value)


# ------------------------------------------------------------------------------


def normalize_or_raise(catalog, value):
    """
    Normaliza un valor.

    Lanza ValueError cuando el alias no existe.
    """

    normalized = normalize(catalog, value)

    if normalized is None:

        raise ValueError(

            f"'{value}' no es un valor válido para "
            f"{catalog.__name__}."

        )

    return normalized


# ------------------------------------------------------------------------------


def is_valid_alias(catalog, value):
    """
    Indica si un valor puede normalizarse.
    """

    return normalize(catalog, value) is not None


# ------------------------------------------------------------------------------


def get_aliases(catalog):
    """
    Devuelve el diccionario oficial de alias
    de un catálogo.
    """

    aliases = CATALOG_ALIASES.get(catalog)

    if aliases is None:

        raise KeyError(

            f"El catálogo '{catalog.__name__}' "
            "no está registrado."

        )

    return aliases


# ------------------------------------------------------------------------------


def get_reverse_aliases(catalog_value):
    """
    Devuelve todos los alias asociados
    a un valor canónico.
    """

    return REVERSE_CATALOG_ALIASES.get(
        catalog_value,
        ()
    )


# ------------------------------------------------------------------------------


def has_catalog(catalog):
    """
    Indica si un catálogo tiene alias registrados.
    """

    return catalog in CATALOG_ALIASES


# ==============================================================================
# EXPORTS
# ==============================================================================

__all__ = [

    # --------------------------------------------------------------------------
    # Diccionarios
    # --------------------------------------------------------------------------

    "PROGRAM_LEVEL_ALIASES",
    "GENDER_ALIASES",
    "DOCUMENT_TYPE_ALIASES",
    "YES_NO_ALIASES",
    "RECORD_STATUS_ALIASES",

    "SOFIA_STATUS_ALIASES",
    "EP_MODALITY_ALIASES",
    "EMPLOYMENT_RELATIONSHIP_ALIASES",
    "TRAINING_RELATIONSHIP_ALIASES",

    "GROUP_MODALITY_ALIASES",
    "GROUP_STATUS_ALIASES",
    "GROUP_MUNICIPALITY_ALIASES",

    "USER_ROLE_ALIASES",
    "USER_STATUS_ALIASES",
    "USER_DOCUMENT_TYPE_ALIASES",

    # --------------------------------------------------------------------------
    # Registros
    # --------------------------------------------------------------------------

    "CATALOG_ALIASES",
    "REVERSE_CATALOG_ALIASES",

    # --------------------------------------------------------------------------
    # Helpers
    # --------------------------------------------------------------------------

    "normalize_alias",

    "normalize",

    "normalize_or_none",

    "normalize_or_raise",

    "is_valid_alias",

    "get_aliases",

    "get_reverse_aliases",

    "has_catalog",

]