"""
app/catalogs/display.py
~~~~~~~~~~~~~~~~~~~~~~~

Etiquetas oficiales (labels) utilizadas para representar los valores
canónicos de los catálogos en la interfaz de usuario (UX/UI).

Este módulo constituye la única fuente oficial de textos visibles para
los catálogos del sistema.

IMPORTANTE
----------
• No contiene lógica de negocio.
• No normaliza datos.
• No realiza validaciones.
• No utiliza alias.

Los valores canónicos provienen de los CatalogEnum definidos en
los diferentes módulos de catalogs.

La normalización se mantiene en los módulos de importación/validación correspondientes.

La validación pertenece a:

    validation.py

El registro pertenece a:

    registry.py

Autor
------
Proyecto GIA
"""

from __future__ import annotations

# ==============================================================================
# COMMON CATALOGS
# ==============================================================================

from .common_catalogs import (
    ProgramLevel,
    Gender,
    DocumentType,
    YesNo,
    RecordStatus,
)

# ==============================================================================
# APPRENTICE CATALOGS
# ==============================================================================

from .apprentice import (
    SofiaStatus,
    EpModality,
    EmploymentRelationship,
    TrainingRelationship,
)

# ==============================================================================
# TRAINING GROUP CATALOGS
# ==============================================================================

from .training_group import (
    GroupModality,
    GroupStatus,
    GroupMunicipality,
)

# ==============================================================================
# USER CATALOGS
# ==============================================================================

from .user import (
    UserRole,
    UserStatus,
    UserDocumentType,
)

# ==============================================================================
# PROGRAM LEVEL
# ==============================================================================

PROGRAM_LEVEL_LABELS = {

    ProgramLevel.OPERARIO:
        "Operario",

    ProgramLevel.AUXILIAR:
        "Auxiliar",

    ProgramLevel.TECNICO:
        "Técnico",

    ProgramLevel.TECNOLOGO:
        "Tecnólogo",

}

# ==============================================================================
# GENDER
# ==============================================================================

GENDER_LABELS = {

    Gender.MASCULINO:
        "Masculino",

    Gender.FEMENINO:
        "Femenino",

}

# ==============================================================================
# DOCUMENT TYPE
# ==============================================================================

DOCUMENT_TYPE_LABELS = {

    DocumentType.CEDULA_CIUDADANIA:
        "Cédula de ciudadanía",

    DocumentType.TARJETA_IDENTIDAD:
        "Tarjeta de identidad",

    DocumentType.CEDULA_EXTRANJERIA:
        "Cédula de extranjería",

    DocumentType.PERMISO_PROTECCION_TEMPORAL:
        "Permiso por Protección Temporal (PPT)",

    DocumentType.PERMISO_ESPECIAL_PERMANENCIA:
        "Permiso Especial de Permanencia (PEP)",

}

# ==============================================================================
# YES / NO
# ==============================================================================

YES_NO_LABELS = {

    YesNo.SI:
        "Sí",

    YesNo.NO:
        "No",

}

# ==============================================================================
# RECORD STATUS
# ==============================================================================

RECORD_STATUS_LABELS = {

    RecordStatus.ACTIVO:
        "Activo",

    RecordStatus.INACTIVO:
        "Inactivo",

}

# ==============================================================================
# SOFIA STATUS
# ==============================================================================

SOFIA_STATUS_LABELS = {

    SofiaStatus.EN_FORMACION:
        "En formación",

    SofiaStatus.CERTIFICADO:
        "Certificado",

    SofiaStatus.POR_CANCELAR:
        "Por cancelar",

    SofiaStatus.CANCELADO:
        "Cancelado",

}

# ==============================================================================
# ETAPA PRODUCTIVA
# ==============================================================================

EP_MODALITY_LABELS = {

    EpModality.CONTRATO_APRENDIZAJE:
        "Contrato de aprendizaje",

    EpModality.CONTRATO_VINCULO_FORMATIVO:
        "Contrato de vínculo formativo",

    EpModality.VINCULO_LABORAL:
        "Vínculo laboral",

    EpModality.PROYECTO_PRODUCTIVO:
        "Proyecto productivo",

    EpModality.MONITORIA:
        "Monitoría",

    EpModality.PRACTICAS_ECONOMIA_POPULAR:
        "Prácticas en la economía popular y/o campesina",

}


# ==============================================================================
# GROUP MODALITY
# ==============================================================================

GROUP_MODALITY_LABELS = {

    GroupModality.PRESENCIAL:
        "Presencial",

    GroupModality.VIRTUAL:
        "Virtual",

    GroupModality.DISTANCIA: "A distancia",
    GroupModality.DUAL: "Dual",

}

# ==============================================================================
# GROUP STATUS
# ==============================================================================

GROUP_STATUS_LABELS = {

    GroupStatus.EN_EJECUCION:
        "En ejecución",

    GroupStatus.CANCELADA:
        "Cancelada",

    GroupStatus.TERMINADO_POR_FECHA:
        "Terminado por fecha",

}

# ==============================================================================
# GROUP MUNICIPALITY
# ==============================================================================

GROUP_MUNICIPALITY_LABELS = {

    municipality: municipality.value.replace("_", " ").title()
    for municipality in GroupMunicipality

}

# ==============================================================================
# USER ROLE
# ==============================================================================

USER_ROLE_LABELS = {

    UserRole.APPRENTICE:
        "Aprendiz",

    UserRole.FOLLOW_UP_INSTRUCTOR:
        "Instructor de seguimiento",

    UserRole.LEAD_FOLLOW_UP_INSTRUCTOR:
        "Instructor de seguimiento líder",

    UserRole.CERTIFIER:
        "Certificador",

    UserRole.CENTER_STAFF:
        "Administrativo del centro/complejo",

    UserRole.SUPPORT:
        "Soporte",

}

# ==============================================================================
# USER STATUS
# ==============================================================================

USER_STATUS_LABELS = {

    UserStatus.ACTIVE:
        "Activo",

    UserStatus.INACTIVE:
        "Inactivo",

    UserStatus.PENDING:
        "Pendiente de activación",

    UserStatus.SUSPENDED:
        "Suspendido",

}

# ==============================================================================
# USER DOCUMENT TYPE
# ==============================================================================

USER_DOCUMENT_TYPE_LABELS = {

    UserDocumentType.NATIONAL_ID:
        "Cédula de ciudadanía",

    UserDocumentType.IDENTITY_CARD:
        "Tarjeta de identidad",

    UserDocumentType.FOREIGNER_ID:
        "Cédula de extranjería",

    UserDocumentType.TEMPORARY_PROTECTION_PERMIT:
        "Permiso por Protección Temporal (PPT)",

    UserDocumentType.SPECIAL_STAY_PERMIT:
        "Permiso Especial de Permanencia (PEP)",

    UserDocumentType.PASSPORT:
        "Pasaporte",

}

# ==============================================================================
# LABEL REGISTRY
# ==============================================================================

CATALOG_LABELS = {

    # --------------------------------------------------------------------------
    # Common catalogs
    # --------------------------------------------------------------------------

    ProgramLevel:
        PROGRAM_LEVEL_LABELS,

    Gender:
        GENDER_LABELS,

    DocumentType:
        DOCUMENT_TYPE_LABELS,

    YesNo:
        YES_NO_LABELS,

    RecordStatus:
        RECORD_STATUS_LABELS,

    # --------------------------------------------------------------------------
    # Apprentice
    # --------------------------------------------------------------------------

    SofiaStatus:
        SOFIA_STATUS_LABELS,

    EpModality:
        EP_MODALITY_LABELS,

    # --------------------------------------------------------------------------
    # Training Group
    # --------------------------------------------------------------------------

    GroupModality:
        GROUP_MODALITY_LABELS,

    GroupStatus:
        GROUP_STATUS_LABELS,

    GroupMunicipality:
        GROUP_MUNICIPALITY_LABELS,

    # --------------------------------------------------------------------------
    # User
    # --------------------------------------------------------------------------

    UserRole:
        USER_ROLE_LABELS,

    UserStatus:
        USER_STATUS_LABELS,

    UserDocumentType:
        USER_DOCUMENT_TYPE_LABELS,

}

# ==============================================================================
# LABEL MAPS
# ==============================================================================

LABEL_MAPS = {

    # --------------------------------------------------------------------------
    # Common
    # --------------------------------------------------------------------------

    "program_level":
        PROGRAM_LEVEL_LABELS,

    "gender":
        GENDER_LABELS,

    "document_type":
        DOCUMENT_TYPE_LABELS,

    "yes_no":
        YES_NO_LABELS,

    "record_status":
        RECORD_STATUS_LABELS,

    # --------------------------------------------------------------------------
    # Apprentice
    # --------------------------------------------------------------------------

    "sofia_status":
        SOFIA_STATUS_LABELS,

    "ep_modality":
        EP_MODALITY_LABELS,


    # --------------------------------------------------------------------------
    # Training Group
    # --------------------------------------------------------------------------

    "group_modality":
        GROUP_MODALITY_LABELS,

    "group_status":
        GROUP_STATUS_LABELS,

    "group_municipality":
        GROUP_MUNICIPALITY_LABELS,

    # --------------------------------------------------------------------------
    # User
    # --------------------------------------------------------------------------

    "user_role":
        USER_ROLE_LABELS,

    "user_status":
        USER_STATUS_LABELS,

    "user_document_type":
        USER_DOCUMENT_TYPE_LABELS,

}

# ==============================================================================
# HELPERS
# ==============================================================================

def get_label(catalog, value) -> str:
    """
    Devuelve la etiqueta (label) asociada a un valor canónico.

    Parameters
    ----------
    catalog : type
        Clase CatalogEnum registrada en CATALOG_LABELS.

    value
        Valor canónico del catálogo.

    Returns
    -------
    str
        Texto visible para el usuario.

    Raises
    ------
    KeyError
        Si el catálogo no está registrado.
    """

    mapping = CATALOG_LABELS.get(catalog)

    if mapping is None:
        raise KeyError(
            f"El catálogo '{catalog.__name__}' no está registrado "
            "en CATALOG_LABELS."
        )

    if value is None:
        return ""

    return mapping.get(value, str(value))


# ------------------------------------------------------------------------------


def get_catalog_label(catalog_name: str, value) -> str:
    """
    Obtiene la etiqueta utilizando el nombre registrado
    en LABEL_MAPS.

    Parameters
    ----------
    catalog_name : str

    value

    Returns
    -------
    str
    """

    mapping = LABEL_MAPS.get(catalog_name)

    if mapping is None:
        raise KeyError(
            f"No existe un catálogo DISPLAY llamado "
            f"'{catalog_name}'."
        )

    if value is None:
        return ""

    return mapping.get(value, str(value))


# ------------------------------------------------------------------------------


def get_label_list(catalog, values):
    """
    Convierte una colección de valores canónicos
    en una lista de etiquetas.

    Parameters
    ----------
    catalog : type

    values : Iterable

    Returns
    -------
    list[str]
    """

    if not values:
        return []

    return [
        get_label(catalog, value)
        for value in values
    ]


# ------------------------------------------------------------------------------


def get_choices(catalog):
    """
    Convierte un catálogo registrado en una lista de opciones.

    Muy útil para:

    - Flask-WTF
    - Select2
    - Combobox
    - APIs
    """

    mapping = CATALOG_LABELS.get(catalog)

    if mapping is None:
        raise KeyError(
            f"El catálogo '{catalog.__name__}' no está registrado."
        )

    return list(mapping.items())

# ------------------------------------------------------------------------------

def get_choices_sorted(catalog):
    """
    Devuelve las opciones de un catálogo ordenadas
    alfabéticamente por su etiqueta.

    Muy útil para:

    - Formularios
    - Select2
    - Filtros
    - Combobox
    """

    choices = get_choices(catalog)

    return sorted(
        choices,
        key=lambda item: item[1]
    )

# ------------------------------------------------------------------------------


def has_catalog(catalog) -> bool:
    """
    Indica si un CatalogEnum está registrado.
    """

    return catalog in CATALOG_LABELS


# ------------------------------------------------------------------------------


def has_catalog_name(catalog_name: str) -> bool:
    """
    Indica si un nombre de catálogo existe
    dentro de LABEL_MAPS.
    """

    return catalog_name in LABEL_MAPS


# ==============================================================================
# EXPORTS
# ==============================================================================

__all__ = [

    # --------------------------------------------------------------------------
    # Diccionarios
    # --------------------------------------------------------------------------

    "PROGRAM_LEVEL_LABELS",
    "GENDER_LABELS",
    "DOCUMENT_TYPE_LABELS",
    "YES_NO_LABELS",
    "RECORD_STATUS_LABELS",

    "SOFIA_STATUS_LABELS",
    "EP_MODALITY_LABELS",

    "GROUP_MODALITY_LABELS",
    "GROUP_STATUS_LABELS",
    "GROUP_MUNICIPALITY_LABELS",

    "USER_ROLE_LABELS",
    "USER_STATUS_LABELS",
    "USER_DOCUMENT_TYPE_LABELS",

    # --------------------------------------------------------------------------
    # Registros
    # --------------------------------------------------------------------------

    "CATALOG_LABELS",
    "LABEL_MAPS",

    # --------------------------------------------------------------------------
    # Helpers
    # --------------------------------------------------------------------------

    "get_label",
    "get_catalog_label",
    "get_label_list",
    "get_choices",
    "get_choices_sorted",
    "has_catalog",
    "has_catalog_name",

]