"""
Catálogos oficiales de GIA.

Este paquete contiene todos los valores canónicos del sistema.

Reglas:

1. La Base de Datos SOLO debe almacenar valores canónicos.
2. Todo dato importado debe normalizarse antes de guardarse.
3. La interfaz utiliza etiquetas amigables (labels).
4. Nunca escribir cadenas de texto ("magic strings") directamente en el código.

Autor:
Proyecto GIA
"""

from .common import CatalogEnum

from .common import (
    CatalogEnum,
    normalize_text,
    normalize_spaces,
    remove_accents,
    clean_string,
    is_empty,
    unique,
)

from .common_catalogs import (
    ProgramLevel,
    Gender,
    DocumentType,
    YesNo,
    RecordStatus,
    TrainingModality,
)

from .apprentice import (
    SofiaStatus,
    EpModality,
    IndividualManagement,
)

# Training Groups
from .training_group import (
    GroupModality,
    GroupStatus,
    GroupMunicipality,
)

__all__ = [
    # Base
    "CatalogEnum",
    "normalize_text",
    "normalize_spaces",
    "remove_accents",
    "clean_string",
    "is_empty",
    "unique",

    # Catálogos compartidos
    "ProgramLevel",
    "Gender",
    "DocumentType",
    "YesNo",
    "RecordStatus",
    "TrainingModality",

    # Catálogos de Apprentice
    "SofiaStatus",
    "EpModality",
    "IndividualManagement",
]
