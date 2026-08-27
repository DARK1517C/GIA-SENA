"""
catalogs/apprentice.py

Catálogos exclusivos del dominio Apprentice.

Los catálogos compartidos del sistema, como ProgramLevel,
Gender y DocumentType, se encuentran en common_catalogs.py.
"""

from __future__ import annotations

from .common import CatalogEnum


# ==============================================================================
# ESTADO DEL APRENDIZ EN SOFIA
# ==============================================================================

class SofiaStatus(CatalogEnum):
    """
    Estado académico del aprendiz en SOFIA Plus.
    """

    EN_FORMACION = "EN_FORMACION"
    CERTIFICADO = "CERTIFICADO"
    POR_CANCELAR = "POR_CANCELAR"
    CANCELADO = "CANCELADO"


# ==============================================================================
# MODALIDAD DE ETAPA PRODUCTIVA
# ==============================================================================

class EpModality(CatalogEnum):
    """
    Modalidades de Etapa Productiva utilizadas por GIA.
    """

    CONTRATO_APRENDIZAJE = "CONTRATO_APRENDIZAJE"
    CONTRATO_VINCULO_FORMATIVO = "CONTRATO_VINCULO_FORMATIVO"
    VINCULO_LABORAL = "VINCULO_LABORAL"
    PROYECTO_PRODUCTIVO = "PROYECTO_PRODUCTIVO"
    MONITORIA = "MONITORIA"
    PRACTICAS_ECONOMIA_POPULAR = "PRACTICAS_ECONOMIA_POPULAR"


# ==============================================================================
# RELACIÓN LABORAL
# ==============================================================================

class EmploymentRelationship(CatalogEnum):
    """
    Tipo de relación laboral del aprendiz durante la etapa productiva.
    """

    TERMINO_INDEFINIDO = "TERMINO_INDEFINIDO"
    TERMINO_FIJO = "TERMINO_FIJO"
    OBRA_LABOR = "OBRA_LABOR"
    TEMPORAL = "TEMPORAL"
    PRESTACION_SERVICIOS = "PRESTACION_SERVICIOS"
    OTRO = "OTRO"


# ==============================================================================
# RELACIÓN CON LA FORMACIÓN
# ==============================================================================

class TrainingRelationship(CatalogEnum):
    """
    Relación mediante la cual se desarrolla la etapa productiva.
    """

    CONVENIO_SENA_EMPRESA = "CONVENIO_SENA_EMPRESA"
    PRACTICA_EMPRESARIAL = "PRACTICA_EMPRESARIAL"
    PROYECTO_PRODUCTIVO = "PROYECTO_PRODUCTIVO"
    MONITORIA = "MONITORIA"
    OTRO = "OTRO"


# ==============================================================================
# GESTIÓN INDIVIDUAL
# ==============================================================================

class IndividualManagement(CatalogEnum):
    """
    Indica si el aprendiz requiere seguimiento
    mediante gestión individual.
    """

    SI = "SI"
    NO = "NO"



# ==============================================================================
# EXPORTACIONES
# ==============================================================================

__all__ = [
    "SofiaStatus",
    "EpModality",
    "EmploymentRelationship",
    "TrainingRelationship",
    "IndividualManagement",
]