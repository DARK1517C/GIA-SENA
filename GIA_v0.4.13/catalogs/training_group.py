"""
app/catalogs/training_group.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Catálogos exclusivos del dominio Training Group.

IMPORTANTE

Este módulo contiene únicamente conceptos propios
de los grupos de formación.

Los catálogos compartidos como:

    ProgramLevel
    Gender
    DocumentType
    YesNo

se encuentran en:

    app.catalogs.common_catalogs

Autor:
Proyecto GIA
"""

from __future__ import annotations

from .common import CatalogEnum


# ==========================================================
# MODALIDAD DE FORMACIÓN
# ==========================================================

class GroupModality(CatalogEnum):
    """
    Modalidad de ejecución del grupo.
    """

    PRESENCIAL = "PRESENCIAL"

    VIRTUAL = "VIRTUAL"

    DISTANCIA = "DISTANCIA"

    DUAL = "DUAL"


# ==========================================================
# ESTADO DEL GRUPO
# ==========================================================

class GroupStatus(CatalogEnum):
    """
    Estado académico del grupo.

    Este catálogo es CERRADO.
    """

    EN_EJECUCION = "EN_EJECUCION"

    CANCELADA = "CANCELADA"

    TERMINADO_POR_FECHA = "TERMINADO_POR_FECHA"


# ==========================================================
# MUNICIPIOS
# ==========================================================

class GroupMunicipality(CatalogEnum):
    """
    Municipios oficiales del área de influencia de GIA.
    
    Este catálogo es CERRADO.

    Todo grupo de formación debe pertenecer a uno de estos municipios.

    La validación rechazará cualquier valor que no pertenezca al catálogo.
    """

    AMALFI = "AMALFI"

    ANORI = "ANORI"

    CISNEROS = "CISNEROS"

    MACEO = "MACEO"

    PUERTO_BERRIO = "PUERTO_BERRIO"

    PUERTO_NARE = "PUERTO_NARE"

    PUERTO_TRIUNFO = "PUERTO_TRIUNFO"

    REMEDIOS = "REMEDIOS"

    SAN_ROQUE = "SAN_ROQUE"

    SEGOVIA = "SEGOVIA"

    VEGACHI = "VEGACHI"

    YALI = "YALI"

    YOLOMBO = "YOLOMBO"


# ==========================================================
# EXPORTACIONES
# ==========================================================

__all__ = [

    "GroupModality",

    "GroupStatus",

    "GroupMunicipality",

]