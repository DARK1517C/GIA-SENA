"""
app/catalogs/common_catalogs.py
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Catálogos compartidos del sistema GIA.

Este módulo contiene únicamente catálogos reutilizables
por múltiples dominios del sistema.

NO contiene:

    • Labels
    • Alias
    • Validaciones
    • Lógica de normalización

Toda la lógica asociada se encuentra en:

    display.py
    registry.py
    validation.py

Autor:
Proyecto GIA
"""

from __future__ import annotations

from .common import CatalogEnum


# ==========================================================
# NIVELES DE FORMACIÓN
# ==========================================================

class ProgramLevel(CatalogEnum):
    """
    Nivel de formación SENA.
    """

    OPERARIO = "OPERARIO"

    AUXILIAR = "AUXILIAR"

    TECNICO = "TECNICO"

    TECNOLOGO = "TECNOLOGO"


# ==========================================================
# GÉNERO
# ==========================================================

class Gender(CatalogEnum):
    """
    Género registrado del aprendiz.
    """

    MASCULINO = "MASCULINO"

    FEMENINO = "FEMENINO"


# ==========================================================
# TIPO DE DOCUMENTO
# ==========================================================

class DocumentType(CatalogEnum):
    """
    Tipo de documento.

    IMPORTANTE

    Se utilizan nombres descriptivos como valores
    canónicos para facilitar integraciones futuras.
    """

    CEDULA_CIUDADANIA = "CEDULA_CIUDADANIA"

    TARJETA_IDENTIDAD = "TARJETA_IDENTIDAD"

    CEDULA_EXTRANJERIA = "CEDULA_EXTRANJERIA"

    PERMISO_PROTECCION_TEMPORAL = "PERMISO_PROTECCION_TEMPORAL"

    PERMISO_ESPECIAL_PERMANENCIA = "PERMISO_ESPECIAL_PERMANENCIA"


# ==========================================================
# RESPUESTAS SI / NO
# ==========================================================

class YesNo(CatalogEnum):
    """
    Respuesta booleana estandarizada.
    """

    SI = "SI"

    NO = "NO"


# ==========================================================
# ESTADO DEL REGISTRO
# ==========================================================

class RecordStatus(CatalogEnum):
    """
    Estado interno del registro.

    No corresponde al estado de SOFIA.
    """

    ACTIVO = "ACTIVO"

    INACTIVO = "INACTIVO"


# ==========================================================
# MODALIDAD DE FORMACIÓN
# ==========================================================

class TrainingModality(CatalogEnum):
    """
    Modalidad general de formación.

    Utilizada por TrainingGroup y futuros módulos.
    """

    PRESENCIAL = "PRESENCIAL"

    VIRTUAL = "VIRTUAL"

    MIXTA = "MIXTA"


# ==========================================================
# EXPORTACIONES
# ==========================================================

__all__ = [

    "ProgramLevel",

    "Gender",

    "DocumentType",

    "YesNo",

    "RecordStatus",

    "TrainingModality",

]