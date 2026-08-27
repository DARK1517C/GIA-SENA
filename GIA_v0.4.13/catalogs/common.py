"""
app/catalogs/common.py
~~~~~~~~~~~~~~~~~~~~~~

Infraestructura base para el sistema de catálogos de GIA.

Este módulo NO contiene catálogos.

Su responsabilidad es proporcionar:

    • Clase base para todos los catálogos.
    • Funciones de normalización.
    • Utilidades comunes.
    • Tipos reutilizables.

Todos los demás archivos de catalogs dependen de este módulo.

Autor:
Proyecto GIA
"""

from __future__ import annotations

import re
import unicodedata

from enum import Enum
from typing import Iterable
from typing import Optional
from typing import TypeVar


# ==========================================================
# CLASE BASE
# ==========================================================

class CatalogEnum(str, Enum):
    """
    Clase base para todos los catálogos del sistema.

    Todos los catálogos deben heredar de esta clase.

    Ejemplo:

        class ProgramLevel(CatalogEnum):

            TECNICO = "TECNICO"

            TECNOLOGO = "TECNOLOGO"
    """

    @classmethod
    def values(cls) -> list[str]:
        """
        Devuelve todos los valores canónicos.

        Returns
        -------
        list[str]
        """

        return [item.value for item in cls]

    # ------------------------------------------------------

    @classmethod
    def count(cls) -> int:
        """
        Número de elementos del catálogo.
        """

        return len(cls.values())

    # ------------------------------------------------------

    @classmethod
    def has_value(cls, value: Optional[str]) -> bool:
        """
        Verifica si un valor pertenece
        al catálogo.
        """

        if value is None:
            return False

        return value in cls.values()

    # ------------------------------------------------------

    @classmethod
    def values_set(cls) -> set[str]:
        """
        Devuelve un set de valores.

        Muy útil para búsquedas O(1).
        """

        return set(cls.values())

    # ------------------------------------------------------

    @classmethod
    def from_value(cls, value: str):
        """
        Convierte un string
        al Enum correspondiente.

        Lanza ValueError
        si no existe.
        """

        return cls(value)


# ==========================================================
# NORMALIZACIÓN
# ==========================================================

def remove_accents(text: str) -> str:
    """
    Elimina acentos.

    Ejemplo

        Técnólogo

        →

        Tecnologo
    """

    normalized = unicodedata.normalize("NFKD", text)

    return "".join(
        char
        for char in normalized
        if not unicodedata.combining(char)
    )


# ----------------------------------------------------------


def normalize_spaces(text: str) -> str:
    """
    Reemplaza múltiples espacios
    por un solo espacio.
    """

    return re.sub(r"\s+", " ", text).strip()


# ----------------------------------------------------------


def normalize_text(value: Optional[str]) -> str:
    """
    Convierte cualquier texto
    a un formato comparable.

    Reglas:

    - elimina acentos
    - elimina espacios sobrantes
    - convierte a MAYÚSCULAS

    Ejemplo

        "  Técnico "

        →

        "TECNICO"
    """

    if value is None:
        return ""

    value = str(value)

    value = normalize_spaces(value)

    value = remove_accents(value)

    value = value.upper()

    return value


# ==========================================================
# UTILIDADES
# ==========================================================

def unique(values: Iterable[str]) -> list[str]:
    """
    Elimina duplicados
    conservando el orden.

    Ejemplo

        ["A","B","A"]

        →

        ["A","B"]
    """

    seen = set()

    result = []

    for value in values:

        if value in seen:

            continue

        seen.add(value)

        result.append(value)

    return result


# ----------------------------------------------------------


def is_empty(value: Optional[str]) -> bool:
    """
    Determina si un valor
    es considerado vacío.
    """

    if value is None:
        return True

    if not str(value).strip():
        return True

    return False


# ----------------------------------------------------------


def clean_string(value: Optional[str]) -> Optional[str]:
    """
    Limpia un texto.

    Devuelve None
    si queda vacío.
    """

    if is_empty(value):

        return None

    return normalize_spaces(str(value))


# ==========================================================
# EXPORTACIONES
# ==========================================================

__all__ = [

    "CatalogEnum",

    "normalize_text",

    "normalize_spaces",

    "remove_accents",

    "clean_string",

    "is_empty",

    "unique",

]