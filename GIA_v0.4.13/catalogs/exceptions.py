"""
Excepciones específicas de los catálogos de GIA.
"""


class CatalogError(Exception):
    """Excepción base para errores relacionados con catálogos."""


class InvalidCatalogValueError(CatalogError):
    """
    Se lanza cuando se intenta utilizar un valor que
    no pertenece a un catálogo válido.
    """


__all__ = [
    "CatalogError",
    "InvalidCatalogValueError",
]