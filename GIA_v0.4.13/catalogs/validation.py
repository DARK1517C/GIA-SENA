"""
app/catalogs/validation.py
~~~~~~~~~~~~~~~~~~~~~~~~~~

Servicios de validación para el sistema de
catálogos de GIA.

Este módulo utiliza CatalogRegistry para validar
y normalizar valores.

No contiene reglas de negocio.

Autor:
Proyecto GIA
"""

from __future__ import annotations

from typing import TypeAlias

from .common import CatalogEnum
from .exceptions import InvalidCatalogValueError
from .registry import CatalogRegistry

CatalogType: TypeAlias = type[CatalogEnum]


class CatalogValidation:
    """
    Servicios de validación para catálogos.
    """

    # --------------------------------------------------
    # VALIDATE
    # --------------------------------------------------

    @classmethod
    def validate(
        cls,
        catalog: CatalogType,
        value: str | None,
    ) -> str | None:
        """
        Valida un valor utilizando
        CatalogRegistry.
        """

        return CatalogRegistry.validate(
            catalog,
            value,
        )

    # --------------------------------------------------
    # REQUIRED
    # --------------------------------------------------

    @classmethod
    def validate_required(
        cls,
        catalog: CatalogType,
        value: str | None,
    ) -> str:
        """
        Valida un campo obligatorio.
        """

        if value is None:

            raise InvalidCatalogValueError(
                "El valor es obligatorio."
            )

        if isinstance(value, str):

            value = value.strip()

            if value == "":

                raise InvalidCatalogValueError(
                    "El valor es obligatorio."
                )

        result = CatalogRegistry.validate(
            catalog,
            value,
        )

        if result is None:

            raise InvalidCatalogValueError(
                "El valor es obligatorio."
            )

        return result

    # --------------------------------------------------
    # OPTIONAL
    # --------------------------------------------------

    @classmethod
    def validate_optional(
        cls,
        catalog: CatalogType,
        value: str | None,
    ) -> str | None:
        """
        Valida un campo opcional.
        """

        if value is None:

            return None

        if isinstance(value, str):

            value = value.strip()

            if value == "":

                return None

        return CatalogRegistry.validate(
            catalog,
            value,
        )

    # --------------------------------------------------
    # MANY
    # --------------------------------------------------

    @classmethod
    def validate_many(
        cls,
        catalog: CatalogType,
        values: list[str],
    ) -> list[str]:
        """
        Valida múltiples valores.
        """

        return [

            cls.validate_required(
                catalog,
                value,
            )

            for value in values

        ]

    # --------------------------------------------------
    # NORMALIZE MANY
    # --------------------------------------------------

    @classmethod
    def normalize_many(
        cls,
        catalog: CatalogType,
        values: list[str],
    ) -> list[str]:
        """
        Normaliza una colección de valores.
        """

        result = []

        for value in values:

            normalized = CatalogRegistry.normalize(
                catalog,
                value,
            )

            if normalized is not None:

                result.append(
                    normalized
                )

        return result

    # --------------------------------------------------
    # IS VALID
    # --------------------------------------------------

    @classmethod
    def is_valid(
        cls,
        catalog: CatalogType,
        value: str | None,
    ) -> bool:
        """
        Indica si un valor es válido.
        """

        try:

            CatalogRegistry.validate(
                catalog,
                value,
            )

            return True

        except InvalidCatalogValueError:

            return False


__all__ = [

    "CatalogValidation",

]