"""
app/catalogs/registry.py
~~~~~~~~~~~~~~~~~~~~~~~~

Registro central del sistema de catálogos de GIA.

Este módulo constituye el punto único de acceso a todos
los catálogos de la aplicación.

Responsabilidades:

    • Obtener valores canónicos.
    • Obtener etiquetas (labels).
    • Resolver alias.
    • Normalizar valores.
    • Validar catálogos.
    • Generar choices para WTForms.
    • Identificar catálogos abiertos y cerrados.

IMPORTANTE

Este módulo NO contiene lógica de negocio.

Autor:
Proyecto GIA
"""

from __future__ import annotations

from functools import lru_cache

from typing import Type

from .aliases import CATALOG_ALIASES
from .common import (
    CatalogEnum,
    normalize_text,
)
from .display import CATALOG_LABELS
from .exceptions import (
    InvalidCatalogValueError,
)

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
)

from .training_group import (
    GroupModality,
    GroupStatus,
    GroupMunicipality,
)

# ==========================================================
# CATÁLOGOS CERRADOS
# ==========================================================

CLOSED_CATALOGS = {

    ProgramLevel,

    Gender,

    DocumentType,

    YesNo,

    RecordStatus,

    SofiaStatus,

    EpModality,

    GroupModality,

    GroupStatus,

    GroupMunicipality,

}

# ==========================================================
# CATÁLOGOS ABIERTOS
# ==========================================================

OPEN_CATALOGS = {

    # MunicipalityOrigin
    #
    # Se añadirán aquí
    # los catálogos abiertos
    # del sistema.

}


# ==========================================================
# REGISTRY
# ==========================================================

class CatalogRegistry:
    """
    Registro central del sistema de catálogos.

    Toda la aplicación debe acceder a los
    catálogos mediante esta clase.
    """

    # --------------------------------------------------
    # VALUES
    # --------------------------------------------------

    @classmethod
    @lru_cache(maxsize=None)
    def values(
        cls,
        catalog: Type[CatalogEnum],
    ) -> tuple[str, ...]:
        """
        Devuelve todos los valores canónicos.
        """

        return tuple(catalog.values())

    # --------------------------------------------------
    # EXISTS
    # --------------------------------------------------

    @classmethod
    def exists(
        cls,
        catalog: Type[CatalogEnum],
        value: str | None,
    ) -> bool:
        """
        Indica si un valor pertenece
        al catálogo.
        """

        if value is None:

            return False

        return catalog.has_value(value)

    # --------------------------------------------------
    # LABELS
    # --------------------------------------------------

    @classmethod
    @lru_cache(maxsize=None)
    def labels(
        cls,
        catalog: Type[CatalogEnum],
    ) -> dict:
        """
        Devuelve el diccionario de labels
        del catálogo.
        """

        return CATALOG_LABELS.get(
            catalog,
            {},
        )

    # --------------------------------------------------
    # LABEL
    # --------------------------------------------------

    @classmethod
    def label(
        cls,
        catalog: Type[CatalogEnum],
        value: str | CatalogEnum,
    ) -> str:
        """
        Devuelve el texto mostrado
        en la interfaz.
        """

        labels = cls.labels(catalog)

        return labels.get(
            value,
            str(value),
        )

    # --------------------------------------------------
    # ALIASES
    # --------------------------------------------------

    @classmethod
    @lru_cache(maxsize=None)
    def aliases(
        cls,
        catalog: Type[CatalogEnum],
    ) -> dict:
        """
        Devuelve el diccionario
        de alias del catálogo.
        """

        return CATALOG_ALIASES.get(
            catalog,
            {},
        )

    # --------------------------------------------------
    # NORMALIZE
    # --------------------------------------------------

    @classmethod
    def normalize(
        cls,
        catalog: Type[CatalogEnum],
        value: str | None,
    ) -> str | None:
        """
        Convierte un texto cualquiera
        en un valor canónico.

        Devuelve None cuando
        no encuentra coincidencia.
        """

        if value is None:

            return None

        normalized = normalize_text(value)

        aliases = cls.aliases(catalog)

        if normalized in aliases:

            return aliases[
                normalized
            ].value

        if catalog.has_value(normalized):

            return normalized

        return None

    # --------------------------------------------------
    # CANONICAL
    # --------------------------------------------------

    @classmethod
    def canonical(
        cls,
        catalog: Type[CatalogEnum],
        value: str | None,
    ) -> CatalogEnum | None:
        """
        Devuelve el Enum correspondiente
        al valor recibido.
        """

        normalized = cls.normalize(
            catalog,
            value,
        )

        if normalized is None:

            return None

        return catalog(normalized)
    
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
        el tipo de catálogo.

        Catálogos cerrados:
            Deben pertenecer al catálogo.

        Catálogos abiertos:
            Si no existe una coincidencia,
            se conserva el valor original.
        """

        normalized = cls.normalize(
            catalog,
            value,
        )

        if normalized is not None:

            return normalized

        if cls.is_open(catalog):

            return value

        raise InvalidCatalogValueError(
            f'"{value}" no pertenece al catálogo '
            f'{catalog.__name__}.'
        )

    # --------------------------------------------------
    # CHOICES
    # --------------------------------------------------

    @classmethod
    @lru_cache(maxsize=None)
    def choices(
        cls,
        catalog: CatalogType,
    ) -> tuple[tuple[str, str], ...]:
        """
        Devuelve los choices utilizados
        por formularios.

        Ejemplo:

            (
                ("TECNICO", "Técnico"),
                ("TECNOLOGO", "Tecnólogo"),
            )
        """

        labels = cls.labels(catalog)

        return tuple(

            (

                item.value,

                labels.get(
                    item,
                    item.value,
                ),

            )

            for item in catalog

        )

    # --------------------------------------------------
    # IS CLOSED
    # --------------------------------------------------

    @classmethod
    def is_closed(
        cls,
        catalog: CatalogType,
    ) -> bool:
        """
        Indica si el catálogo
        es cerrado.
        """

        return catalog in CLOSED_CATALOGS

    # --------------------------------------------------
    # IS OPEN
    # --------------------------------------------------

    @classmethod
    def is_open(
        cls,
        catalog: CatalogType,
    ) -> bool:
        """
        Indica si el catálogo
        es abierto.
        """

        return catalog in OPEN_CATALOGS


# ==========================================================
# EXPORTACIONES
# ==========================================================

__all__ = [

    "CatalogRegistry",

    "CLOSED_CATALOGS",

    "OPEN_CATALOGS",

]