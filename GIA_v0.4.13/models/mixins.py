"""
models/mixins.py

Mixins reutilizables para los modelos del proyecto GIA.

Este módulo concentra pequeños bloques de comportamiento y campos
que pueden ser compartidos por varias entidades del sistema, sin
mezclar lógica de negocio específica.

La idea es reducir duplicación en modelos como:
- centros
- empresas
- evidencias
- actividades
- grupos
- catálogos internos
"""

from __future__ import annotations

from datetime import datetime, timezone

from sqlalchemy import Boolean, DateTime, Integer, String, Text, false, true
from sqlalchemy.orm import Mapped, mapped_column


class CodeMixin:
    """
    Aporta un código único para entidades que necesitan un identificador
    legible o estable distinto del ID numérico.

    Ejemplos:
    - EVID-001
    - CAT-ARL
    - GRP-2026-01
    """

    code: Mapped[str] = mapped_column(
        String(80),
        unique=True,
        index=True,
        nullable=False,
    )


class NameMixin:
    """
    Aporta un nombre principal para la entidad.
    """

    name: Mapped[str] = mapped_column(
        String(150),
        nullable=False,
        index=True,
    )


class TitleMixin:
    """
    Aporta un título corto para entidades que se muestran en interfaz
    o se agrupan por nombre visible.
    """

    title: Mapped[str] = mapped_column(
        String(180),
        nullable=False,
        index=True,
    )


class DescriptionMixin:
    """
    Aporta una descripción opcional.
    """

    description: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )


class OrderMixin:
    """
    Aporta una posición de orden para listas, tarjetas, catálogos o bloques.
    """

    sort_order: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=0,
        server_default="0",
        index=True,
    )


class ActiveMixin:
    """
    Aporta un indicador de estado activo/inactivo.

    Útil en entidades que no deben eliminarse físicamente,
    pero sí deshabilitarse.
    """

    is_active: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=True,
        server_default=true(),
        index=True,
    )

    def activate(self) -> None:
        """
        Activa el registro.
        """
        self.is_active = True

    def deactivate(self) -> None:
        """
        Desactiva el registro.
        """
        self.is_active = False


class SoftDeleteMixin:
    """
    Aporta eliminación lógica.

    En lugar de borrar físicamente el registro,
    se marca como eliminado y se guarda la fecha.
    """

    is_deleted: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=False,
        server_default=false(),
        index=True,
    )

    deleted_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    def mark_deleted(self) -> None:
        """
        Marca el registro como eliminado lógicamente.
        """
        self.is_deleted = True
        self.deleted_at = datetime.now(timezone.utc)

    def restore(self) -> None:
        """
        Restaura el registro eliminado lógicamente.
        """
        self.is_deleted = False
        self.deleted_at = None


class SortableNameMixin(NameMixin, OrderMixin):
    """
    Mezcla útil para entidades que necesitan nombre y orden.
    """

    pass


__all__ = [
    "CodeMixin",
    "NameMixin",
    "TitleMixin",
    "DescriptionMixin",
    "OrderMixin",
    "ActiveMixin",
    "SoftDeleteMixin",
    "SortableNameMixin",
]