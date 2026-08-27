"""
models/base.py

Clase base abstracta para todos los modelos del proyecto GIA.

Este módulo concentra la funcionalidad común que comparten las entidades
persistentes del sistema, evitando duplicación de código en los modelos
de dominio.

Responsabilidades:
- Proveer un identificador primario.
- Gestionar fechas de creación y actualización.
- Ofrecer serialización básica a diccionario.
- Ofrecer actualización genérica de atributos.
- Definir representación legible del objeto.

No contiene lógica de negocio específica de ningún dominio.
"""

from __future__ import annotations

from datetime import datetime
from typing import Any

from sqlalchemy import DateTime, Integer, func
from sqlalchemy.orm import Mapped, mapped_column

from extensions import db


class BaseModel(db.Model):
    """
    Clase base abstracta para los modelos del sistema.

    Todos los modelos concretos deben heredar de esta clase.
    """

    __abstract__ = True

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    # Keep both Python-side and server-side defaults.
    # The Python defaults are important for legacy SQLite databases whose
    # existing columns may be NOT NULL without a server DEFAULT clause.
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        default=func.now(),
        server_default=func.now(),
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True),
        nullable=False,
        default=func.now(),
        server_default=func.now(),
        onupdate=func.now(),
    )

    def to_dict(self) -> dict[str, Any]:
        """
        Serializa el modelo a un diccionario básico.

        Incluye todos los atributos públicos de la instancia que no
        comienzan con "_" y que no son llamados.
        """
        result: dict[str, Any] = {}

        for key in self.__mapper__.columns.keys():
            result[key] = getattr(self, key)

        return result

    def update(self, **kwargs: Any) -> None:
        """
        Actualiza atributos del modelo de forma dinámica.

        Ejemplo:
            instance.update(name="Nuevo nombre", active=True)
        """
        for key, value in kwargs.items():
            if hasattr(self, key):
                setattr(self, key, value)

    def __repr__(self) -> str:
        """
        Representación legible del objeto para depuración.
        """
        attrs = []
        if hasattr(self, "id"):
            attrs.append(f"id={self.id!r}")
        return f"<{self.__class__.__name__} {' '.join(attrs)}>"