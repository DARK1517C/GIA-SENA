"""
models/evidence.py

Modelo del dominio de evidencias del sistema GIA.

El módulo organiza el flujo de evidencias con un enfoque tipo LMS:

- categorías institucionales
- plantillas oficiales
- actividades asignadas a una ficha
- entregas de los aprendices
- revisión y corrección
- aprobación
- firma de documentos PDF
- comentarios y retroalimentación

Estados oficiales de una entrega:

1. no_entregado
2. pendiente_revision
3. requiere_correccion
4. aprobado

La firma se almacena como parte de la entrega mediante
signed_file_name, signed_file_path y signed_at.
"""

from __future__ import annotations

import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, TypeVar

from sqlalchemy import (
    Boolean,
    DateTime,
    column,
    ForeignKey,
    Integer,
    Index,
    String,
    Text,
    false,
    true,
)
from sqlalchemy.orm import (
    Mapped,
    backref,
    mapped_column,
    relationship,
    validates,
)

from catalogs.common import CatalogEnum, normalize_spaces
from catalogs.exceptions import InvalidCatalogValueError

from .base import BaseModel


CatalogEnumT = TypeVar(
    "CatalogEnumT",
    bound=CatalogEnum,
)


# ==============================================================================
# ESTADOS DE LA ENTREGA
# ==============================================================================

EVIDENCE_STATUS_NOT_SUBMITTED = "no_entregado"

EVIDENCE_STATUS_PENDING_REVIEW = "pendiente_revision"

EVIDENCE_STATUS_REQUIRES_CORRECTION = "requiere_correccion"

EVIDENCE_STATUS_APPROVED = "aprobado"


EVIDENCE_STATUS_LABELS = {
    EVIDENCE_STATUS_NOT_SUBMITTED: "No entregado",
    EVIDENCE_STATUS_PENDING_REVIEW: "Pendiente de revisión",
    EVIDENCE_STATUS_REQUIRES_CORRECTION: "Requiere corrección",
    EVIDENCE_STATUS_APPROVED: "Aprobado",
}


EVIDENCE_STATUS_COLORS = {
    EVIDENCE_STATUS_NOT_SUBMITTED: "#d13438",
    EVIDENCE_STATUS_PENDING_REVIEW: "#fdc300",
    EVIDENCE_STATUS_REQUIRES_CORRECTION: "#ff8c00",
    EVIDENCE_STATUS_APPROVED: "#39a900",
}


EVIDENCE_STATUS_ORDER = (
    EVIDENCE_STATUS_NOT_SUBMITTED,
    EVIDENCE_STATUS_PENDING_REVIEW,
    EVIDENCE_STATUS_REQUIRES_CORRECTION,
    EVIDENCE_STATUS_APPROVED,
)


EVIDENCE_STATUSES = EVIDENCE_STATUS_ORDER


# ==============================================================================
# ORIGEN DE LA ACTIVIDAD
# ==============================================================================

EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE = "template"

EVIDENCE_ACTIVITY_ORIGIN_CUSTOM = "custom"


EVIDENCE_ACTIVITY_ORIGINS = (
    EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE,
    EVIDENCE_ACTIVITY_ORIGIN_CUSTOM,
)


EVIDENCE_ACTIVITY_ORIGIN_LABELS = {
    EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE: "Plantilla institucional",
    EVIDENCE_ACTIVITY_ORIGIN_CUSTOM: "Personalizada por el instructor",
}


# ==============================================================================
# CATÁLOGO DE EVIDENCIAS
# ==============================================================================
#
# La arquitectura oficial vive en la BD mediante:
#   EvidenceCategory -> EvidenceTemplate -> EvidenceActivity -> EvidenceSubmission
#
# Los catálogos históricos fueron retirados del dominio.
# Las categorías y plantillas institucionales se cargan mediante Alembic.


# ==============================================================================
# HELPERS DE NORMALIZACIÓN
# ==============================================================================

def _normalize_required_text(
    value: Any,
    field_name: str,
) -> str:
    if value is None:
        raise ValueError(
            f"El campo {field_name} es obligatorio."
        )

    normalized = normalize_spaces(str(value))

    if not normalized:
        raise ValueError(
            f"El campo {field_name} es obligatorio."
        )

    return normalized


def _normalize_optional_text(
    value: Any,
) -> str | None:
    if value is None:
        return None

    normalized = normalize_spaces(str(value))

    return normalized or None


def _normalize_required_code(
    value: Any,
    field_name: str,
) -> str:
    text = _normalize_required_text(
        value,
        field_name,
    )

    text = text.upper()

    text = re.sub(
        r"[^A-Z0-9]+",
        "_",
        text,
    )

    text = re.sub(
        r"_+",
        "_",
        text,
    ).strip("_")

    if not text:
        raise ValueError(
            f"El campo {field_name} es obligatorio."
        )

    return text


def _normalize_optional_code(
    value: Any,
) -> str | None:
    if value is None:
        return None

    text = normalize_spaces(str(value))

    if not text:
        return None

    text = text.upper()

    text = re.sub(
        r"[^A-Z0-9]+",
        "_",
        text,
    )

    text = re.sub(
        r"_+",
        "_",
        text,
    ).strip("_")

    return text or None


def _normalize_optional_int(
    value: Any,
) -> int | None:
    if value is None:
        return None

    try:
        number = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            "Se esperaba un valor numérico entero."
        ) from exc

    if number < 0:
        raise ValueError(
            "El valor numérico no puede ser negativo."
        )

    return number


def _normalize_required_int(
    value: Any,
    field_name: str,
) -> int:
    normalized = _normalize_optional_int(value)

    if normalized is None:
        raise ValueError(
            f"El campo {field_name} es obligatorio."
        )

    return normalized


def _normalize_bool(
    value: Any,
) -> bool:
    if isinstance(value, bool):
        return value

    if value is None:
        return False

    text = normalize_spaces(
        str(value)
    ).lower()

    if text in {
        "1",
        "true",
        "yes",
        "si",
        "sí",
        "on",
    }:
        return True

    if text in {
        "0",
        "false",
        "no",
        "off",
        "",
    }:
        return False

    return bool(value)


def _normalize_allowed_extensions(
    value: Any,
) -> str | None:
    if value is None:
        return None

    if isinstance(value, str):
        raw_items = re.split(
            r"[,\s;|]+",
            value,
        )

    elif isinstance(value, Iterable):
        raw_items = list(value)

    else:
        raw_items = [value]

    normalized_items: list[str] = []
    seen: set[str] = set()

    for item in raw_items:
        token = normalize_spaces(
            str(item)
        ).lower()

        if not token:
            continue

        token = token.lstrip(".")
        token = token.replace("*", "")
        token = token.strip()

        if not token:
            continue

        token = f".{token}"

        if token not in seen:
            seen.add(token)
            normalized_items.append(token)

    return (
        ",".join(normalized_items)
        if normalized_items
        else None
    )


def _split_allowed_extensions(
    value: str | None,
) -> tuple[str, ...]:
    if not value:
        return tuple()

    return tuple(
        item.strip()
        for item in value.split(",")
        if item.strip()
    )


def _normalize_mime_type(
    value: Any,
) -> str | None:
    if value is None:
        return None

    text = normalize_spaces(
        str(value)
    ).lower()

    return text or None


def _normalize_file_path(
    value: Any,
) -> str | None:
    if value is None:
        return None

    text = normalize_spaces(
        str(value)
    )

    return text or None


def _is_pdf_mime(
    mime_type: str | None,
) -> bool:
    return bool(
        mime_type
        and mime_type.lower() == "application/pdf"
    )


def _is_pdf_extension(
    file_name: str | None,
) -> bool:
    if not file_name:
        return False

    return (
        Path(file_name)
        .suffix
        .lower()
        == ".pdf"
    )


def _normalize_status(
    value: Any,
) -> str:
    if value is None:
        return EVIDENCE_STATUS_NOT_SUBMITTED

    normalized = normalize_spaces(
        str(value)
    ).lower()

    valid_statuses = {
        EVIDENCE_STATUS_NOT_SUBMITTED,
        EVIDENCE_STATUS_PENDING_REVIEW,
        EVIDENCE_STATUS_REQUIRES_CORRECTION,
        EVIDENCE_STATUS_APPROVED,
    }

    if normalized in valid_statuses:
        return normalized

    raise InvalidCatalogValueError(
        f"Estado de evidencia no válido: {value!r}"
    )


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


# ==============================================================================
# MODELO DE CATEGORÍA
# ==============================================================================

class EvidenceCategory(BaseModel):
    """
    Categoría institucional de evidencias.

    Ejemplos:

    - Requisitos Iniciales
    - Bitacoras
    - Momentos de Seguimiento
    - Requisitos de Certificacion
    """

    __tablename__ = "evidence_category"

    code: Mapped[str] = mapped_column(
        String(80),
        unique=True,
        nullable=False,
        index=True,
    )

    name: Mapped[str] = mapped_column(
        String(120),
        unique=True,
        nullable=False,
        index=True,
    )

    description: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    icon: Mapped[str | None] = mapped_column(
        String(80),
        nullable=True,
    )

    color: Mapped[str | None] = mapped_column(
        String(20),
        nullable=True,
    )

    sort_order: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=0,
    )

    is_active: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=True,
        server_default=true(),
        index=True,
    )

    templates = relationship(
        "EvidenceTemplate",
        back_populates="category",
        lazy="selectin",
    )

    activities = relationship(
        "EvidenceActivity",
        back_populates="category",
        lazy="selectin",
    )

    @validates("code")
    def validate_code(
        self,
        key: str,
        value: Any,
    ) -> str:
        return _normalize_required_code(
            value,
            "código de categoría",
        )

    @validates("name")
    def validate_name(
        self,
        key: str,
        value: Any,
    ) -> str:
        return _normalize_required_text(
            value,
            "nombre de categoría",
        )

    @validates(
        "description",
        "icon",
        "color",
    )
    def validate_optional_text_fields(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_optional_text(value)

    @validates("sort_order")
    def validate_sort_order(
        self,
        key: str,
        value: Any,
    ) -> int:
        return _normalize_required_int(
            value,
            "orden de categoría",
        )

    @validates("is_active")
    def validate_is_active(
        self,
        key: str,
        value: Any,
    ) -> bool:
        return _normalize_bool(value)

    def __str__(self) -> str:
        return self.name

    def __repr__(self) -> str:
        return (
            f"<EvidenceCategory "
            f"id={getattr(self, 'id', None)!r} "
            f"code={self.code!r} "
            f"name={self.name!r}>"
        )


# ==============================================================================
# MODELO DE PLANTILLA
# ==============================================================================

class EvidenceTemplate(BaseModel):
    """
    Plantilla institucional de evidencia.

    Sirve como base para crear actividades de evidencia
    asignadas a una ficha.
    """

    __tablename__ = "evidence_template"

    category_id: Mapped[int] = mapped_column(
        ForeignKey("evidence_category.id"),
        nullable=False,
        index=True,
    )

    code: Mapped[str] = mapped_column(
        String(120),
        unique=True,
        nullable=False,
        index=True,
    )

    title: Mapped[str] = mapped_column(
        String(180),
        nullable=False,
        index=True,
    )

    description: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    allowed_extensions: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    max_file_size_mb: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )

    requires_signature: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=False,
    )

    is_required: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=True,
    )

    sort_order: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=0,
    )

    is_active: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=True,
        index=True,
    )

    created_by_id: Mapped[int | None] = mapped_column(
        ForeignKey("user.id"),
        nullable=True,
        index=True,
    )

    category = relationship(
        "EvidenceCategory",
        back_populates="templates",
        lazy="select",
    )

    created_by = relationship(
        "User",
        foreign_keys=[created_by_id],
        backref=backref(
            "created_evidence_templates",
            lazy="selectin",
        ),
    )

    activities = relationship(
        "EvidenceActivity",
        back_populates="template",
        lazy="selectin",
    )

    @validates("category_id")
    def validate_category_id(
        self,
        key: str,
        value: Any,
    ) -> int:
        return _normalize_required_int(
            value,
            "categoría",
        )

    @validates("code")
    def validate_code(
        self,
        key: str,
        value: Any,
    ) -> str:
        return _normalize_required_code(
            value,
            "código de plantilla",
        )

    @validates("title")
    def validate_title(
        self,
        key: str,
        value: Any,
    ) -> str:
        return _normalize_required_text(
            value,
            "título de plantilla",
        )

    @validates("description")
    def validate_description(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_optional_text(value)

    @validates("allowed_extensions")
    def validate_allowed_extensions(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_allowed_extensions(value)

    @validates("max_file_size_mb")
    def validate_max_file_size_mb(
        self,
        key: str,
        value: Any,
    ) -> int | None:
        return _normalize_optional_int(value)

    @validates("sort_order")
    def validate_sort_order(
        self,
        key: str,
        value: Any,
    ) -> int:
        return _normalize_required_int(
            value,
            "orden de plantilla",
        )

    @validates("created_by_id")
    def validate_created_by_id(
        self,
        key: str,
        value: Any,
    ) -> int | None:
        if value is None:
            return None

        return _normalize_required_int(
            value,
            "usuario creador",
        )

    @validates(
        "requires_signature",
        "is_required",
        "is_active",
    )
    def validate_boolean_fields(
        self,
        key: str,
        value: Any,
    ) -> bool:
        return _normalize_bool(value)

    @property
    def allowed_extensions_list(
        self,
    ) -> tuple[str, ...]:
        return _split_allowed_extensions(
            self.allowed_extensions
        )

    @property
    def is_pdf_only(self) -> bool:
        return (
            self.allowed_extensions_list
            == (".pdf",)
        )

    @classmethod
    def from_defaults(
        cls,
        category_id: int,
        code: str,
        title: str,
        description: str | None = None,
        allowed_extensions: (
            str | Iterable[str] | None
        ) = None,
        max_file_size_mb: int | None = None,
        requires_signature: bool = False,
        is_required: bool = True,
        sort_order: int = 0,
        is_active: bool = True,
        created_by_id: int | None = None,
    ) -> "EvidenceTemplate":

        return cls(
            category_id=category_id,
            code=code,
            title=title,
            description=description,
            allowed_extensions=(
                _normalize_allowed_extensions(
                    allowed_extensions
                )
            ),
            max_file_size_mb=max_file_size_mb,
            requires_signature=requires_signature,
            is_required=is_required,
            sort_order=sort_order,
            is_active=is_active,
            created_by_id=created_by_id,
        )

    def __str__(self) -> str:
        return self.title

    def __repr__(self) -> str:
        return (
            f"<EvidenceTemplate "
            f"id={getattr(self, 'id', None)!r} "
            f"code={self.code!r} "
            f"title={self.title!r}>"
        )


# ==============================================================================
# MODELO DE ACTIVIDAD
# ==============================================================================

class EvidenceActivity(BaseModel):
    """
    Actividad de evidencia asignada a una ficha.

    Puede provenir de una plantilla institucional
    o ser creada de forma personalizada por un instructor.
    """

    __tablename__ = "evidence_activity"

    group_id: Mapped[int] = mapped_column(
        ForeignKey("training_group.id"),
        nullable=False,
        index=True,
    )

    template_id: Mapped[int | None] = mapped_column(
        ForeignKey("evidence_template.id"),
        nullable=True,
        index=True,
    )

    category_id: Mapped[int] = mapped_column(
        ForeignKey("evidence_category.id"),
        nullable=False,
        index=True,
    )

    code: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
        index=True,
    )

    title: Mapped[str] = mapped_column(
        String(180),
        nullable=False,
        index=True,
    )

    description: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    due_start: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    due_end: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    allowed_extensions: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    max_file_size_mb: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )

    requires_signature: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=False,
    )

    is_required: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=True,
        server_default=true(),
    )

    is_visible: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=True,
    )

    is_default: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=True,
    )

    origin: Mapped[str] = mapped_column(
        String(20),
        nullable=False,
        default=EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE,
        server_default=EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE,
        index=True,
    )

    sort_order: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=0,
    )

    created_by_id: Mapped[int | None] = mapped_column(
        ForeignKey("user.id"),
        nullable=True,
        index=True,
    )

    group = relationship(
        "TrainingGroup",
        back_populates="evidence_activities",
        lazy="select",
    )

    template = relationship(
        "EvidenceTemplate",
        back_populates="activities",
        lazy="select",
    )

    category = relationship(
        "EvidenceCategory",
        back_populates="activities",
        lazy="select",
    )

    created_by = relationship(
        "User",
        foreign_keys=[created_by_id],
        backref=backref(
            "created_evidence_activities",
            lazy="selectin",
        ),
    )

    submissions = relationship(
        "EvidenceSubmission",
        back_populates="activity",
        lazy="selectin",
        cascade="all, delete-orphan",
    )

    @validates(
        "group_id",
        "category_id",
        "template_id",
        "created_by_id",
    )
    def validate_fk_ids(
        self,
        key: str,
        value: Any,
    ) -> int | None:

        if value is None:
            return None

        return _normalize_required_int(
            value,
            key,
        )

    @validates("code")
    def validate_code(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_optional_code(value)

    @validates("title")
    def validate_title(
        self,
        key: str,
        value: Any,
    ) -> str:
        return _normalize_required_text(
            value,
            "título de actividad",
        )

    @validates(
        "description",
        "due_start",
        "due_end",
    )
    def validate_optional_text_fields(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_optional_text(value)

    @validates("allowed_extensions")
    def validate_allowed_extensions(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_allowed_extensions(value)

    @validates("max_file_size_mb")
    def validate_max_file_size_mb(
        self,
        key: str,
        value: Any,
    ) -> int | None:
        return _normalize_optional_int(value)

    @validates("sort_order")
    def validate_sort_order(
        self,
        key: str,
        value: Any,
    ) -> int:
        return _normalize_required_int(
            value,
            "orden de actividad",
        )

    @validates(
        "requires_signature",
        "is_required",
        "is_visible",
        "is_default",
    )
    def validate_boolean_fields(
        self,
        key: str,
        value: Any,
    ) -> bool:
        return _normalize_bool(value)

    @validates("origin")
    def validate_origin(
        self,
        key: str,
        value: Any,
    ) -> str:

        if value is None:
            return EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE

        normalized = normalize_spaces(
            str(value)
        ).lower()

        if normalized in EVIDENCE_ACTIVITY_ORIGINS:
            return normalized

        raise InvalidCatalogValueError(
            f"Origen de actividad no válido: {value!r}"
        )

    @property
    def allowed_extensions_list(
        self,
    ) -> tuple[str, ...]:
        return _split_allowed_extensions(
            self.allowed_extensions
        )

    @property
    def is_pdf_only(self) -> bool:
        return (
            self.allowed_extensions_list
            == (".pdf",)
        )

    @property
    def is_custom(self) -> bool:
        return (
            self.origin
            == EVIDENCE_ACTIVITY_ORIGIN_CUSTOM
        )

    @property
    def is_from_template(self) -> bool:
        return (
            self.origin
            == EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE
        )

    @property
    def title_with_group(self) -> str:
        if self.group:
            return (
                f"{self.group.group_number} - "
                f"{self.title}"
            )

        return self.title

    @classmethod
    def from_template(
        cls,
        template: EvidenceTemplate,
        group_id: int,
        created_by_id: int | None = None,
        **overrides: Any,
    ) -> "EvidenceActivity":

        activity = cls(
            group_id=group_id,
            template_id=template.id,
            category_id=template.category_id,
            code=(
                overrides.get("code")
                or template.code
            ),
            title=(
                overrides.get("title")
                or template.title
            ),
            description=(
                overrides.get("description")
                or template.description
            ),
            due_start=overrides.get(
                "due_start"
            ),
            due_end=overrides.get(
                "due_end"
            ),
            allowed_extensions=(
                overrides.get(
                    "allowed_extensions"
                )
                or template.allowed_extensions
            ),
            max_file_size_mb=(
                overrides.get(
                    "max_file_size_mb"
                )
                if overrides.get(
                    "max_file_size_mb"
                ) is not None
                else template.max_file_size_mb
            ),
            requires_signature=overrides.get(
                "requires_signature",
                template.requires_signature,
            ),
            is_required=overrides.get(
                "is_required",
                template.is_required,
            ),
            is_visible=overrides.get(
                "is_visible",
                True,
            ),
            is_default=overrides.get(
                "is_default",
                True,
            ),
            origin=overrides.get(
                "origin",
                EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE,
            ),
            sort_order=overrides.get(
                "sort_order",
                template.sort_order,
            ),
            created_by_id=created_by_id,
        )
        activity.validate_domain_consistency(template=template)
        return activity

    def validate_domain_consistency(self, template=None, category=None) -> None:
        """Valida invariantes que no puede expresar una FK simple.

        - Una actividad de plantilla debe apuntar a esa plantilla.
        - La categoría de actividad debe coincidir con la de la plantilla.
        - Una actividad personalizada no puede depender de una plantilla.
        """
        if self.origin == EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE:
            if self.template_id is None and template is None:
                raise ValueError(
                    "Una actividad de plantilla requiere template_id."
                )
            if template is not None and self.category_id != template.category_id:
                raise ValueError(
                    "La categoría de la actividad no coincide con la plantilla."
                )
        elif self.origin == EVIDENCE_ACTIVITY_ORIGIN_CUSTOM:
            if self.template_id is not None:
                raise ValueError(
                    "Una actividad personalizada no puede tener template_id."
                )

        if template is not None and self.category_id != template.category_id:
            raise ValueError(
                "La categoría de la actividad no coincide con la plantilla."
            )

    def __str__(self) -> str:
        return self.title

    def __repr__(self) -> str:
        return (
            f"<EvidenceActivity "
            f"id={getattr(self, 'id', None)!r} "
            f"title={self.title!r}>"
        )


# ==============================================================================
# MODELO DE ENTREGA
# ==============================================================================

class EvidenceSubmission(BaseModel):
    """
    Entrega realizada por un aprendiz sobre una actividad.

    El ciclo normal es:

        no_entregado
            ↓
        pendiente_revision
            ↓
        aprobado

    o:

        pendiente_revision
            ↓
        requiere_correccion
            ↓
        pendiente_revision

    La firma de un PDF se conserva separadamente del archivo
    original mediante signed_file_path.
    """

    __tablename__ = "evidence_submission"

    __table_args__ = (
        Index(
            "uq_evidence_submission_latest_per_activity_apprentice",
            "activity_id",
            "apprentice_id",
            unique=True,
            sqlite_where=(column("is_latest") == True),
            postgresql_where=(column("is_latest") == True),
        ),
    )

    activity_id: Mapped[int] = mapped_column(
        ForeignKey("evidence_activity.id"),
        nullable=False,
        index=True,
    )

    apprentice_id: Mapped[int] = mapped_column(
        ForeignKey("apprentice.id"),
        nullable=False,
        index=True,
    )

    status: Mapped[str] = mapped_column(
        String(40),
        nullable=False,
        default=EVIDENCE_STATUS_NOT_SUBMITTED,
        server_default=EVIDENCE_STATUS_NOT_SUBMITTED,
        index=True,
    )

    file_name: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    file_path: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    mime_type: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
        index=True,
    )

    file_size_bytes: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )

    uploaded_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    # ------------------------------------------------------------------
    # Revisión
    # ------------------------------------------------------------------

    reviewed_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    reviewed_by: Mapped[int | None] = mapped_column(
        ForeignKey("user.id"),
        nullable=True,
        index=True,
    )

    # ------------------------------------------------------------------
    # Aprobación
    # ------------------------------------------------------------------

    approved_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    approved_by_id: Mapped[int | None] = mapped_column(
        ForeignKey("user.id"),
        nullable=True,
        index=True,
    )

    # ------------------------------------------------------------------
    # Archivo firmado
    # ------------------------------------------------------------------

    signed_file_name: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    signed_file_path: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    signed_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    # ------------------------------------------------------------------
    # Versionado
    # ------------------------------------------------------------------

    version_number: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=1,
    )

    attempt_number: Mapped[int] = mapped_column(
        Integer,
        nullable=False,
        default=1,
    )

    is_latest: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=True,
        server_default=true(),
        index=True,
    )

    # ------------------------------------------------------------------
    # Relaciones
    # ------------------------------------------------------------------

    activity = relationship(
        "EvidenceActivity",
        back_populates="submissions",
        lazy="select",
    )

    apprentice = relationship(
        "Apprentice",
        back_populates="evidence_submissions",
        lazy="select",
    )

    reviewed_by_user = relationship(
        "User",
        foreign_keys=[reviewed_by],
        backref=backref(
            "reviewed_evidence_submissions",
            lazy="selectin",
        ),
    )

    approved_by = relationship(
        "User",
        foreign_keys=[approved_by_id],
        backref=backref(
            "approved_evidence_submissions",
            lazy="selectin",
        ),
    )

    comments = relationship(
        "EvidenceComment",
        back_populates="submission",
        lazy="selectin",
        cascade="all, delete-orphan",
        order_by="EvidenceComment.created_at.asc()",
    )

    attempts = relationship(
        "EvidenceSubmissionAttempt",
        back_populates="submission",
        lazy="selectin",
        cascade="all, delete-orphan",
        order_by="EvidenceSubmissionAttempt.attempt_number.asc()",
    )

    @property
    def latest_attempt(self):
        return max(self.attempts, key=lambda item: (item.attempt_number, item.version_number), default=None)

    @property
    def attempt_history(self):
        return tuple(sorted(self.attempts, key=lambda item: item.attempt_number))

    # ------------------------------------------------------------------
    # Validaciones
    # ------------------------------------------------------------------

    @validates(
        "activity_id",
        "apprentice_id",
        "reviewed_by",
        "approved_by_id",
    )
    def validate_fk_ids(
        self,
        key: str,
        value: Any,
    ) -> int | None:

        if value is None:
            return None

        return _normalize_required_int(
            value,
            key,
        )

    @validates("status")
    def validate_status(
        self,
        key: str,
        value: Any,
    ) -> str:
        return _normalize_status(value)

    @validates(
        "file_name",
        "file_path",
        "signed_file_name",
        "signed_file_path",
    )
    def validate_optional_text_fields(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_optional_text(value)

    @validates("mime_type")
    def validate_mime_type(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_mime_type(value)

    @validates("file_size_bytes")
    def validate_file_size_bytes(
        self,
        key: str,
        value: Any,
    ) -> int | None:
        return _normalize_optional_int(value)

    @validates(
        "version_number",
        "attempt_number",
    )
    def validate_positive_int_fields(
        self,
        key: str,
        value: Any,
    ) -> int:

        normalized = _normalize_required_int(
            value,
            key,
        )

        if normalized < 1:
            raise ValueError(
                f"El campo {key} debe ser mayor que cero."
            )

        return normalized

    @validates("is_latest")
    def validate_is_latest(
        self,
        key: str,
        value: Any,
    ) -> bool:
        return _normalize_bool(value)

    # ------------------------------------------------------------------
    # Propiedades
    # ------------------------------------------------------------------

    @property
    def status_label(self) -> str:
        return EVIDENCE_STATUS_LABELS.get(
            self.status,
            self.status,
        )

    @property
    def status_color(self) -> str:
        return EVIDENCE_STATUS_COLORS.get(
            self.status,
            "#737373",
        )

    @property
    def has_file(self) -> bool:
        return bool(self.file_path)

    @property
    def has_signed_file(self) -> bool:
        return bool(self.signed_file_path)

    @property
    def file_extension(self) -> str | None:
        filename = (
            self.file_name
            or self.file_path
        )

        if not filename:
            return None

        extension = Path(
            filename
        ).suffix.lower()

        return extension or None

    @property
    def is_pdf(self) -> bool:
        return (
            _is_pdf_mime(self.mime_type)
            or _is_pdf_extension(self.file_name)
            or _is_pdf_extension(self.file_path)
        )

    @property
    def can_be_signed(self) -> bool:
        """
        Unicamente los PDF pendientes de revisión
        pueden entrar al proceso de firma.
        """

        return (
            self.is_pdf
            and self.status
            == EVIDENCE_STATUS_PENDING_REVIEW
        )

    @property
    def requires_signature(self) -> bool:
        return bool(
            self.activity
            and self.activity.requires_signature
        )

    @property
    def is_approved(self) -> bool:
        return (
            self.status
            == EVIDENCE_STATUS_APPROVED
        )

    @property
    def is_pending_review(self) -> bool:
        return (
            self.status
            == EVIDENCE_STATUS_PENDING_REVIEW
        )

    @property
    def is_requires_correction(self) -> bool:
        return (
            self.status
            == EVIDENCE_STATUS_REQUIRES_CORRECTION
        )

    @property
    def can_be_resubmitted(self) -> bool:
        """
        Indica si el aprendiz puede realizar
        una nueva entrega.
        """

        return self.status in {
            EVIDENCE_STATUS_NOT_SUBMITTED,
            EVIDENCE_STATUS_REQUIRES_CORRECTION,
        }

    @property
    def is_finished(self) -> bool:
        """
        La evidencia termina su ciclo cuando
        queda aprobada.
        """

        return self.is_approved

    @property
    def is_not_submitted(self) -> bool:
        return (
            self.status
            == EVIDENCE_STATUS_NOT_SUBMITTED
        )

    # ------------------------------------------------------------------
    # Limpieza de aprobación y firma
    # ------------------------------------------------------------------

    def clear_signed_file(self) -> None:
        self.signed_file_name = None
        self.signed_file_path = None
        self.signed_at = None

    def clear_approval(self) -> None:
        self.approved_at = None
        self.approved_by_id = None

    def clear_review(self) -> None:
        self.reviewed_at = None
        self.reviewed_by = None

    # ------------------------------------------------------------------
    # Estados
    # ------------------------------------------------------------------

    def mark_not_submitted(self) -> None:
        """
        Regresa la entrega al estado inicial.
        """

        self.status = (
            EVIDENCE_STATUS_NOT_SUBMITTED
        )

        self.file_name = None
        self.file_path = None
        self.mime_type = None
        self.file_size_bytes = None
        self.uploaded_at = None

        latest = self.latest_attempt
        if latest is not None:
            latest.status = EVIDENCE_STATUS_NOT_SUBMITTED
            latest.file_name = None
            latest.file_path = None
            latest.mime_type = None
            latest.file_size_bytes = None
            latest.uploaded_at = None
            latest.reviewed_at = None
            latest.reviewed_by = None
            latest.approved_at = None
            latest.approved_by_id = None
            latest.signed_file_name = None
            latest.signed_file_path = None
            latest.signed_at = None

        self.clear_review()
        self.clear_signed_file()
        self.clear_approval()

    def mark_pending_review(self) -> None:
        """Coloca la evidencia en revisión desde una entrega válida."""
        if not self.has_file:
            raise ValueError(
                "No se puede poner en revisión una evidencia sin archivo."
            )
        if self.status not in {
            EVIDENCE_STATUS_NOT_SUBMITTED,
            EVIDENCE_STATUS_REQUIRES_CORRECTION,
        }:
            raise ValueError(
                "La evidencia no puede pasar a revisión desde su estado actual."
            )

        self.status = (
            EVIDENCE_STATUS_PENDING_REVIEW
        )

        self.reviewed_at = None
        self.reviewed_by = None

        self.clear_approval()
        self.clear_signed_file()

    def submit(
        self,
        file_name: str,
        file_path: str,
        mime_type: str | None = None,
        file_size_bytes: int | None = None,
        uploaded_at: datetime | None = None,
    ) -> None:
        """Crea un nuevo intento y lo convierte en la versión vigente."""
        if self.status not in {EVIDENCE_STATUS_NOT_SUBMITTED, EVIDENCE_STATUS_REQUIRES_CORRECTION}:
            raise ValueError("La evidencia no puede reenviarse desde su estado actual.")

        previous = self.latest_attempt
        attempt_number = (previous.attempt_number + 1) if previous else 1
        version_number = (previous.version_number + 1) if previous else 1
        attempt = EvidenceSubmissionAttempt(
            submission=self,
            attempt_number=attempt_number,
            version_number=version_number,
            status=EVIDENCE_STATUS_PENDING_REVIEW,
            file_name=_normalize_required_text(file_name, "nombre del archivo"),
            file_path=_normalize_required_text(file_path, "ruta del archivo"),
            mime_type=_normalize_mime_type(mime_type),
            file_size_bytes=_normalize_optional_int(file_size_bytes),
            uploaded_at=uploaded_at or _utcnow(),
        )
        self.status = EVIDENCE_STATUS_PENDING_REVIEW
        self.version_number = version_number
        self.attempt_number = attempt_number
        self.is_latest = True
        self.file_name = attempt.file_name
        self.file_path = attempt.file_path
        self.mime_type = attempt.mime_type
        self.file_size_bytes = attempt.file_size_bytes
        self.uploaded_at = attempt.uploaded_at
        self.reviewed_at = None
        self.reviewed_by = None
        self.approved_at = None
        self.approved_by_id = None
        self.clear_signed_file()

    def attach_signed_file(
        self,
        file_name: str,
        file_path: str,
        signed_at: datetime | None = None,
    ) -> None:
        """
        Asocia el PDF firmado a la entrega.

        El archivo original no se elimina ni se reemplaza
        desde el modelo.
        """

        if not self.is_pdf:
            raise ValueError(
                "Solo los archivos PDF pueden ser firmados."
            )

        self.signed_file_name = (
            _normalize_required_text(
                file_name,
                "nombre del archivo firmado",
            )
        )

        self.signed_file_path = (
            _normalize_required_text(
                file_path,
                "ruta del archivo firmado",
            )
        )

        self.signed_at = (
            signed_at
            or _utcnow()
        )
        latest = self.latest_attempt
        if latest is not None:
            latest.signed_file_name = self.signed_file_name
            latest.signed_file_path = self.signed_file_path
            latest.signed_at = self.signed_at

    def approve(
        self,
        approved_by_id: int,
        approved_at: datetime | None = None,
    ) -> None:
        """Aprueba exclusivamente una entrega pendiente de revisión."""
        if not self.has_file:
            raise ValueError(
                "No se puede aprobar una evidencia sin archivo."
            )
        if self.status != EVIDENCE_STATUS_PENDING_REVIEW:
            raise ValueError(
                "Solo una evidencia pendiente de revisión puede aprobarse."
            )

        self.status = (
            EVIDENCE_STATUS_APPROVED
        )

        self.approved_by_id = (
            _normalize_required_int(
                approved_by_id,
                "usuario aprobador",
            )
        )

        self.approved_at = (
            approved_at
            or _utcnow()
        )

        self.reviewed_by = (
            self.approved_by_id
        )

        self.reviewed_at = (
            self.approved_at
        )
        latest = self.latest_attempt
        if latest is not None:
            latest.status = self.status
            latest.approved_by_id = self.approved_by_id
            latest.approved_at = self.approved_at
            latest.reviewed_by = self.reviewed_by
            latest.reviewed_at = self.reviewed_at

    def request_revision(
        self,
        reviewed_by_id: int | None = None,
    ) -> None:
        """
        Solicita correcciones sobre una evidencia.

        Las observaciones se conservan para que el aprendiz
        pueda conocer qué debe corregir.
        """

        if not self.has_file:
            raise ValueError(
                "No se pueden solicitar correcciones sin una entrega."
            )
        if self.status != EVIDENCE_STATUS_PENDING_REVIEW:
            raise ValueError(
                "Solo una evidencia pendiente de revisión puede requerir corrección."
            )

        self.status = (
            EVIDENCE_STATUS_REQUIRES_CORRECTION
        )

        self.reviewed_at = _utcnow()
        if reviewed_by_id is not None:
            self.reviewed_by = _normalize_required_int(
                reviewed_by_id,
                "usuario revisor",
            )

        self.clear_approval()
        self.clear_signed_file()

        latest = self.latest_attempt
        if latest is not None:
            latest.status = self.status
            latest.reviewed_at = self.reviewed_at
            latest.reviewed_by = self.reviewed_by

    def add_observation(
        self,
        observation: str,
        author_id: int | None = None,
        *,
        is_correction_request: bool = False,
    ) -> "EvidenceComment":
        """Registra un comentario visible, opcionalmente marcado como solicitud de corrección."""
        comment = EvidenceComment(
            submission=self,
            attempt=self.latest_attempt,
            author_id=author_id,
            comment=_normalize_required_text(observation, "comentario"),
            is_internal=False,
            is_correction_request=bool(is_correction_request),
        )
        # Adjuntar explícitamente al agregado para que SQLAlchemy lo persista
        # mediante la relación comments/cascade al hacer commit. Crear el
        # objeto y devolverlo sin asociarlo a una colección dejaba el comentario
        # fuera del flush en ciertas rutas, aunque la notificación sí se creara.
        self.comments.append(comment)
        return comment

    def __str__(self) -> str:
        return (
            f"{self.apprentice_id} - "
            f"{self.activity_id}"
        )

    def __repr__(self) -> str:
        return (
            f"<EvidenceSubmission "
            f"id={getattr(self, 'id', None)!r} "
            f"activity_id={self.activity_id!r} "
            f"apprentice_id={self.apprentice_id!r} "
            f"status={self.status!r}>"
        )


# ==============================================================================
# MODELO DE INTENTO DE ENTREGA
# ==============================================================================

class EvidenceSubmissionAttempt(BaseModel):
    """Versión inmutable de una entrega concreta.

    EvidenceSubmission es el agregado; cada carga/reentrega genera un intento.
    El estado y los archivos históricos viven aquí para conservar trazabilidad.
    """

    __tablename__ = "evidence_submission_attempt"

    submission_id: Mapped[int] = mapped_column(
        ForeignKey("evidence_submission.id"),
        nullable=False,
        index=True,
    )
    attempt_number: Mapped[int] = mapped_column(Integer, nullable=False)
    version_number: Mapped[int] = mapped_column(Integer, nullable=False)
    status: Mapped[str] = mapped_column(
        String(40), nullable=False, default=EVIDENCE_STATUS_PENDING_REVIEW,
        server_default=EVIDENCE_STATUS_PENDING_REVIEW, index=True,
    )
    file_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    file_path: Mapped[str | None] = mapped_column(String(255), nullable=True)
    mime_type: Mapped[str | None] = mapped_column(String(120), nullable=True, index=True)
    file_size_bytes: Mapped[int | None] = mapped_column(Integer, nullable=True)
    uploaded_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    reviewed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    reviewed_by: Mapped[int | None] = mapped_column(ForeignKey("user.id"), nullable=True, index=True)
    approved_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    approved_by_id: Mapped[int | None] = mapped_column(ForeignKey("user.id"), nullable=True, index=True)
    signed_file_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    signed_file_path: Mapped[str | None] = mapped_column(String(255), nullable=True)
    signed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)

    submission = relationship("EvidenceSubmission", back_populates="attempts")
    reviewed_by_user = relationship("User", foreign_keys=[reviewed_by])
    approved_by_user = relationship("User", foreign_keys=[approved_by_id])

    @property
    def has_file(self) -> bool:
        return bool(self.file_path)

    @property
    def is_pdf(self) -> bool:
        return (self.mime_type or "").lower() == "application/pdf" or (self.file_name or "").lower().endswith(".pdf")



# ==============================================================================
# MODELO DE COMENTARIO
# ==============================================================================

class EvidenceComment(BaseModel):
    """
    Comentario o retroalimentación sobre una entrega.
    """

    __tablename__ = "evidence_comment"

    submission_id: Mapped[int] = mapped_column(
        ForeignKey("evidence_submission.id"),
        nullable=False,
        index=True,
    )

    attempt_id: Mapped[int | None] = mapped_column(
        ForeignKey("evidence_submission_attempt.id"),
        nullable=True,
        index=True,
    )

    author_id: Mapped[int | None] = mapped_column(
        ForeignKey("user.id"),
        nullable=True,
        index=True,
    )

    comment: Mapped[str] = mapped_column(
        Text,
        nullable=False,
    )

    is_internal: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=False,
        server_default=false(),
        index=True,
    )

    is_correction_request: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        default=False,
        server_default=false(),
        index=True,
    )

    submission = relationship(
        "EvidenceSubmission",
        back_populates="comments",
        lazy="select",
    )

    attempt = relationship(
        "EvidenceSubmissionAttempt",
        foreign_keys=[attempt_id],
        backref=backref(
            "comments",
            lazy="selectin",
        ),
    )

    author = relationship(
        "User",
        foreign_keys=[author_id],
        backref=backref(
            "evidence_comments",
            lazy="selectin",
        ),
    )

    @validates(
        "submission_id",
        "attempt_id",
        "author_id",
    )
    def validate_fk_ids(
        self,
        key: str,
        value: Any,
    ) -> int | None:
        if value is None:
            return None
        return _normalize_required_int(value, key)

    @validates("comment")
    def validate_comment(
        self,
        key: str,
        value: Any,
    ) -> str:
        return _normalize_required_text(
            value,
            "comentario",
        )

    @validates("is_internal")
    def validate_is_internal(
        self,
        key: str,
        value: Any,
    ) -> bool:
        return _normalize_bool(value)

    @validates("is_correction_request")
    def validate_is_correction_request(
        self,
        key: str,
        value: Any,
    ) -> bool:
        return _normalize_bool(value)

    @property
    def author_role_label(self) -> str:
        if not self.author:
            return "Sistema"
        role = getattr(self.author, "role", None)
        return {
            "APPRENTICE": "Aprendiz",
            "FOLLOW_UP_INSTRUCTOR": "Instructor de seguimiento",
            "LEAD_FOLLOW_UP_INSTRUCTOR": "Instructor de seguimiento líder",
            "CERTIFIER": "Certificador",
            "SUPPORT": "Soporte",
        }.get(role, role or "Usuario")

    @property
    def author_name(self) -> str:
        if not self.author:
            return ""

        return (
            getattr(
                self.author,
                "display_name",
                None,
            )
            or getattr(
                self.author,
                "full_name",
                None,
            )
            or getattr(
                self.author,
                "login_identifier",
                "",
            )
        )

    def __str__(self) -> str:
        return self.comment[:80]

    def __repr__(self) -> str:
        return (
            f"<EvidenceComment "
            f"id={getattr(self, 'id', None)!r} "
            f"submission_id={self.submission_id!r} "
            f"author_id={self.author_id!r}>"
        )


# ==============================================================================
# EXPORTACIONES
# ==============================================================================

__all__ = [
    # Estados
    "EVIDENCE_STATUS_NOT_SUBMITTED",
    "EVIDENCE_STATUS_PENDING_REVIEW",
    "EVIDENCE_STATUS_REQUIRES_CORRECTION",
    "EVIDENCE_STATUS_APPROVED",
    "EVIDENCE_STATUSES",
    "EVIDENCE_STATUS_LABELS",
    "EVIDENCE_STATUS_COLORS",
    "EVIDENCE_STATUS_ORDER",

    # Origen
    "EVIDENCE_ACTIVITY_ORIGIN_TEMPLATE",
    "EVIDENCE_ACTIVITY_ORIGIN_CUSTOM",
    "EVIDENCE_ACTIVITY_ORIGINS",
    "EVIDENCE_ACTIVITY_ORIGIN_LABELS",

    # Catálogos

    # Modelos
    "EvidenceCategory",
    "EvidenceTemplate",
    "EvidenceActivity",
    "EvidenceSubmission",
    "EvidenceComment",
]