"""
models/training_group.py

Modelo de grupo de formación del sistema GIA.

Este modelo representa la información central de una ficha o grupo,
incluyendo su relación con el usuario creador, el programa, los
instructores asociados, el municipio, el estado académico y las
relaciones con aprendices y actividades de evidencia.

Responsabilidades:
- Relacionar el grupo con el usuario creador.
- Almacenar la información general de la ficha.
- Validar campos estandarizados mediante catálogos.
- Exponer relaciones con aprendices y actividades de evidencia.

No contiene lógica de evidencias ni de certificación.
"""

from __future__ import annotations

from typing import Any, TypeVar

from sqlalchemy import ForeignKey, String, Text
from sqlalchemy.orm import Mapped, backref, mapped_column, relationship, validates

from catalogs.common import CatalogEnum, normalize_spaces
from catalogs.common_catalogs import ProgramLevel
from catalogs.exceptions import InvalidCatalogValueError
from catalogs.training_group import GroupMunicipality, GroupModality, GroupStatus
from catalogs.validation import CatalogValidation

from .base import BaseModel

CatalogEnumT = TypeVar("CatalogEnumT", bound=CatalogEnum)


def _coerce_optional_catalog_value(
    catalog: type[CatalogEnumT],
    value: str | CatalogEnumT | None,
) -> str | None:
    """
    Convierte un valor de catálogo opcional a su valor canónico.
    """
    if value is None:
        return None

    if isinstance(value, catalog):
        return value.value

    return CatalogValidation.validate_optional(catalog, value)


def _normalize_required_text(value: Any, field_name: str) -> str:
    """
    Normaliza un texto obligatorio.
    """
    if value is None:
        raise ValueError(f"El campo {field_name} es obligatorio.")

    normalized = normalize_spaces(str(value))

    if not normalized:
        raise ValueError(f"El campo {field_name} es obligatorio.")

    return normalized


def _normalize_optional_text(value: Any) -> str | None:
    """
    Normaliza textos opcionales conservando el contenido original
    salvo espacios sobrantes.
    """
    if value is None:
        return None

    normalized = normalize_spaces(str(value))
    return normalized or None


class TrainingGroup(BaseModel):
    """
    Grupo de formación (ficha) del sistema GIA.
    """

    __tablename__ = "training_group"

    created_by: Mapped[int] = mapped_column(
        ForeignKey("user.id"),
        nullable=False,
        index=True,
    )

    group_number: Mapped[str] = mapped_column(
        String(30),
        unique=True,
        nullable=False,
        index=True,
    )

    program_name: Mapped[str] = mapped_column(
        String(150),
        nullable=False,
        index=True,
    )

    lead_instructor: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    followup_instructor: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    municipality: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
        index=True,
    )

    program_level: Mapped[str | None] = mapped_column(
        String(80),
        nullable=True,
        index=True,
    )

    modality: Mapped[str | None] = mapped_column(
        String(80),
        nullable=True,
        index=True,
    )

    sofia_group_status: Mapped[str | None] = mapped_column(
        String(80),
        nullable=True,
        index=True,
    )

    group_validity: Mapped[str | None] = mapped_column(
        String(80),
        nullable=True,
    )

    group_start_date: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    training_end_date: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    ep_start_date: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    apprentices_statistics: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
    )

    apprentices_training: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    apprentices_enabled: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    apprentices_rap_pending: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    apprentices_practice: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    apprentices_without_alternative: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    apprentices_certified: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    productive_modalities: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
    )

    learning_contract: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    internship: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    productive_project: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    employment_link: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    # ------------------------------------------------------------------
    # RELACIONES
    # ------------------------------------------------------------------

    created_by_user = relationship(
        "User",
        foreign_keys=[created_by],
        backref=backref("created_training_groups", lazy="selectin"),
    )

    apprentices = relationship(
        "Apprentice",
        back_populates="group",
        lazy="selectin",
    )

    evidence_activities = relationship(
        "EvidenceActivity",
        back_populates="group",
        lazy="selectin",
        cascade="all, delete-orphan",
    )

    # ------------------------------------------------------------------
    # VALIDACIONES
    # ------------------------------------------------------------------

    @validates("group_number")
    def validate_group_number(self, key: str, value: Any) -> str:
        return _normalize_required_text(value, "número de ficha")

    @validates("program_name")
    def validate_program_name(self, key: str, value: Any) -> str:
        return _normalize_required_text(value, "nombre del programa")

    @validates("lead_instructor")
    def validate_lead_instructor(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("followup_instructor")
    def validate_followup_instructor(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("municipality")
    def validate_municipality(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(GroupMunicipality, value)

    @validates("program_level")
    def validate_program_level(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(ProgramLevel, value)

    @validates("modality")
    def validate_modality(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(GroupModality, value)

    @validates("sofia_group_status")
    def validate_sofia_group_status(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(GroupStatus, value)

    @validates(
        "group_validity",
        "group_start_date",
        "training_end_date",
        "ep_start_date",
        "apprentices_statistics",
        "apprentices_training",
        "apprentices_enabled",
        "apprentices_rap_pending",
        "apprentices_practice",
        "apprentices_without_alternative",
        "apprentices_certified",
        "productive_modalities",
        "learning_contract",
        "internship",
        "productive_project",
        "employment_link",
    )
    def validate_optional_text_fields(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    # ------------------------------------------------------------------
    # PROPIEDADES
    # ------------------------------------------------------------------

    @property
    def display_name(self) -> str:
        """
        Nombre visible del grupo.
        """
        return self.group_number

    @property
    def title(self) -> str:
        """
        Título descriptivo del grupo.
        """
        if self.program_name:
            return f"{self.group_number} - {self.program_name}"
        return self.group_number

    @property
    def program_level_enum(self) -> ProgramLevel | None:
        return ProgramLevel(self.program_level) if self.program_level else None

    @property
    def modality_enum(self) -> GroupModality | None:
        return GroupModality(self.modality) if self.modality else None

    @property
    def sofia_group_status_enum(self) -> GroupStatus | None:
        return GroupStatus(self.sofia_group_status) if self.sofia_group_status else None

    @property
    def municipality_enum(self) -> GroupMunicipality | None:
        return GroupMunicipality(self.municipality) if self.municipality else None

    @property
    def has_apprentices(self) -> bool:
        return bool(self.apprentices)

    @property
    def apprentices_count(self) -> int:
        return len(self.apprentices or [])

    @property
    def has_evidence_activities(self) -> bool:
        return bool(self.evidence_activities)

    @property
    def is_active(self) -> bool:
        """
        Indica si la ficha se encuentra en ejecución.
        """
        return self.sofia_group_status == GroupStatus.EN_EJECUCION.value

    # ------------------------------------------------------------------
    # UTILIDADES DE REPORTE
    # ------------------------------------------------------------------

    @classmethod
    def record_fields(cls) -> list[tuple[str, str]]:
        """
        Devuelve la lista de campos usados en el récord de fichas.
        """
        return [
            ("consecutive", "CONSECUTIVO"),
            ("group_number", "N° DE FICHA"),
            ("lead_instructor", "NOMBRE DE INSTRUCTOR(A) LÍDER DE LA FICHA"),
            ("followup_instructor", "NOMBRE DE INSTRUCTOR(A) DE SEGUIMIENTO ETAPA PRODUCTIVA (EP)"),
            ("program_name", "NOMBRE DEL PROGRAMA DE FORMACIÓN"),
            ("municipality", "MUNICIPIO"),
            ("program_level", "NIVEL DE PROGRAMA"),
            ("modality", "MODALIDAD"),
            ("sofia_group_status", "ESTADO DE LA FICHA EN SOFÍA PLUS"),
            ("group_start_date", "FECHA INICIO DE LA FICHA EN SOFÍA PLUS"),
            ("training_end_date", "FECHA FIN DE LA FORMACIÓN EN SOFÍA PLUS"),
            ("ep_start_date", "FECHA INICIO DE ETAPA PRODUCTIVA"),
            ("group_validity", "VIGENCIA DE LA FICHA"),
            ("apprentices_training", "APRENDICES EN FORMACIÓN"),
            ("apprentices_enabled", "APRENDICES HABILITADOS PARA INICIAR ETAPA PRODUCTIVA"),
            ("apprentices_rap_pending", "APRENDICES QUE DEBEN RAP"),
            ("apprentices_practice", "APRENDICES EN PRÁCTICA"),
            ("learning_contract", "CONTRATO DE APRENDIZAJE"),
            ("internship", "VÍNCULO FORMATIVO"),
            ("productive_project", "PROYECTO PRODUCTIVO"),
            ("apprentices_without_alternative", "APRENDICES SIN ALTERNATIVA DE PRÁCTICA"),
            ("apprentices_certified", "APRENDICES CERTIFICADOS"),
        ]

    # ------------------------------------------------------------------
    # REPRESENTACIÓN
    # ------------------------------------------------------------------

    def __str__(self) -> str:
        return self.title

    def __repr__(self) -> str:
        return (
            f"<TrainingGroup id={getattr(self, 'id', None)!r} "
            f"group_number={self.group_number!r} "
            f"program_name={self.program_name!r}>"
        )