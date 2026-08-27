"""
models/apprentice.py

Modelo de aprendiz del sistema GIA.

Este modelo representa la información académica y operativa
del aprendiz dentro de la plataforma, enlazando su identidad
de usuario con el grupo y con los datos asociados a su etapa productiva.

Responsabilidades:
- Relacionar el aprendiz con el usuario que lo representa en la plataforma.
- Relacionar el aprendiz con el usuario creador del registro.
- Relacionar el aprendiz con su grupo.
- Almacenar datos académicos, de etapa productiva y de empresa.
- Validar campos estandarizados mediante catálogos.

No contiene lógica de evidencias ni de certificación.
"""

from __future__ import annotations

import re
from typing import Any, TypeVar

from sqlalchemy import ForeignKey, String, Text
from sqlalchemy.orm import Mapped, backref, mapped_column, relationship, validates

from catalogs.apprentice import EpModality, SofiaStatus
from catalogs.common import CatalogEnum, normalize_spaces
from catalogs.common_catalogs import Gender, ProgramLevel, YesNo
from catalogs.common_catalogs import DocumentType as _LegacyDocumentTypeAlias  # noqa: F401
from catalogs.exceptions import InvalidCatalogValueError
from catalogs.user import UserDocumentType
from catalogs.validation import CatalogValidation

from .base import BaseModel

CatalogEnumT = TypeVar("CatalogEnumT", bound=CatalogEnum)


def _coerce_required_catalog_value(
    catalog: type[CatalogEnumT],
    value: str | CatalogEnumT | None,
) -> str:
    """
    Convierte un valor de catálogo obligatorio a su valor canónico.
    """
    if isinstance(value, catalog):
        return value.value

    normalized = CatalogValidation.validate_required(catalog, value)

    if normalized is None:
        raise InvalidCatalogValueError(
            f"El valor no es válido para el catálogo {catalog.__name__}."
        )

    return normalized


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


def _normalize_optional_phone(value: Any) -> str | None:
    """
    Normaliza un número telefónico opcional.
    """
    if value is None:
        return None

    normalized = normalize_spaces(str(value))
    return normalized or None


def _normalize_optional_email(value: Any) -> str | None:
    """
    Normaliza un correo electrónico opcional.
    """
    if value is None:
        return None

    normalized = normalize_spaces(str(value)).lower()
    if not normalized:
        return None

    if "@" not in normalized or normalized.startswith("@") or normalized.endswith("@"):
        raise ValueError("El correo electrónico no es válido.")

    return normalized


class Apprentice(BaseModel):
    """
    Perfil académico y operativo de un aprendiz.
    """

    __tablename__ = "apprentice"

    created_by: Mapped[int] = mapped_column(
        ForeignKey("user.id"),
        nullable=False,
        index=True,
    )

    student_user_id: Mapped[int | None] = mapped_column(
        ForeignKey("user.id"),
        nullable=True,
        index=True,
    )

    group_id: Mapped[int | None] = mapped_column(
        ForeignKey("training_group.id"),
        nullable=True,
        index=True,
    )

    group_number: Mapped[str] = mapped_column(
        String(30),
        nullable=False,
        index=True,
    )

    document_type: Mapped[str] = mapped_column(
        String(30),
        nullable=False,
        index=True,
    )

    document_number: Mapped[str] = mapped_column(
        String(30),
        unique=True,
        nullable=False,
        index=True,
    )

    first_names: Mapped[str] = mapped_column(
        String(120),
        nullable=False,
    )

    last_names: Mapped[str] = mapped_column(
        String(120),
        nullable=False,
    )

    gender: Mapped[str | None] = mapped_column(
        String(20),
        nullable=True,
        index=True,
    )

    phone: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    email: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
        index=True,
    )

    municipality_origin: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
        index=True,
    )

    program_name: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
        index=True,
    )

    program_level: Mapped[str | None] = mapped_column(
        String(80),
        nullable=True,
        index=True,
    )

    group_validity: Mapped[str | None] = mapped_column(
        String(80),
        nullable=True,
    )

    lead_instructor: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    followup_instructor: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    followup_instructor_email: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    ep_modality: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
        index=True,
    )

    sofia_status: Mapped[str | None] = mapped_column(
        String(80),
        nullable=True,
        index=True,
    )

    practice_start_date: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    practice_end_date: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    followup_moment1_start: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    followup_moment1_end: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    followup_moment2_start: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    followup_moment2_end: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    followup_moment3_start: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    followup_moment3_end: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    followup_moment4_start: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    followup_moment4_end: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    company_name: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    company_municipality: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
    )

    company_address: Mapped[str | None] = mapped_column(
        String(180),
        nullable=True,
    )

    coformador_name: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    coformador_email: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    coformador_phone: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    arl_responsible: Mapped[str | None] = mapped_column(
        String(150),
        nullable=True,
    )

    continues_company: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    individual_management: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    followup_moments: Mapped[str | None] = mapped_column(
        String(200),
        nullable=True,
    )

    evaluation_date: Mapped[str | None] = mapped_column(
        String(40),
        nullable=True,
    )

    english_results: Mapped[str | None] = mapped_column(
        String(120),
        nullable=True,
    )

    # ------------------------------------------------------------------
    # RELACIONES
    # ------------------------------------------------------------------

    created_by_user = relationship(
        "User",
        foreign_keys=[created_by],
        backref=backref("created_apprentice_records", lazy="selectin"),
    )

    student_user = relationship(
        "User",
        foreign_keys=[student_user_id],
        backref=backref("apprentice_profile", uselist=False),
        lazy="select",
    )

    group = relationship(
        "TrainingGroup",
        foreign_keys=[group_id],
        back_populates="apprentices",
        lazy="joined",
    )

    evidence_submissions = relationship(
        "EvidenceSubmission",
        back_populates="apprentice",
        lazy="select",
        cascade="all, delete-orphan",
    )

    # ------------------------------------------------------------------
    # VALIDACIONES
    # ------------------------------------------------------------------

    @validates("group_number")
    def validate_group_number(self, key: str, value: Any) -> str:
        return _normalize_required_text(value, "número de grupo")

    @validates("document_type")
    def validate_document_type(self, key: str, value: Any) -> str:
        return _coerce_required_catalog_value(UserDocumentType, value)

    @validates("document_number")
    def validate_document_number(self, key: str, value: Any) -> str:
        normalized = _normalize_required_text(value, "número de documento")
        return re.sub(r"\s+", "", normalized).upper()

    @validates("first_names")
    def validate_first_names(self, key: str, value: Any) -> str:
        return _normalize_required_text(value, "nombres")

    @validates("last_names")
    def validate_last_names(self, key: str, value: Any) -> str:
        return _normalize_required_text(value, "apellidos")

    @validates("gender")
    def validate_gender(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(Gender, value)

    @validates("phone")
    def validate_phone(self, key: str, value: Any) -> str | None:
        return _normalize_optional_phone(value)

    @validates("email")
    def validate_email(self, key: str, value: Any) -> str | None:
        return _normalize_optional_email(value)

    @validates("municipality_origin")
    def validate_municipality_origin(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("program_name")
    def validate_program_name(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("program_level")
    def validate_program_level(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(ProgramLevel, value)

    @validates("group_validity")
    def validate_group_validity(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("lead_instructor")
    def validate_lead_instructor(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("followup_instructor")
    def validate_followup_instructor(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("followup_instructor_email")
    def validate_followup_instructor_email(
        self,
        key: str,
        value: Any,
    ) -> str | None:
        return _normalize_optional_email(value)

    @validates("ep_modality")
    def validate_ep_modality(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(EpModality, value)

    @validates("sofia_status")
    def validate_sofia_status(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(SofiaStatus, value)

    @validates("practice_start_date", "practice_end_date")
    def validate_practice_dates(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates(
        "followup_moment1_start",
        "followup_moment1_end",
        "followup_moment2_start",
        "followup_moment2_end",
        "followup_moment3_start",
        "followup_moment3_end",
        "followup_moment4_start",
        "followup_moment4_end",
    )
    def validate_followup_moment_dates(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("company_name")
    def validate_company_name(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("company_municipality")
    def validate_company_municipality(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("company_address")
    def validate_company_address(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("coformador_name")
    def validate_coformador_name(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("coformador_email")
    def validate_coformador_email(self, key: str, value: Any) -> str | None:
        return _normalize_optional_email(value)

    @validates("coformador_phone")
    def validate_coformador_phone(self, key: str, value: Any) -> str | None:
        return _normalize_optional_phone(value)

    @validates("arl_responsible")
    def validate_arl_responsible(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("continues_company")
    def validate_continues_company(self, key: str, value: Any) -> str | None:
        return _coerce_optional_catalog_value(YesNo, value)

    @validates("individual_management")
    def validate_individual_management(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("followup_moments")
    def validate_followup_moments(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("evaluation_date")
    def validate_evaluation_date(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("english_results")
    def validate_english_results(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    # ------------------------------------------------------------------
    # PROPIEDADES
    # ------------------------------------------------------------------

    @property
    def full_name(self) -> str:
        return f"{self.first_names} {self.last_names}".strip()

    @property
    def display_name(self) -> str:
        return self.full_name

    @property
    def document_type_enum(self) -> UserDocumentType:
        return UserDocumentType(self.document_type)

    @property
    def gender_enum(self) -> Gender | None:
        return Gender(self.gender) if self.gender else None

    @property
    def program_level_enum(self) -> ProgramLevel | None:
        return ProgramLevel(self.program_level) if self.program_level else None

    @property
    def ep_modality_enum(self) -> EpModality | None:
        return EpModality(self.ep_modality) if self.ep_modality else None

    @property
    def sofia_status_enum(self) -> SofiaStatus | None:
        return SofiaStatus(self.sofia_status) if self.sofia_status else None

    @property
    def continues_company_enum(self) -> YesNo | None:
        return YesNo(self.continues_company) if self.continues_company else None

    @property
    def has_group(self) -> bool:
        return self.group_id is not None

    @property
    def has_student_user(self) -> bool:
        return self.student_user_id is not None

    @property
    def has_company(self) -> bool:
        return bool(self.company_name)

    @property
    def is_in_productive_stage(self) -> bool:
        return self.ep_modality is not None

    @property
    def is_certified(self) -> bool:
        return self.sofia_status == SofiaStatus.CERTIFICADO.value

    # ------------------------------------------------------------------
    # REPRESENTACIÓN
    # ------------------------------------------------------------------

    def __str__(self) -> str:
        return self.full_name or self.document_number

    def __repr__(self) -> str:
        return (
            f"<Apprentice id={getattr(self, 'id', None)!r} "
            f"document_number={self.document_number!r} "
            f"full_name={self.full_name!r}>"
        )