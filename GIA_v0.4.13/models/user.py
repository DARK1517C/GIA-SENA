"""
models/user.py

Modelo de usuario del sistema GIA.

Este modelo concentra la identidad, autenticación y rol funcional
de cada usuario del sistema. Los permisos concretos se resolverán
posteriormente mediante servicios y lógica de autorización.

Responsabilidades:
- Identidad básica del usuario.
- Autenticación (contraseña).
- Rol funcional del sistema.
- Estado de la cuenta.
- Datos de firma para aprobación de PDFs.
- Utilidades de uso frecuente en la aplicación.

No contiene lógica de negocio de otros dominios.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, TypeVar

from flask_login import UserMixin
from sqlalchemy import DateTime, String
from sqlalchemy.orm import Mapped, mapped_column, validates
from werkzeug.security import check_password_hash, generate_password_hash

from catalogs.common import CatalogEnum, normalize_spaces
from catalogs.exceptions import InvalidCatalogValueError
from catalogs.user import UserDocumentType, UserRole, UserStatus
from catalogs.validation import CatalogValidation

from .base import BaseModel

CatalogEnumT = TypeVar("CatalogEnumT", bound=CatalogEnum)


def _coerce_catalog_value(
    catalog: type[CatalogEnumT],
    value: str | CatalogEnumT | None,
) -> str:
    """
    Convierte un valor de catálogo a su valor canónico en cadena.
    """
    if isinstance(value, catalog):
        return value.value

    normalized = CatalogValidation.validate_required(catalog, value)

    if normalized is None:
        raise InvalidCatalogValueError(
            f"El valor no es válido para el catálogo {catalog.__name__}."
        )

    return normalized


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
    Normaliza un texto opcional.
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


class User(UserMixin, BaseModel):
    """
    Usuario del sistema.

    Un usuario puede representar:
    - Aprendiz
    - Instructor de seguimiento
    - Instructor de seguimiento líder
    - Certificador
    - Administrativo del centro
    - Soporte
    """

    __tablename__ = "user"

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
        index=True,
    )

    last_names: Mapped[str] = mapped_column(
        String(120),
        nullable=False,
        index=True,
    )

    email: Mapped[str | None] = mapped_column(
        String(255),
        unique=True,
        nullable=True,
        index=True,
    )

    phone: Mapped[str | None] = mapped_column(
        String(30),
        nullable=True,
    )

    role: Mapped[str] = mapped_column(
        String(60),
        nullable=False,
        index=True,
    )

    status: Mapped[str] = mapped_column(
        String(30),
        nullable=False,
        default=UserStatus.ACTIVE.value,
        server_default=UserStatus.ACTIVE.value,
        index=True,
    )

    password_hash: Mapped[str] = mapped_column(
        String(255),
        nullable=False,
    )

    signature_file_name: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    signature_file_path: Mapped[str | None] = mapped_column(
        String(255),
        nullable=True,
    )

    signature_updated_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    last_login_at: Mapped[datetime | None] = mapped_column(
        DateTime(timezone=True),
        nullable=True,
    )

    # ---------------------------------------------------------------------
    # VALIDACIONES
    # ---------------------------------------------------------------------

    @validates("document_type")
    def validate_document_type(self, key: str, value: Any) -> str:
        return _coerce_catalog_value(UserDocumentType, value)

    @validates("document_number")
    def validate_document_number(self, key: str, value: Any) -> str:
        normalized = _normalize_required_text(value, "número de documento")
        return normalized.replace(" ", "").upper()

    @validates("first_names")
    def validate_first_names(self, key: str, value: Any) -> str:
        return _normalize_required_text(value, "nombres")

    @validates("last_names")
    def validate_last_names(self, key: str, value: Any) -> str:
        return _normalize_required_text(value, "apellidos")

    @validates("email")
    def validate_email(self, key: str, value: Any) -> str | None:
        return _normalize_optional_email(value)

    @validates("phone")
    def validate_phone(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    @validates("role")
    def validate_role(self, key: str, value: Any) -> str:
        return _coerce_catalog_value(UserRole, value)

    @validates("status")
    def validate_status(self, key: str, value: Any) -> str:
        return _coerce_catalog_value(UserStatus, value)

    @validates("password_hash")
    def validate_password_hash(self, key: str, value: Any) -> str:
        if value is None:
            raise ValueError("La contraseña no puede estar vacía.")

        normalized = str(value).strip()
        if not normalized:
            raise ValueError("La contraseña no puede estar vacía.")

        return normalized

    @validates("signature_file_name", "signature_file_path")
    def validate_signature_fields(self, key: str, value: Any) -> str | None:
        return _normalize_optional_text(value)

    # ---------------------------------------------------------------------
    # AUTENTICACIÓN
    # ---------------------------------------------------------------------

    def set_password(self, password: str) -> None:
        """
        Genera y almacena el hash de la contraseña.
        """
        if password is None or not str(password).strip():
            raise ValueError("La contraseña no puede estar vacía.")

        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
        """
        Verifica una contraseña en texto plano contra el hash almacenado.
        """
        if not self.password_hash:
            return False

        return check_password_hash(self.password_hash, password)

    def touch_last_login(self) -> None:
        """
        Actualiza la fecha del último acceso.
        """
        self.last_login_at = datetime.now(timezone.utc)

    # ---------------------------------------------------------------------
    # FIRMA
    # ---------------------------------------------------------------------

    def set_signature(self, file_name: str, file_path: str) -> None:
        """
        Asocia una firma al usuario.
        """
        self.signature_file_name = _normalize_required_text(file_name, "nombre de la firma")
        self.signature_file_path = _normalize_required_text(file_path, "ruta de la firma")
        self.signature_updated_at = datetime.now(timezone.utc)

    def clear_signature(self) -> None:
        """
        Elimina la firma asociada al usuario.
        """
        self.signature_file_name = None
        self.signature_file_path = None
        self.signature_updated_at = None

    @property
    def has_signature(self) -> bool:
        """
        Indica si el usuario tiene firma configurada.
        """
        return bool(self.signature_file_path)

    # ---------------------------------------------------------------------
    # PROPIEDADES DE CATÁLOGO
    # ---------------------------------------------------------------------

    @property
    def document_type_enum(self) -> UserDocumentType:
        return UserDocumentType(self.document_type)

    @property
    def role_enum(self) -> UserRole:
        return UserRole(self.role)

    @property
    def status_enum(self) -> UserStatus:
        return UserStatus(self.status)

    # ---------------------------------------------------------------------
    # ESTADO DE CUENTA
    # ---------------------------------------------------------------------

    @property
    def is_active(self) -> bool:
        """
        Compatibilidad con Flask-Login.
        """
        return self.status == UserStatus.ACTIVE.value

    @is_active.setter
    def is_active(self, value: bool) -> None:
        self.status = UserStatus.ACTIVE.value if value else UserStatus.INACTIVE.value

    def activate(self) -> None:
        self.status = UserStatus.ACTIVE.value

    def deactivate(self) -> None:
        self.status = UserStatus.INACTIVE.value

    def suspend(self) -> None:
        self.status = UserStatus.SUSPENDED.value

    def set_pending(self) -> None:
        self.status = UserStatus.PENDING.value

    # ---------------------------------------------------------------------
    # ROLE CHECKS
    # ---------------------------------------------------------------------

    @property
    def is_apprentice(self) -> bool:
        return self.role == UserRole.APPRENTICE.value

    @property
    def is_follow_up_instructor(self) -> bool:
        return self.role == UserRole.FOLLOW_UP_INSTRUCTOR.value

    @property
    def is_lead_follow_up_instructor(self) -> bool:
        return self.role == UserRole.LEAD_FOLLOW_UP_INSTRUCTOR.value

    @property
    def is_certifier(self) -> bool:
        return self.role == UserRole.CERTIFIER.value

    @property
    def is_center_staff(self) -> bool:
        return self.role == UserRole.CENTER_STAFF.value

    @property
    def is_support(self) -> bool:
        return self.role == UserRole.SUPPORT.value

    @property
    def is_instructor(self) -> bool:
        return self.is_follow_up_instructor or self.is_lead_follow_up_instructor

    @property
    def is_admin(self) -> bool:
        return self.is_center_staff or self.is_support

    # ---------------------------------------------------------------------
    # CAPACIDADES DE DOMINIO
    # ---------------------------------------------------------------------

    @property
    def can_sign_pdf(self) -> bool:
        """
        Solo los instructores pueden firmar PDFs dentro de la plataforma.
        """
        return self.is_instructor

    @property
    def can_approve_evidences(self) -> bool:
        """Aprobación según la política canónica de permisos."""
        return self.is_instructor or self.is_certifier or self.is_support

    @property
    def can_manage_apprentices(self) -> bool:
        return self.is_instructor or self.is_support

    @property
    def can_manage_groups(self) -> bool:
        return self.is_instructor or self.is_support

    @property
    def can_import_excel(self) -> bool:
        return self.is_instructor or self.is_support

    @property
    def can_export_reports(self) -> bool:
        return self.is_instructor or self.is_admin or self.is_support

    @property
    def can_manage_users(self) -> bool:
        return self.is_support

    @property
    def can_manage_catalogs(self) -> bool:
        return self.is_support

    @property
    def can_review_certification(self) -> bool:
        return self.is_certifier or self.is_support

    @property
    def can_view_global_statistics(self) -> bool:
        return self.is_lead_follow_up_instructor or self.is_admin or self.is_support

    # ---------------------------------------------------------------------
    # AYUDAS DE USO FRECUENTE
    # ---------------------------------------------------------------------

    @property
    def full_name(self) -> str:
        return f"{self.first_names} {self.last_names}".strip()

    @property
    def display_name(self) -> str:
        return self.full_name

    @property
    def login_identifier(self) -> str:
        """
        Identificador de acceso recomendado.
        Prioriza el correo si existe; de lo contrario, el documento.
        """
        return self.email or self.document_number

    @property
    def avatar_initial(self) -> str:
        """
        Primera letra útil para avatares.
        """
        text = self.first_names.strip() if self.first_names else self.document_number.strip()
        return text[:1].upper() if text else ""

    # ---------------------------------------------------------------------
    # SERIALIZACIÓN
    # ---------------------------------------------------------------------

    def to_dict(self) -> dict[str, Any]:
        """
        Serializa el usuario excluyendo el hash de contraseña.
        """
        data = super().to_dict()
        data.pop("password_hash", None)
        return data

    # ---------------------------------------------------------------------
    # REPRESENTACIÓN
    # ---------------------------------------------------------------------

    def __str__(self) -> str:
        return self.full_name

    def __repr__(self) -> str:
        return (
            f"<User id={getattr(self, 'id', None)!r} "
            f"full_name={self.full_name!r} "
            f"role={self.role!r}>"
        )