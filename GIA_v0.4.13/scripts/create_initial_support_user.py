"""Create the initial SUPPORT account for a fresh local GIA installation.

Usage:
    python scripts/create_initial_support_user.py

The script is intentionally interactive for the password so credentials are not
stored in source code, shell history, or .env files.
"""
from __future__ import annotations

import getpass
from datetime import datetime, timezone
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app import create_app  # noqa: E402
from catalogs.user import UserDocumentType, UserRole, UserStatus  # noqa: E402
from extensions import db  # noqa: E402
from models import User  # noqa: E402


def prompt(label: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{label}{suffix}: ").strip()
    return value or (default or "")


def main() -> int:
    app = create_app()
    with app.app_context():
        print("GIA - creación de usuario inicial SUPPORT")
        print("La contraseña se solicitará de forma oculta y no se almacenará en el código.\n")

        email = prompt("Correo", "soporte@gia.local").lower()
        document_number = prompt("Número de documento", "GIA-SUPPORT-001").upper()
        first_names = prompt("Nombres", "Soporte")
        last_names = prompt("Apellidos", "GIA")

        if not email or "@" not in email:
            print("ERROR: el correo no parece válido.")
            return 2
        if not document_number:
            print("ERROR: el documento es obligatorio.")
            return 2
        if not first_names or not last_names:
            print("ERROR: nombres y apellidos son obligatorios.")
            return 2

        existing = User.query.filter(
            (User.email == email) | (User.document_number == document_number)
        ).first()
        if existing is not None:
            print(
                "ERROR: ya existe un usuario con ese correo o número de documento "
                f"(id={existing.id}, email={existing.email!r}, role={existing.role!r})."
            )
            print("No se modificó ningún registro.")
            return 3

        password = getpass.getpass("Contraseña inicial (mínimo 8 caracteres): ")
        confirm = getpass.getpass("Confirmar contraseña: ")

        if len(password) < 8:
            print("ERROR: la contraseña debe tener al menos 8 caracteres.")
            return 2
        if password != confirm:
            print("ERROR: las contraseñas no coinciden.")
            return 2

        user = User(
            document_type=UserDocumentType.NATIONAL_ID.value,
            document_number=document_number,
            first_names=first_names,
            last_names=last_names,
            email=email,
            role=UserRole.SUPPORT.value,
            status=UserStatus.ACTIVE.value,
            password_hash="TEMP",
            created_at=datetime.now(timezone.utc),
            updated_at=datetime.now(timezone.utc),
        )
        user.set_password(password)

        db.session.add(user)
        db.session.commit()

        print("\nUSER_CREATED=PASS")
        print(f"USER_ID={user.id}")
        print(f"USER_EMAIL={user.email}")
        print(f"USER_ROLE={user.role}")
        print(f"USER_STATUS={user.status}")
        print("Ya puedes iniciar sesión en GIA con el correo y contraseña introducidos.")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
