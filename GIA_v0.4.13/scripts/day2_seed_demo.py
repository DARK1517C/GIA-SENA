from __future__ import annotations

import sys
from pathlib import Path

import getpass
from datetime import date, timedelta

from werkzeug.security import generate_password_hash

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / ".env", override=False)

from app import create_app
from extensions import db
from models import Apprentice, TrainingGroup, User
from catalogs.user import UserDocumentType, UserRole, UserStatus
from catalogs.common_catalogs import ProgramLevel
from catalogs.training_group import GroupModality, GroupStatus
from catalogs.apprentice import SofiaStatus, EpModality
from services.evidence_service import ensure_submissions_for_apprentice, ensure_template_activities_for_group


def prompt_password(label: str) -> str:
    while True:
        value = getpass.getpass(label)
        if len(value) < 8:
            print("La contraseña debe tener al menos 8 caracteres.")
            continue
        confirm = getpass.getpass("Confirmar contraseña: ")
        if value != confirm:
            print("Las contraseñas no coinciden.")
            continue
        return value


def get_or_create_user(email: str, document_number: str, first_names: str, last_names: str, role: str, password: str) -> User:
    user = User.query.filter_by(email=email).first()
    if user:
        user.role = role
        user.status = UserStatus.ACTIVE.value
        if not user.password_hash:
            user.password_hash = generate_password_hash(password)
        return user

    user = User(
        document_type=UserDocumentType.NATIONAL_ID.value,
        document_number=document_number,
        first_names=first_names,
        last_names=last_names,
        email=email,
        role=role,
        status=UserStatus.ACTIVE.value,
        password_hash=generate_password_hash(password),
    )
    db.session.add(user)
    db.session.flush()
    return user


def main() -> int:
    app = create_app()
    with app.app_context():
        support = User.query.filter_by(role=UserRole.SUPPORT.value).order_by(User.id).first()
        if support is None:
            print("ERROR: primero crea el usuario SUPPORT con scripts/create_initial_support_user.py")
            return 2

        instructor_password = prompt_password("Contraseña del instructor de prueba: ")
        apprentice_password = prompt_password("Contraseña del aprendiz de prueba: ")

        instructor = get_or_create_user(
            "instructor.demo@gia.local",
            "GIA-INST-001",
            "Instructor",
            "Demo",
            UserRole.FOLLOW_UP_INSTRUCTOR.value,
            instructor_password,
        )
        apprentice_user = get_or_create_user(
            "aprendiz.demo@gia.local",
            "GIA-APR-001",
            "Aprendiz",
            "Demo",
            UserRole.APPRENTICE.value,
            apprentice_password,
        )

        group = TrainingGroup.query.filter_by(group_number="DIA2-3002645").first()
        if group is None:
            group = TrainingGroup(
                created_by=support.id,
                group_number="DIA2-3002645",
                program_name="Tecnólogo en Análisis y Desarrollo de Software",
                lead_instructor=instructor.full_name,
                followup_instructor=instructor.full_name,
                municipality=None,
                program_level=ProgramLevel.TECNOLOGO.value,
                modality=GroupModality.PRESENCIAL.value,
                sofia_group_status=GroupStatus.EN_EJECUCION.value,
                group_start_date=str(date.today() - timedelta(days=120)),
                training_end_date=str(date.today() + timedelta(days=180)),
                ep_start_date=str(date.today() - timedelta(days=15)),
            )
            db.session.add(group)
            db.session.flush()
        else:
            group.followup_instructor = instructor.full_name
            group.lead_instructor = instructor.full_name

        apprentice = Apprentice.query.filter_by(document_number="GIA-APR-001").first()
        if apprentice is None:
            apprentice = Apprentice(
                created_by=support.id,
                student_user_id=apprentice_user.id,
                group_id=group.id,
                group_number=group.group_number,
                document_type=UserDocumentType.NATIONAL_ID.value,
                document_number="GIA-APR-001",
                first_names=apprentice_user.first_names,
                last_names=apprentice_user.last_names,
                program_name=group.program_name,
                program_level=ProgramLevel.TECNOLOGO.value,
                followup_instructor=instructor.full_name,
                followup_instructor_email=instructor.email,
                sofia_status=SofiaStatus.EN_FORMACION.value,
                ep_modality=EpModality.CONTRATO_APRENDIZAJE.value,
                practice_start_date=group.ep_start_date,
                practice_end_date=group.training_end_date,
            )
            db.session.add(apprentice)
            db.session.flush()
        else:
            apprentice.student_user_id = apprentice_user.id
            apprentice.group_id = group.id
            apprentice.group_number = group.group_number
            apprentice.followup_instructor = instructor.full_name
            apprentice.followup_instructor_email = instructor.email

        ensure_template_activities_for_group(group)
        db.session.flush()
        ensure_submissions_for_apprentice(apprentice)
        db.session.commit()

        print("DAY2_SEED=PASS")
        print(f"SUPPORT={support.email}")
        print(f"INSTRUCTOR={instructor.email}")
        print(f"APPRENTICE={apprentice_user.email}")
        print(f"GROUP={group.group_number}")
        print(f"ACTIVITIES={len(group.evidence_activities)}")
        print(f"SUBMISSIONS={len(apprentice.evidence_submissions)}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
