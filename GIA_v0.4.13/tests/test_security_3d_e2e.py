"""Prueba E2E de seguridad 3.D para los seis roles GIA.

Ejecutar en el entorno real del proyecto:
    python -m pytest -q tests/test_security_3d_e2e.py

La prueba usa Flask test_client contra una BD SQLite temporal y no toca
instance/gia.db ni ninguna BD de desarrollo.
"""
from pathlib import Path

import pytest

from app import create_app
from extensions import db
from models import (
    Apprentice,
    EvidenceActivity,
    EvidenceCategory,
    EvidenceSubmission,
    TrainingGroup,
    User,
)
from models.evidence import EVIDENCE_STATUS_PENDING_REVIEW
from services.permissions import (
    ROLE_APPRENTICE,
    ROLE_FOLLOWUP_INSTRUCTOR,
    ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_CENTER_STAFF,
    ROLE_CERTIFIER,
    ROLE_SUPPORT,
)

PASSWORD = "E2E-Password-2026!"

ROLES = {
    "apprentice": ROLE_APPRENTICE,
    "instructor": ROLE_FOLLOWUP_INSTRUCTOR,
    "leader": ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    "admin": ROLE_CENTER_STAFF,
    "certifier": ROLE_CERTIFIER,
    "support": ROLE_SUPPORT,
}


def _user(role, n, first_names=None):
    user = User(
        document_type="NATIONAL_ID",
        document_number=f"E2E{n:06d}",
        first_names=first_names or role.title(),
        last_names="E2E",
        email=f"e2e.{role}@example.test",
        role=role,
        status="ACTIVE",
    )
    user.set_password(PASSWORD)
    return user


@pytest.fixture()
def e2e_app(tmp_path):
    db_path = Path(tmp_path) / "e2e_security.sqlite3"
    app = create_app({
        "TESTING": True,
        "WTF_CSRF_ENABLED": False,
        "SQLALCHEMY_DATABASE_URI": f"sqlite:///{db_path}",
        "UPLOAD_DIR": str(Path(tmp_path) / "uploads"),
    })

    with app.app_context():
        db.create_all()

        users = {}
        users["apprentice"] = _user(ROLE_APPRENTICE, 1, "Aprendiz")
        users["instructor"] = _user(ROLE_FOLLOWUP_INSTRUCTOR, 2, "Instructor")
        users["leader"] = _user(ROLE_LEAD_FOLLOWUP_INSTRUCTOR, 3, "Lider")
        users["admin"] = _user(ROLE_CENTER_STAFF, 4, "Administrativo")
        users["certifier"] = _user(ROLE_CERTIFIER, 5, "Certificador")
        users["support"] = _user(ROLE_SUPPORT, 6, "Soporte")
        db.session.add_all(users.values())
        db.session.flush()

        own_group = TrainingGroup(
            created_by=users["instructor"].id,
            group_number="E2E-OWN",
            program_name="Programa E2E",
            followup_instructor=users["instructor"].full_name,
        )
        foreign_group = TrainingGroup(
            created_by=users["support"].id,
            group_number="E2E-FOREIGN",
            program_name="Programa E2E Foreign",
            followup_instructor="Otro Instructor",
        )
        db.session.add_all([own_group, foreign_group])
        db.session.flush()

        apprentice_own = Apprentice(
            created_by=users["instructor"].id,
            student_user_id=users["apprentice"].id,
            group_id=own_group.id,
            group_number=own_group.group_number,
            document_type="NATIONAL_ID",
            document_number="E2E-APP-001",
            first_names="Aprendiz",
            last_names="Propio",
        )
        apprentice_foreign = Apprentice(
            created_by=users["support"].id,
            student_user_id=None,
            group_id=foreign_group.id,
            group_number=foreign_group.group_number,
            document_type="NATIONAL_ID",
            document_number="E2E-APP-002",
            first_names="Aprendiz",
            last_names="Ajeno",
        )
        db.session.add_all([apprentice_own, apprentice_foreign])
        db.session.flush()

        category = EvidenceCategory(
            code="E2E-CAT",
            name="E2E Category",
            sort_order=0,
            is_active=True,
        )
        db.session.add(category)
        db.session.flush()

        activity_own = EvidenceActivity(
            group_id=own_group.id,
            category_id=category.id,
            title="E2E Activity Own",
            origin="template",
        )
        activity_foreign = EvidenceActivity(
            group_id=foreign_group.id,
            category_id=category.id,
            title="E2E Activity Foreign",
            origin="template",
        )
        db.session.add_all([activity_own, activity_foreign])
        db.session.flush()

        submission_own = EvidenceSubmission(
            activity_id=activity_own.id,
            apprentice_id=apprentice_own.id,
            status=EVIDENCE_STATUS_PENDING_REVIEW,
            version_number=1,
            attempt_number=1,
            is_latest=True,
        )
        submission_foreign = EvidenceSubmission(
            activity_id=activity_foreign.id,
            apprentice_id=apprentice_foreign.id,
            status=EVIDENCE_STATUS_PENDING_REVIEW,
            version_number=1,
            attempt_number=1,
            is_latest=True,
        )
        db.session.add_all([submission_own, submission_foreign])
        db.session.commit()

        ids = {
            "own_group": own_group.id,
            "foreign_group": foreign_group.id,
            "own_apprentice": apprentice_own.id,
            "foreign_apprentice": apprentice_foreign.id,
            "own_submission": submission_own.id,
            "foreign_submission": submission_foreign.id,
        }

    yield app, users, ids

    with app.app_context():
        db.drop_all()


def login(client, user):
    # The fixture closes its application context before yielding, so ORM
    # instances returned in ``users`` are detached and their expired
    # scalar attributes cannot be lazy-loaded. Re-load the user inside a
    # live application context instead of relying on a detached instance.
    with client.application.app_context():
        persisted_user = db.session.get(User, user.id)
        assert persisted_user is not None
        identifier = persisted_user.login_identifier

    response = client.post(
        "/auth/login",
        data={"identifier": identifier, "password": PASSWORD},
        follow_redirects=False,
    )
    assert response.status_code in {302, 303}, response.data[:500]


def test_all_six_roles_can_authenticate(e2e_app):
    app, users, _ = e2e_app
    for name in ROLES:
        client = app.test_client()
        login(client, users[name])


def test_role_gates_are_enforced_over_http(e2e_app):
    app, users, _ = e2e_app

    expected = {
        "apprentice": {"groups": False, "users": False, "catalog": False},
        "instructor": {"groups": True, "users": False, "catalog": False},
        "leader": {"groups": True, "users": False, "catalog": False},
        "admin": {"groups": False, "users": False, "catalog": False},
        "certifier": {"groups": False, "users": False, "catalog": False},
        "support": {"groups": True, "users": True, "catalog": True},
    }

    for name, matrix in expected.items():
        client = app.test_client()
        login(client, users[name])

        group = client.get("/groups/create")
        users_create = client.get("/users/create")
        catalog_create = client.get("/evidence-admin/categories/create")

        assert (group.status_code == 200) is matrix["groups"], (name, group.status_code)
        assert (users_create.status_code == 200) is matrix["users"], (name, users_create.status_code)
        assert (catalog_create.status_code == 200) is matrix["catalog"], (name, catalog_create.status_code)


def test_scope_blocks_instructor_from_foreign_records(e2e_app):
    app, users, ids = e2e_app
    client = app.test_client()
    login(client, users["instructor"])

    assert client.get(f"/groups/{ids['foreign_group']}").status_code in {302, 403}
    assert client.get(f"/apprentices/{ids['foreign_apprentice']}").status_code in {302, 403}
    assert client.get(f"/evidencias/{ids['foreign_submission']}").status_code == 403

    assert client.get(f"/groups/{ids['own_group']}").status_code == 200
    assert client.get(f"/apprentices/{ids['own_apprentice']}").status_code == 200
    assert client.get(f"/evidencias/{ids['own_submission']}").status_code == 200


def test_apprentice_isolation(e2e_app):
    app, users, ids = e2e_app
    client = app.test_client()
    login(client, users["apprentice"])

    assert client.get(f"/evidencias/{ids['own_submission']}").status_code == 200
    assert client.get(f"/evidencias/{ids['foreign_submission']}").status_code == 403
    assert client.get(f"/apprentices/{ids['own_apprentice']}").status_code == 200
    assert client.get(f"/apprentices/{ids['foreign_apprentice']}").status_code in {302, 403}


def test_global_roles_can_view_foreign_evidence(e2e_app):
    app, users, ids = e2e_app
    for name in ("leader", "admin", "certifier", "support"):
        client = app.test_client()
        login(client, users[name])
        assert client.get(f"/evidencias/{ids['foreign_submission']}").status_code == 200


def test_evidence_mutation_permission_and_scope(e2e_app):
    app, users, ids = e2e_app

    # Instructor: own evidence allowed, foreign evidence forbidden.
    client = app.test_client()
    login(client, users["instructor"])
    own = client.post(
        f"/evidencias/{ids['own_submission']}/observe",
        data={"observation": "E2E observation"},
        follow_redirects=False,
    )
    foreign = client.post(
        f"/evidencias/{ids['foreign_submission']}/observe",
        data={"observation": "Should be blocked"},
        follow_redirects=False,
    )
    assert own.status_code != 403
    assert foreign.status_code == 403

    # Center staff: global visibility but no evidence-management permission.
    client = app.test_client()
    login(client, users["admin"])
    assert client.post(
        f"/evidencias/{ids['foreign_submission']}/observe",
        data={"observation": "Should be blocked"},
        follow_redirects=False,
    ).status_code == 403

    # Certifier: global evidence approval, but no signature permission.
    client = app.test_client()
    login(client, users["certifier"])
    assert client.post(
        f"/evidencias/{ids['foreign_submission']}/approve",
        follow_redirects=False,
    ).status_code != 403
    assert client.post(
        f"/evidencias/{ids['foreign_submission']}/sign",
        data={"x": "0", "y": "0", "page": "1"},
        follow_redirects=False,
    ).status_code == 403

    # Support: global management, but signature remains restricted by policy.
    client = app.test_client()
    login(client, users["support"])
    assert client.post(
        f"/evidencias/{ids['foreign_submission']}/observe",
        data={"observation": "Support E2E"},
        follow_redirects=False,
    ).status_code != 403
    assert client.post(
        f"/evidencias/{ids['foreign_submission']}/sign",
        data={"x": "0", "y": "0", "page": "1"},
        follow_redirects=False,
    ).status_code == 403


def test_support_only_user_administration(e2e_app):
    app, users, _ = e2e_app
    for name in ("apprentice", "instructor", "leader", "admin", "certifier"):
        client = app.test_client()
        login(client, users[name])
        assert client.get("/users/create").status_code == 403

    client = app.test_client()
    login(client, users["support"])
    assert client.get("/users/create").status_code == 200
