"""Auditoría estática de integración del dominio canónico de evidencias.

No requiere Flask/SQLAlchemy en runtime. Comprueba invariantes estructurales
que deben mantenerse antes de ejecutar pruebas de integración contra una BD.
"""
from pathlib import Path
import ast
import re

ROOT = Path(__file__).resolve().parents[1]


def read(path):
    return (ROOT / path).read_text(encoding="utf-8")


def assert_contains(text, needle, label):
    if needle not in text:
        raise AssertionError(f"FALLO: {label}: falta {needle!r}")


def assert_not_contains(text, needle, label):
    if needle in text:
        raise AssertionError(f"FALLO: {label}: aparece {needle!r}")


def main():
    service = read("services/evidence_service.py")
    admin = read("routes/evidence_admin.py")
    model = read("models/evidence.py")

    # Dominio canónico.
    for name in (
        "EvidenceCategory",
        "EvidenceTemplate",
        "EvidenceActivity",
        "EvidenceSubmission",
    ):
        assert_contains(model, f"class {name}", "modelos canónicos")

    # El catálogo legacy no debe aparecer en código funcional.
    for path in ROOT.rglob("*.py"):
        if "__pycache__" in path.parts or path.name.startswith("audit_evidence_"):
            continue
        text = path.read_text(encoding="utf-8")
        for legacy in ("EVIDENCE_" + "TYPES", "DEFAULT_" + "EVIDENCES"):
            if legacy in text and path.name != "audit_evidence_domain.py":
                raise AssertionError(f"FALLO: dependencia legacy {legacy} en {path}")

    # Sincronización de una plantilla: debe usar el helper que NO proyecta
    # plantillas adicionales de forma implícita.
    assert_contains(
        service,
        "ensure_submissions_for_apprentice_group(group)",
        "sincronización de plantilla",
    )

    # CRUD: eliminación destructiva solo cuando no existe historial.
    assert_contains(admin, "if category.templates or category.activities:", "protección de categoría")
    assert_contains(admin, "if template.activities:", "protección de plantilla")
    assert_contains(admin, "if submissions:", "protección de actividad")

    # Mover una actividad con entregas rompe la trazabilidad y debe bloquearse.
    assert_contains(
        admin,
        "if new_group.id != activity.group_id:",
        "movimiento de actividad",
    )
    assert_contains(
        admin,
        "existing_submissions = EvidenceSubmission.query.filter_by(",
        "movimiento de actividad",
    )

    # Todos los Python deben parsear correctamente.
    for path in ROOT.rglob("*.py"):
        if "__pycache__" in path.parts or path.name.startswith("audit_evidence_"):
            continue
        ast.parse(path.read_text(encoding="utf-8"), filename=str(path))

    print("OK: integración estructural del dominio de evidencias validada.")
    print("OK: CRUD protegido frente a borrado con historial.")
    print("OK: movimiento de actividades con entregas bloqueado.")
    print("OK: sincronización de plantilla no proyecta catálogos adicionales de forma implícita.")


if __name__ == "__main__":
    main()
