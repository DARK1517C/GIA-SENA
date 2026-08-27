"""Auditoría estática/BD del ciclo de vida canónico de evidencias.

Uso dentro de la aplicación: python scripts/audit_evidence_lifecycle.py
"""
from __future__ import annotations

from collections import Counter
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import create_app
from extensions import db
from models import EvidenceActivity, EvidenceSubmission


def main() -> int:
    app = create_app()
    with app.app_context():
        problems: list[str] = []

        activities = EvidenceActivity.query.all()
        for activity in activities:
            try:
                activity.validate_domain_consistency(template=activity.template)
            except Exception as exc:
                problems.append(
                    f"actividad #{activity.id}: {exc}"
                )

            if activity.template is not None and activity.category_id != activity.template.category_id:
                problems.append(
                    f"actividad #{activity.id}: category_id no coincide con template_id"
                )

        submissions = EvidenceSubmission.query.all()
        latest = Counter(
            (item.activity_id, item.apprentice_id)
            for item in submissions
            if item.is_latest
        )
        for key, count in latest.items():
            if count > 1:
                problems.append(
                    f"submissions duplicadas como latest para activity={key[0]}, apprentice={key[1]}"
                )

        for submission in submissions:
            activity = submission.activity
            apprentice = submission.apprentice
            group = getattr(apprentice, "group_number", None)
            activity_group = getattr(activity.group, "group_number", None)
            if group is not None and activity_group is not None and group != activity_group:
                problems.append(
                    f"submission #{submission.id}: aprendiz y actividad pertenecen a fichas distintas"
                )

        if problems:
            print("AUDITORIA FALLIDA")
            for problem in problems:
                print(f"- {problem}")
            return 1

        print("AUDITORIA OK: invariantes del dominio de evidencias satisfechas.")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
