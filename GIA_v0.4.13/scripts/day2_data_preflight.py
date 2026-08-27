from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / ".env", override=False)

from app import create_app
from models import Apprentice, EvidenceActivity, EvidenceSubmission, TrainingGroup, User


def main() -> int:
    app = create_app()
    with app.app_context():
        counts = {
            "users": User.query.count(),
            "groups": TrainingGroup.query.count(),
            "apprentices": Apprentice.query.count(),
            "activities": EvidenceActivity.query.count(),
            "submissions": EvidenceSubmission.query.count(),
        }
        print(f"COUNTS={counts}")
        expected = {
            "instructor.demo@gia.local": User.query.filter_by(email="instructor.demo@gia.local").first(),
            "aprendiz.demo@gia.local": User.query.filter_by(email="aprendiz.demo@gia.local").first(),
        }
        if any(value is None for value in expected.values()):
            print("DAY2_DATA_PREFLIGHT=FAIL missing demo users")
            return 1
        group = TrainingGroup.query.filter_by(group_number="DIA2-3002645").first()
        apprentice = Apprentice.query.filter_by(document_number="GIA-APR-001").first()
        if group is None or apprentice is None:
            print("DAY2_DATA_PREFLIGHT=FAIL missing demo group/apprentice")
            return 1
        if apprentice.group_id != group.id:
            print("DAY2_DATA_PREFLIGHT=FAIL apprentice/group relation")
            return 1
        print("DAY2_DATA_PREFLIGHT=PASS")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
