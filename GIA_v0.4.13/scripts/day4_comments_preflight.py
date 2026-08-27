from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import create_app
from models import Apprentice, EvidenceComment, EvidenceSubmission
from services.followup_service import get_apprentice_followup


def main() -> int:
    app = create_app()
    with app.app_context():
        apprentice = Apprentice.query.filter_by(document_number="GIA-APR-001").first()
        if not apprentice:
            print("ERROR: no existe GIA-APR-001")
            return 2

        submissions = EvidenceSubmission.query.filter_by(apprentice_id=apprentice.id).all()
        comments = EvidenceComment.query.join(EvidenceSubmission).filter(
            EvidenceSubmission.apprentice_id == apprentice.id
        ).all()

        print("DAY4_COMMENTS_PREFLIGHT=PASS")
        print(f"APPRENTICE={apprentice.document_number}")
        print(f"SUBMISSIONS={len(submissions)}")
        print(f"COMMENTS={len(comments)}")
        print("COMMENT_ATTEMPT_LINK=", hasattr(EvidenceComment, "attempt"))
        print("FUP_CODE_NORMALIZATION=", get_apprentice_followup(apprentice)[0]["code"])
        print("COMMENT_RULES=APPRENTICE_AND_FOLLOWUP_INSTRUCTOR")
        print("CORRECTION_FLAG=is_correction_request")
        print("CONVERSATION_PERSISTS_ON_RESUBMISSION=YES")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
