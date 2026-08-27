from __future__ import annotations

import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv
load_dotenv(ROOT / '.env', override=False)

from app import create_app
from extensions import db
from models import Apprentice, EvidenceSubmission
from models.evidence import EVIDENCE_STATUS_APPROVED
from services.evidence_service import ensure_submissions_for_apprentice
from services.certification_service import build_certification_checklist

CERT_CATEGORY = 'certification_requirements'
FUP_CODES = {'FUP_01', 'FUP_02', 'FUP_03'}


def main() -> int:
    app = create_app()
    with app.app_context():
        apprentice = Apprentice.query.filter_by(document_number='GIA-APR-001').first()
        if apprentice is None:
            print('ERROR: no existe el aprendiz demo GIA-APR-001.')
            print('Ejecuta primero: python scripts/day2_seed_demo.py')
            return 2

        # Dejamos las fechas de EP en un rango pasado para que el caso sea
        # inequívocamente finalizado a efectos de la prueba de certificación.
        today = datetime.now(timezone.utc).date()
        end = today - timedelta(days=20)
        start = end - timedelta(days=180)
        apprentice.practice_start_date = start.isoformat()
        apprentice.practice_end_date = end.isoformat()
        apprentice.sofia_status = 'EN_FORMACION'
        apprentice.evaluation_date = None

        ensure_submissions_for_apprentice(apprentice)
        db.session.flush()

        submissions = (
            EvidenceSubmission.query
            .join(EvidenceSubmission.activity)
            .join(EvidenceSubmission.activity.property.mapper.class_.category)
            .filter(
                EvidenceSubmission.apprentice_id == apprentice.id,
                EvidenceSubmission.is_latest.is_(True),
            )
            .all()
        )

        now = datetime.now(timezone.utc)
        certified_count = 0
        followup_count = 0

        for submission in submissions:
            category_code = getattr(getattr(submission.activity, 'category', None), 'code', None)
            activity_code = getattr(submission.activity, 'code', None)

            if category_code == CERT_CATEGORY or activity_code in FUP_CODES:
                submission.status = EVIDENCE_STATUS_APPROVED
                submission.reviewed_at = now
                submission.approved_at = now
                # SUPPORT is the deterministic reviewer available in the demo DB.
                reviewer = getattr(apprentice, 'created_by', None)
                if reviewer is not None:
                    reviewer_id = getattr(reviewer, 'id', reviewer if isinstance(reviewer, int) else None)
                    if reviewer_id is not None:
                        submission.reviewed_by = reviewer_id
                        submission.approved_by_id = reviewer_id
                submission.uploaded_at = now - timedelta(minutes=5)
                if not submission.file_name:
                    submission.file_name = f'PRUEBA_CERT_{activity_code or submission.id}.pdf'
                    submission.mime_type = 'application/pdf'
                    submission.file_size_bytes = 1024
                if category_code == CERT_CATEGORY:
                    certified_count += 1
                elif activity_code in FUP_CODES:
                    followup_count += 1

        db.session.commit()

        print('DAY3_CERTIFICATION_ELIGIBLE=PASS')
        print(f'APPRENTICE={apprentice.email or apprentice.document_number}')
        print(f'CERTIFICATION_REQUIREMENTS_APPROVED={certified_count}')
        print(f'FOLLOWUP_MOMENTS_APPROVED={followup_count}')
        checklist = build_certification_checklist(apprentice)
        ready = bool(checklist.get('ready'))

        print('SOFIA_STATUS=EN_FORMACION')
        print(f'REQUIREMENTS_OK={bool(checklist.get('requirements_ok'))}')
        print(f'FOLLOWUP_OK={bool(checklist.get('followup_ok'))}')
        print(f'READY_FOR_CERTIFICATION={"YES" if ready else "NO"}')
        return 0 if ready else 3


if __name__ == '__main__':
    raise SystemExit(main())
