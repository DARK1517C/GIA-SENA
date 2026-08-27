from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv
load_dotenv(ROOT / '.env', override=False)

from app import create_app
from models import Apprentice
from services.certification_service import build_certification_checklist


def main() -> int:
    app = create_app()
    with app.app_context():
        apprentice = Apprentice.query.filter_by(document_number='GIA-APR-001').first()
        if apprentice is None:
            print('ERROR=APRENDIZ_DEMO_NO_EXISTE')
            return 2

        checklist = build_certification_checklist(apprentice)
        print(f'REQUIREMENTS_OK={checklist["requirements_ok"]}')
        print(f'FOLLOWUP_OK={checklist["followup_ok"]}')
        print(f'READY={checklist["ready"]}')

        assert checklist['requirements_ok'] is True
        assert checklist['followup_ok'] is True
        assert checklist['ready'] is True
        assert apprentice.is_certified is False
        print('DAY3_CERTIFICATION_ELIGIBLE_PREFLIGHT=PASS')
        return 0


if __name__ == '__main__':
    raise SystemExit(main())
