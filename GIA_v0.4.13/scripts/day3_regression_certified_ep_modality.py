from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import create_app
from extensions import db
from models import Apprentice, TrainingGroup
from catalogs.apprentice import EpModality, SofiaStatus
from routes.dashboard import _apprentice_statistics
from routes.groups import compute_group_stats


def main() -> int:
    app = create_app()
    with app.app_context():
        apprentice = Apprentice.query.filter_by(document_number='GIA-APR-001').first()
        if apprentice is None:
            print('ERROR: no existe GIA-APR-001')
            return 2

        if not apprentice.is_certified:
            print(f'ERROR: el caso de prueba no está certificado: {apprentice.sofia_status!r}')
            return 3

        original_modality = apprentice.ep_modality
        print('APPRENTICE=', apprentice.document_number)
        print('STATUS=', apprentice.sofia_status)
        print('EP_MODALITY=', original_modality)

        # Simula el comportamiento observado: certificado con modalidad persistida.
        dashboard_stats = _apprentice_statistics({'type': 'global'})
        print('DASHBOARD_BY_EP_MODALITY=', dashboard_stats['by_ep_modality'])

        if any(item.get('count', 0) for item in dashboard_stats['by_ep_modality']):
            print('ERROR: un certificado sigue contando por modalidad EP en dashboard')
            return 4

        group = apprentice.group
        if group is not None:
            group_stats = compute_group_stats(group)
            values = {
                'contrato_aprendizaje': group_stats.contrato_aprendizaje,
                'contrato_vinculo_formativo': group_stats.contrato_vinculo_formativo,
                'vinculo_laboral': group_stats.vinculo_laboral,
                'proyecto_productivo': group_stats.proyecto_productivo,
                'monitoria': group_stats.monitoria,
                'practicas_economia_popular': group_stats.practicas_economia_popular,
            }
            print('GROUP_EP_MODALITIES=', values)
            if any(values.values()):
                print('ERROR: un certificado sigue contando por modalidad EP en grupo')
                return 5

        print('DAY3_CERTIFIED_EP_MODALITY_REGRESSION=PASS')
        return 0


if __name__ == '__main__':
    raise SystemExit(main())
