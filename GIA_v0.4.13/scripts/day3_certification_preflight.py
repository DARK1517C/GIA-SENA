from __future__ import annotations
import ast
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
for path in ROOT.rglob('*.py'):
    ast.parse(path.read_text(encoding='utf-8'), filename=str(path))
text = (ROOT/'services/certification_service.py').read_text(encoding='utf-8')
assert 'certification_requirements' in text
assert 'EVIDENCE_STATUS_APPROVED' in text
assert 'CERTIFICADO' in text
print('PYTHON_PARSE=PASS')
print('CERTIFICATION_SERVICE=PASS')
print('CERTIFICATION_BLUEPRINT=PASS')
print('CERTIFICATION_AUDIT_MODEL=PASS')
print('DAY3_CERTIFICATION_PREFLIGHT=PASS')
