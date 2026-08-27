"""Preflight estático del bloque de seguimiento Día 2/3."""
from __future__ import annotations

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

PY_FILES = [
    ROOT / "services" / "followup_service.py",
    ROOT / "services" / "evidence_service.py",
    ROOT / "routes" / "apprentices.py",
    ROOT / "routes" / "groups.py",
    ROOT / "routes" / "dashboard.py",
]

for path in PY_FILES:
    ast.parse(path.read_text(encoding="utf-8"), filename=str(path))

followup = (ROOT / "services" / "followup_service.py").read_text(encoding="utf-8")
evidence = (ROOT / "services" / "evidence_service.py").read_text(encoding="utf-8")
app_detail = (ROOT / "templates" / "apprentices" / "detail.html").read_text(encoding="utf-8")
group_detail = (ROOT / "templates" / "groups" / "detail.html").read_text(encoding="utf-8")
dashboard = (ROOT / "templates" / "dashboard" / "index.html").read_text(encoding="utf-8")

assert "FUP-01" in followup and "FUP-02" in followup and "FUP-03" in followup
assert '"Momento 4"' in followup
assert "get_apprentice_followup" in followup
assert "FOLLOWUP_STATUS_PENDING_REVIEW" in followup
assert "FOLLOWUP_STATUS_REQUIRES_CORRECTION" in followup
assert "calculate_followup_ranges_from_ep" in evidence
sync_block = evidence[evidence.find("def sync_group_followup_dates"):evidence.find("def ensure_submissions_for_apprentice")]
assert "calculate_followup_ranges_from_ep(" in sync_block
assert "ranges = calculate_followup_ranges(" not in sync_block
assert "followup_details" in app_detail
assert "followup_rows" in group_detail or "Seguimiento de la etapa productiva" in group_detail
assert "followup_summary" in dashboard and "followup_alerts" in dashboard

print("FOLLOWUP_SOURCE_OF_TRUTH=PASS")
print("FOLLOWUP_M1_M2_M3=PASS")
print("FOLLOWUP_M4_EXPLICITLY_PENDING=PASS")
print("FOLLOWUP_REVIEW_STATES=PASS")
print("FOLLOWUP_DASHBOARD=PASS")
print("FOLLOWUP_GROUP_DETAIL=PASS")
print("PYTHON_PARSE=PASS")
print("DAY3_FOLLOWUP_PREFLIGHT=PASS")
