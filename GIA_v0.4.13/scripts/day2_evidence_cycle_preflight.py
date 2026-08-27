from pathlib import Path
import ast

ROOT = Path(__file__).resolve().parents[1]
files = [
    ROOT / "models" / "evidence.py",
    ROOT / "routes" / "evidences.py",
    ROOT / "templates" / "evidences" / "detail.html",
]

for f in files:
    if f.suffix == ".py":
        ast.parse(f.read_text(encoding="utf-8"), filename=str(f))

model = (ROOT / "models" / "evidence.py").read_text(encoding="utf-8")
routes = (ROOT / "routes" / "evidences.py").read_text(encoding="utf-8")
tpl = (ROOT / "templates" / "evidences" / "detail.html").read_text(encoding="utf-8")

assert "def request_revision(" in model
assert "reviewed_by_id: int | None = None" in model
assert "submission.request_revision(reviewed_by_id=current_user.id)" in routes
assert "submission.status == EVIDENCE_STATUS_PENDING_REVIEW" in tpl
assert "submission.attempt_history" in tpl
assert "Pasantia" not in model
print("REDELIVERY_MODEL=PASS")
print("REVIEWER_AUDIT=PASS")
print("APPROVAL_STATE_GUARD=PASS")
print("ATTEMPT_HISTORY_UI=PASS")
print("PYTHON_PARSE=PASS")
print("DAY2_EVIDENCE_CYCLE_PREFLIGHT=PASS")
