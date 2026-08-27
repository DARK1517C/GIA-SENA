from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
model = (ROOT / "models" / "evidence.py").read_text(encoding="utf-8")
route = (ROOT / "routes" / "evidences.py").read_text(encoding="utf-8")
tpl = (ROOT / "templates" / "evidences" / "detail.html").read_text(encoding="utf-8")
checks = {
    "MODEL_CORRECTION_FLAG": "is_correction_request" in model,
    "GENERIC_COMMENT_ROUTE": "def add_comment" in route and '"/<int:submission_id>/comments"' in route,
    "CORRECTION_NOTIF": "notify_apprentice_correction" in route,
    "APPRENTICE_COMMENT_FORM": "apprentice-comment" in tpl,
    "INSTRUCTOR_CORRECTION_CHECKBOX": "request_correction" in tpl,
    "COMMENT_CONVERSATION": "Conversación de la evidencia" in tpl,
}
for key, ok in checks.items(): print(f"{key}={'PASS' if ok else 'FAIL'}")
if not all(checks.values()): raise SystemExit(1)
print("DAY2_COMMENTS_PREFLIGHT=PASS")
