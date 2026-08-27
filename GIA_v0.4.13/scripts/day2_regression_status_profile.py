from pathlib import Path
import ast

ROOT = Path(__file__).resolve().parents[1]

checks = {
    "users_helper_module_scope": "def _profile_catalog_labels" in (ROOT / "routes" / "users.py").read_text(encoding="utf-8") and (ROOT / "routes" / "users.py").read_text(encoding="utf-8").count("def _profile_catalog_labels") == 1,
    "evidence_index_dynamic_status": "status_colors.get(canonical_status" in (ROOT / "templates" / "evidences" / "index.html").read_text(encoding="utf-8"),
    "evidence_detail_dynamic_status": "submission.status_color" in (ROOT / "templates" / "evidences" / "detail.html").read_text(encoding="utf-8"),
    "local_datetime_formatter": "format_datetime_local" in (ROOT / "app.py").read_text(encoding="utf-8"),
    "bogota_default": "America/Bogota" in (ROOT / "config.py").read_text(encoding="utf-8"),
}

for path in ROOT.rglob("*.py"):
    if ".venv" in path.parts:
        continue
    ast.parse(path.read_text(encoding="utf-8"), filename=str(path))

for key, value in checks.items():
    print(f"{key}={'PASS' if value else 'FAIL'}")

if not all(checks.values()):
    raise SystemExit(1)
print("DAY2_REGRESSION=PASS")
