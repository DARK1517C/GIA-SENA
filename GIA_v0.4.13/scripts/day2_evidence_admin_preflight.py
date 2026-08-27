from pathlib import Path
import ast

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "routes" / "evidence_admin.py"

text = TARGET.read_text(encoding="utf-8")
ast.parse(text, filename=str(TARGET))
required = {
    "get_active_evidence_categories",
    "get_active_evidence_templates",
}
missing = [name for name in sorted(required) if name not in text]
print(f"TARGET={TARGET}")
print("PYTHON_PARSE=PASS")
print(f"CATALOG_HELPERS_REFERENCED={len(required)}")
print("IMPORT_FIX=PASS" if not missing else f"IMPORT_FIX=FAIL missing={missing}")
