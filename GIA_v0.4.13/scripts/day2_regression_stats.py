from pathlib import Path
import ast

ROOT = Path(__file__).resolve().parents[1]
checks = {
    "dashboard_scope_reuses_visible_groups": "visible_group_ids" in (ROOT / "routes" / "dashboard.py").read_text(encoding="utf-8"),
    "no_dashboard_lectiva_kpi": "En etapa lectiva" not in (ROOT / "templates" / "dashboard" / "index.html").read_text(encoding="utf-8"),
    "no_dashboard_productiva_kpi": "En etapa productiva" not in (ROOT / "templates" / "dashboard" / "index.html").read_text(encoding="utf-8"),
    "group_detail_normalized_labels": "apprentice.program_level_label" in (ROOT / "routes" / "groups.py").read_text(encoding="utf-8"),
    "local_group_timestamp": "_local_now_label" in (ROOT / "routes" / "groups.py").read_text(encoding="utf-8"),
    "correction_semantics_preserved": "submission.request_revision(" in (ROOT / "routes" / "evidences.py").read_text(encoding="utf-8"),
}
for py in [ROOT / "routes" / "dashboard.py", ROOT / "routes" / "groups.py", ROOT / "routes" / "evidences.py"]:
    ast.parse(py.read_text(encoding="utf-8"), filename=str(py))
    print(f"PYTHON_PARSE=PASS {py.relative_to(ROOT)}")
for k,v in checks.items():
    print(f"{k}={'PASS' if v else 'FAIL'}")
if not all(checks.values()):
    raise SystemExit(1)
print("DAY2_STATS_REGRESSION=PASS")
