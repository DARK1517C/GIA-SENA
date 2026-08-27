"""Preflight seguro del motor de fechas GIA 0.4.10. No modifica BD."""
from __future__ import annotations

from pathlib import Path
import importlib.util
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

# Load the pure date rule modules without importing services/__init__.py.
base_spec = importlib.util.spec_from_file_location("gia_utils_base", ROOT / "services" / "utils_base.py")
base_mod = importlib.util.module_from_spec(base_spec)
base_spec.loader.exec_module(base_mod)
sys.modules["gia_utils_base"] = base_mod

rules_spec = importlib.util.spec_from_file_location("gia_date_rules", ROOT / "services" / "date_rules.py")
rules_mod = importlib.util.module_from_spec(rules_spec)
# Provide the relative import dependency manually.
rules_mod.__package__ = "gia_rules_pkg"
sys.modules["gia_rules_pkg.utils_base"] = base_mod
rules_spec.loader.exec_module(rules_mod)

print(f"PROJECT={ROOT}")
print("PASANTIA_ALLOWED=NO")
print("GROUP_VALIDITY_RULE=TRAINING_END_PLUS_6_MONTHS")
print("FOLLOWUP_M1_RULE=EP_START_PLUS_15_DAY_WINDOW")
print("FOLLOWUP_M2_RULE=15_DAY_WINDOW_CENTERED_ON_EP_MIDPOINT")
print("FOLLOWUP_M3_RULE=15_DAY_WINDOW_ENDING_AT_EP_END")
print("FOLLOWUP_M4_RULE=NOT_CONFIGURED")

ranges = rules_mod.calculate_followup_ranges_from_ep("01/07/2026", "31/12/2026")
print("SAMPLE_M1=", ranges["followup_moment1_start"], ranges["followup_moment1_end"])
print("SAMPLE_M2=", ranges["followup_moment2_start"], ranges["followup_moment2_end"])
print("SAMPLE_M3=", ranges["followup_moment3_start"], ranges["followup_moment3_end"])
print("SAMPLE_M4=", ranges["followup_moment4_start"], ranges["followup_moment4_end"])
print("SAMPLE_GROUP_VALIDITY=", rules_mod.calculate_group_validity_from_training_end("30/06/2027"))

from catalogs.apprentice import EpModality
if hasattr(EpModality, "PASANTIA"):
    raise SystemExit("ERROR: PASANTIA sigue presente en EpModality")
print("PYTHON_PARSE=PASS")
print("DATE_ENGINE_PREFLIGHT=PASS")
