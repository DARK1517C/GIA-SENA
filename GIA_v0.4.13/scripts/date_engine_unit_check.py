from datetime import date
from services import date_rules

r = date_rules.calculate_followup_ranges_from_ep("01/07/2026", "31/12/2026")
assert r["followup_moment1_start"] == "01/07/2026"
assert r["followup_moment1_end"] == "15/07/2026"
assert r["followup_moment3_end"] == "31/12/2026"
assert r["followup_moment4_start"] == ""
assert date_rules.calculate_group_validity_from_training_end("30/06/2027") == "30/12/2027"
print("DATE_ENGINE_UNIT=PASS")
