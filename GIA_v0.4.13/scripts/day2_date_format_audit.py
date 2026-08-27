from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
checks = {
    "templates/apprentices/detail.html": [
        "record.practice_start_date|date_dmY",
        "record.practice_end_date|date_dmY",
        "record.followup_moment1_start|date_dmY",
        "record.followup_moment2_start|date_dmY",
        "record.followup_moment3_start|date_dmY",
        "record.followup_moment4_start|date_dmY",
    ],
    "templates/groups/detail.html": [
        "|date_dmY",
    ],
    "templates/auth/profile.html": [
        "apprentice.practice_start_date|date_dmY",
        "apprentice.practice_end_date|date_dmY",
        "apprentice.evaluation_date|date_dmY",
        "apprentice.group_validity|date_dmY",
    ],
    "app.py": ["app.jinja_env.filters[\"date_dmY\"]"],
}

for rel, needles in checks.items():
    text = (ROOT / rel).read_text(encoding="utf-8")
    for needle in needles:
        if needle not in text:
            raise SystemExit(f"FAIL: {rel} falta {needle}")

# Catch the most important raw date fields in these display templates.
for rel in ("templates/apprentices/detail.html", "templates/groups/detail.html", "templates/auth/profile.html"):
    text = (ROOT / rel).read_text(encoding="utf-8")
    raw_patterns = [
        r"record\.(practice_start_date|practice_end_date|evaluation_date|group_validity)\s+or",
        r"apprentice\.(practice_start_date|practice_end_date|evaluation_date|group_validity)\s+or",
    ]
    for pat in raw_patterns:
        if re.search(pat, text):
            raise SystemExit(f"FAIL: {rel} contiene fecha sin filtro: {pat}")

print("DATE_DISPLAY_FORMAT=DD/MM/YYYY")
print("APP_JINJA_DATE_FILTER=PASS")
print("APPRENTICE_DETAIL=PASS")
print("GROUP_DETAIL=PASS")
print("PROFILE_DATE_DISPLAY=PASS")
print("DAY2_DATE_FORMAT_AUDIT=PASS")
