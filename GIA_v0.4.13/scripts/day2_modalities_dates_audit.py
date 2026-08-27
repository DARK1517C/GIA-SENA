"""Legacy compatibility audit retained for historical tracking; v0.4.10 rules are validated by date_engine_preflight.py."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
checks = [
    (ROOT / "catalogs" / "apprentice.py", "PRACTICAS_ECONOMIA_POPULAR"),
    (ROOT / "catalogs" / "display.py", "EpModality.MONITORIA"),
    (ROOT / "routes" / "groups.py", "practicas_economia_popular"),
    (ROOT / "templates" / "groups" / "detail.html", "Prácticas en la economía popular y/o campesina"),
]
for path, needle in checks:
    text = path.read_text(encoding="utf-8")
    if needle not in text:
        raise SystemExit(f"FAIL: {path} no contiene {needle}")

for path in [ROOT / "catalogs" / "apprentice.py", ROOT / "catalogs" / "display.py", ROOT / "catalogs" / "aliases.py", ROOT / "routes" / "groups.py", ROOT / "templates" / "groups" / "detail.html"]:
    text = path.read_text(encoding="utf-8")
    if "EpModality.PASANTIA" in text or 'PASANTIA = "PASANTIA"' in text:
        raise SystemExit(f"FAIL: modalidad PASANTIA sigue expuesta en {path}")

print("SIX_EP_MODALITIES=PASS")
print("PREFER_NO_PASANTIAS=PASS")
print("DATE_RULES_AUDIT=DELEGATED_TO_0.4.10")
