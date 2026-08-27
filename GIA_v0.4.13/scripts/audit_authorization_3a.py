"""Auditoría estática sin dependencias externas para Bloque 3.A."""
from pathlib import Path
import ast

ROOT = Path(__file__).resolve().parents[1]
EXPECTED = {
    "routes/users.py": ("users.manage", 4),
    "routes/groups.py": ("groups.manage", 6),
    "routes/apprentices.py": ("apprentices.manage", 6),
    "routes/evidence_admin.py": ("evidences.", 10),
}

errors = []
for rel, (needle, minimum) in EXPECTED.items():
    path = ROOT / rel
    text = path.read_text(encoding="utf-8")
    count = text.count("@permission_required(")
    if count < minimum:
        errors.append(f"{rel}: expected >= {minimum} permission decorators, found {count}")
    if needle not in text:
        errors.append(f"{rel}: missing expected permission family {needle}")

perm_text = (ROOT / "services/permissions.py").read_text(encoding="utf-8")
for required in [
    '"users.manage"',
    '"apprentices.manage"',
    '"groups.manage"',
    '"evidences.manage"',
    '"evidences.approve"',
    '"evidences.upload"',
    '"evidences.sign"',
    '"evidences.catalog.manage"',
    '"evidences.activities.manage"',
    '"data.global_view"',
]:
    if required not in perm_text:
        errors.append(f"services/permissions.py: missing {required}")

# Parse changed Python modules to catch syntax errors without importing Flask.
for rel in [
    "services/permissions.py",
    "services/auth_helpers.py",
    "utils/auth.py",
    "routes/users.py",
    "routes/groups.py",
    "routes/apprentices.py",
    "routes/evidence_admin.py",
]:
    try:
        ast.parse((ROOT / rel).read_text(encoding="utf-8"), filename=rel)
    except SyntaxError as exc:
        errors.append(f"{rel}: SyntaxError {exc}")

if errors:
    print("AUTHORIZATION_3A_FAIL")
    for error in errors:
        print("-", error)
    raise SystemExit(1)

print("AUTHORIZATION_3A_OK")
print("Canonical permissions: OK")
print("Administrative route decorators: OK")
print("Python syntax: OK")
