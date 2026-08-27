"""Smoke test del bootstrap de GIA.

Ejecutar con el entorno virtual instalado:
    python scripts/smoke_bootstrap.py
"""

from app import create_app


app = create_app({"TESTING": True})

required = {
    "auth.login",
    "auth.logout",
    "dashboard.index",
    "apprentices.index",
    "groups.index",
    "evidences.index",
    "users.index",
}

registered = {rule.endpoint for rule in app.url_map.iter_rules()}
missing = sorted(required - registered)

if missing:
    raise SystemExit(f"Bootstrap incompleto. Blueprints/endpoints faltantes: {missing}")

print(f"BOOTSTRAP_OK: {len(registered)} endpoints registrados")
