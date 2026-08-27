"""Smoke test del Día 1.

Debe ejecutarse después de instalar requirements.txt. No modifica datos.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parents[1]
# Permite importar el paquete/aplicación desde scripts\ aunque el cwd sea distinto.
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
load_dotenv(PROJECT_ROOT / ".env", override=False)


def main() -> int:
    missing = []
    for name in ("SECRET_KEY",):
        if not os.getenv(name):
            missing.append(name)
    if missing:
        print("ERROR: faltan variables de entorno: " + ", ".join(missing))
        return 2

    try:
        from app import create_app
    except Exception as exc:
        print(f"ERROR: no se pudo importar la aplicación: {type(exc).__name__}: {exc}")
        return 1

    try:
        app = create_app({"TESTING": True})
        required = {
            "auth.login", "auth.logout", "dashboard.index",
            "apprentices.index", "groups.index", "evidences.index", "users.index",
        }
        registered = {rule.endpoint for rule in app.url_map.iter_rules()}
        missing = sorted(required - registered)
        if missing:
            print("ERROR: endpoints faltantes: " + ", ".join(missing))
            return 1
        print(f"BOOTSTRAP_OK endpoints={len(registered)}")
    except Exception as exc:
        print(f"ERROR: bootstrap falló: {type(exc).__name__}: {exc}")
        return 1
    print("DAY1_SMOKE=PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
