from __future__ import annotations

import getpass
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from dotenv import load_dotenv
load_dotenv(ROOT / ".env", override=False)

from app import create_app
from extensions import db
from models import User


def main() -> int:
    app = create_app()
    with app.app_context():
        user = User.query.filter_by(email="soporte@gia.local").first()
        if user is None:
            print("ERROR: no existe soporte@gia.local")
            return 2

        password = getpass.getpass("Nueva contraseña de SUPPORT (mínimo 8 caracteres): ")
        confirm = getpass.getpass("Confirmar contraseña: ")
        if len(password) < 8:
            print("ERROR: la contraseña debe tener mínimo 8 caracteres.")
            return 2
        if password != confirm:
            print("ERROR: las contraseñas no coinciden.")
            return 2

        user.set_password(password)
        user.status = "ACTIVE"
        db.session.commit()
        print("SUPPORT_PASSWORD_RESET=PASS")
        print(f"USER={user.email}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
