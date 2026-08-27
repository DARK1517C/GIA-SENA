"""Recreate the local SQLite database from the Alembic baseline.

WARNING: this intentionally destroys the local SQLite database and all of its data.
It does not touch PostgreSQL or any DATABASE_URL configured externally.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
INSTANCE = ROOT / "instance"
DB = INSTANCE / "gia.db"


def main() -> int:
    database_url = os.getenv("DATABASE_URL", "")
    if database_url:
        raise SystemExit(
            "DATABASE_URL está configurada. Este script solo resetea la SQLite local. "
            "Desactiva DATABASE_URL antes de continuar."
        )

    INSTANCE.mkdir(parents=True, exist_ok=True)

    for path in (DB, Path(f"{DB}-wal"), Path(f"{DB}-shm")):
        if path.exists():
            path.unlink()

    env = os.environ.copy()
    env["FLASK_APP"] = "app:create_app"

    result = subprocess.run(
        [sys.executable, "-m", "flask", "db", "upgrade"],
        cwd=ROOT,
        env=env,
        check=False,
    )

    return result.returncode


if __name__ == "__main__":
    raise SystemExit(main())
