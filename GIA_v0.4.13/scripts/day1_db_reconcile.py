"""Safely reconcile an existing local SQLite database with Alembic.

Use this only when the database already contains the current GIA schema but
has no alembic_version table (typical of the shipped development DB).
The script never deletes or recreates the database. It makes a timestamped
backup, verifies a minimum head-schema signature, then stamps the configured
Alembic head. Afterward `flask db upgrade` should be a no-op.
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import subprocess
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
HEAD = "a5d8e7f4c2b1"

REQUIRED_TABLES = {
    "user",
    "training_group",
    "apprentice",
    "evidence_category",
    "evidence_template",
    "evidence_activity",
    "evidence_submission",
    "evidence_comment",
    "evidence_submission_attempt",
}

REQUIRED_COLUMNS = {
    "evidence_submission_attempt": {"submission_id", "attempt_number", "version_number", "status"},
    "evidence_activity": {"category_id", "template_id"},
    "evidence_submission": {"attempt_number", "version_number", "file_path", "status"},
    "evidence_comment": {"submission_id", "comment"},
    "user": {"id", "email", "password_hash", "role"},
}


def fail(message: str) -> None:
    print(f"ERROR: {message}")
    raise SystemExit(1)


def sqlite_path() -> Path:
    url = os.getenv("DATABASE_URL", "").strip()
    if url and not url.startswith("sqlite"):
        fail("DATABASE_URL no es SQLite; este reconciliador solo opera sobre SQLite local.")
    instance = PROJECT_ROOT / "instance"
    instance.mkdir(parents=True, exist_ok=True)
    path = instance / "gia.db"
    if url.startswith("sqlite:///"):
        raw = url[len("sqlite:///"):]
        if os.path.isabs(raw):
            path = Path(raw)
        else:
            path = (PROJECT_ROOT / raw).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def inspect_db(path: Path) -> tuple[set[str], dict[str, set[str]], bool]:
    if not path.exists():
        fail(f"No existe la base SQLite: {path}")
    con = sqlite3.connect(path)
    try:
        tables = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        cols = {}
        for table in REQUIRED_TABLES:
            if table in tables:
                cols[table] = {r[1] for r in con.execute(f'PRAGMA table_info("{table}")')}
        has_version = "alembic_version" in tables
        return tables, cols, has_version
    finally:
        con.close()


def main() -> int:
    print(f"PROJECT={PROJECT_ROOT}")
    print(f"HEAD={HEAD}")
    path = sqlite_path()
    print(f"SQLITE_PATH={path}")

    tables, cols, has_version = inspect_db(path)
    missing_tables = REQUIRED_TABLES - tables
    missing_columns = {
        table: sorted(expected - cols.get(table, set()))
        for table, expected in REQUIRED_COLUMNS.items()
        if expected - cols.get(table, set())
    }

    print(f"REQUIRED_TABLES_MISSING={sorted(missing_tables)}")
    print(f"REQUIRED_COLUMNS_MISSING={missing_columns}")
    print(f"ALEMBIC_VERSION_TABLE={has_version}")

    if missing_tables or missing_columns:
        fail("La BD no coincide con la firma mínima esperada del head; no se hará stamp automático.")

    if has_version:
        con = sqlite3.connect(path)
        try:
            rows = con.execute("SELECT version_num FROM alembic_version").fetchall()
        finally:
            con.close()
        print(f"ALEMBIC_VERSION={rows}")
        if rows == [(HEAD,)]:
            print("DAY1_DB_RECONCILE=ALREADY_AT_HEAD")
            return 0
        fail("La tabla alembic_version existe pero no está en el head esperado; detener y revisar.")

    stamp_backup = path.with_name(path.stem + f".pre_stamp_{datetime.now():%Y%m%d_%H%M%S}.bak")
    shutil.copy2(path, stamp_backup)
    print(f"BACKUP={stamp_backup}")

    env = os.environ.copy()
    subprocess.run(
        [sys.executable, "-m", "flask", "--app", "app:create_app", "db", "stamp", HEAD],
        cwd=PROJECT_ROOT,
        env=env,
        check=True,
    )

    con = sqlite3.connect(path)
    try:
        rows = con.execute("SELECT version_num FROM alembic_version").fetchall()
    finally:
        con.close()
    if rows != [(HEAD,)]:
        fail(f"Stamp no dejó el head esperado. Encontrado: {rows}")

    print("DAY1_DB_RECONCILE=PASS")
    print("Siguiente: python -m flask --app app:create_app db upgrade")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
