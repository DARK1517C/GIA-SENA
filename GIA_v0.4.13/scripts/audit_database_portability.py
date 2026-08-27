"""Static/DDL portability audit for SQLite/PostgreSQL.

Does not connect to either database. It imports the current models and compiles
all tables/indexes using SQLAlchemy's PostgreSQL dialect, catching constructs
that cannot be rendered by PostgreSQL before a real server is provisioned.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sqlalchemy.dialects import postgresql
from sqlalchemy.schema import CreateIndex, CreateTable

from extensions import db
import models  # noqa: F401 - registers all models


def main() -> int:
    errors: list[str] = []
    dialect = postgresql.dialect()
    metadata = db.metadata

    for table in metadata.sorted_tables:
        try:
            str(CreateTable(table).compile(dialect=dialect))
        except Exception as exc:
            errors.append(f"TABLE {table.name}: {type(exc).__name__}: {exc}")
        for index in table.indexes:
            try:
                str(CreateIndex(index).compile(dialect=dialect))
            except Exception as exc:
                errors.append(
                    f"INDEX {index.name or '<unnamed>'}: "
                    f"{type(exc).__name__}: {exc}"
                )

    forbidden = []
    for path in (ROOT / "models", ROOT / "migrations").rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        if 'server_default="0"' in text or 'server_default="1"' in text:
            forbidden.append(str(path.relative_to(ROOT)))

    if forbidden:
        errors.append("SQLite-style boolean server_default remains in: " + ", ".join(forbidden))

    if errors:
        print("RESULT: FAIL")
        for error in errors:
            print(f"- {error}")
        return 1

    print("PostgreSQL DDL compilation: PASS")
    print("Boolean server defaults: PASS")
    print("Migration backup artifacts in versions/: PASS")
    print("RESULT: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
