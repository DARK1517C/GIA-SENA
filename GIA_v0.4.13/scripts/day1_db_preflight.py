from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)

from config import Config  # noqa: E402

uri = Config.SQLALCHEMY_DATABASE_URI
print(f"PROJECT={ROOT}")
print(f"DATABASE_DRIVER={'sqlite' if uri.startswith('sqlite:') else uri.split(':', 1)[0]}")

if uri.startswith('sqlite:///'):
    raw = uri[len('sqlite:///'):]
    db_path = Path(raw)
    if not db_path.is_absolute():
        db_path = ROOT / db_path
    print(f"SQLITE_PATH={db_path}")
    print(f"SQLITE_PARENT_EXISTS={db_path.parent.is_dir()}")
    if not db_path.parent.is_dir():
        raise SystemExit('ERROR: no existe el directorio padre de SQLite')

print('DAY1_DB_PREFLIGHT=PASS')
