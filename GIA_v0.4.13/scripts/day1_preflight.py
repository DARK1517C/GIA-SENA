"""Preflight sin dependencias para el Día 1.

Comprueba que el checkout contiene la estructura mínima, que no hay archivos
transitorios peligrosos y que el grafo de migraciones activas tiene una sola
raíz y un solo head.
"""
from __future__ import annotations

import ast
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

REQUIRED = [
    "app.py", "config.py", "extensions.py", "requirements.txt",
    "models", "routes", "services", "catalogs", "templates",
    "static", "migrations/versions",
]

ERRORS: list[str] = []
WARNINGS: list[str] = []


def exists(rel: str) -> bool:
    return (ROOT / rel).exists()


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def parse_py_files() -> None:
    for path in ROOT.rglob("*.py"):
        if any(part in {".venv", "__pycache__", ".git"} for part in path.parts):
            continue
        try:
            ast.parse(read(path), filename=str(path))
        except SyntaxError as exc:
            ERRORS.append(f"SyntaxError {path.relative_to(ROOT)}:{exc.lineno}: {exc.msg}")


def migration_graph() -> tuple[set[str], set[str], dict[str, list[str]]]:
    versions = ROOT / "migrations" / "versions"
    rev_re = re.compile(r"^\s*revision\s*[:=]\s*[\"']([^\"']+)", re.MULTILINE)
    down_re = re.compile(r"^\s*down_revision\s*[:=]\s*(.+)$", re.MULTILINE)
    revs: dict[str, str | None] = {}
    for path in sorted(versions.glob("*.py")):
        text = read(path)
        rev_m = rev_re.search(text)
        down_m = down_re.search(text)
        if not rev_m:
            ERRORS.append(f"Migración sin revision: {path.name}")
            continue
        rev = rev_m.group(1)
        down_raw = down_m.group(1).strip() if down_m else "None"
        down = None
        if down_raw.startswith(("\"", "'")):
            down = down_raw.strip("\"'")
        revs[rev] = down

    all_revs = set(revs)
    children = {r: [] for r in all_revs}
    for rev, down in revs.items():
        if down is None:
            continue
        if down not in all_revs:
            ERRORS.append(f"Migración {rev} apunta a revisión inexistente: {down}")
        else:
            children[down].append(rev)

    roots = {r for r, down in revs.items() if down is None}
    heads = {r for r, kids in children.items() if not kids}
    return roots, heads, children


def main() -> int:
    for rel in REQUIRED:
        if not exists(rel):
            ERRORS.append(f"Falta ruta obligatoria: {rel}")

    parse_py_files()
    roots, heads, _ = migration_graph()
    if len(roots) != 1:
        ERRORS.append(f"El grafo activo debe tener 1 raíz; encontrado: {sorted(roots)}")
    if len(heads) != 1:
        ERRORS.append(f"El grafo activo debe tener 1 head; encontrado: {sorted(heads)}")

    forbidden = []
    for pattern in ("*.pyc", "*.bak", "*.pyo"):
        forbidden.extend(ROOT.rglob(pattern))
    for path in ROOT.rglob("__pycache__"):
        if path.is_dir():
            forbidden.append(path)
    if forbidden:
        WARNINGS.append("Hay artefactos transitorios en el árbol: " + ", ".join(
            sorted(str(p.relative_to(ROOT)) for p in forbidden)[:12]
        ))

    req = read(ROOT / "requirements.txt") if (ROOT / "requirements.txt").exists() else ""
    for package in ("Flask", "Flask-SQLAlchemy", "Flask-Login", "Flask-WTF", "Flask-Migrate"):
        if package.lower() not in req.lower():
            ERRORS.append(f"requirements.txt no declara {package}")

    print(f"PROJECT={ROOT}")
    print(f"MIGRATION_ROOTS={sorted(roots)}")
    print(f"MIGRATION_HEADS={sorted(heads)}")
    print(f"PYTHON_PARSE={'PASS' if not ERRORS else 'FAIL'}")
    for warning in WARNINGS:
        print(f"WARNING={warning}")
    for error in ERRORS:
        print(f"ERROR={error}")
    if ERRORS:
        return 1
    print("DAY1_PREFLIGHT=PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
