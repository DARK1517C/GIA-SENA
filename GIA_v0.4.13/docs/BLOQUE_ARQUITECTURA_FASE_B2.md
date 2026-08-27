# Bloque Arquitectura — Fase B.2
## BD limpia y árbol Alembic único

Esta entrega corrige la bifurcación de migraciones detectada durante la
validación de una base SQLite vacía.

### Árbol canónico activo

`b7e2c1a4f901 -> 9c4d2e7a1b60 -> 3f8a7c2d1e90 -> 4a6b9c1d2e30`

El árbol legacy (`93671fcbae56 -> 7c1f2e9a4b10`) queda conservado únicamente
en `migrations/legacy_versions/` y ya no es descubierto por Alembic.

### Validación prevista

Desde la raíz del proyecto, con `DATABASE_URL` sin definir:

```powershell
python .\scripts\reset_database.py
python -m flask db heads
python -m flask db current
python -c "import sqlite3; c=sqlite3.connect(r'instance\gia.db'); print(c.execute('PRAGMA integrity_check').fetchone()[0]); c.close()"
```

Resultado esperado:

- `flask db heads`: un único `head`, `4a6b9c1d2e30`
- `flask db current`: `4a6b9c1d2e30 (head)`
- `PRAGMA integrity_check`: `ok`

El reset elimina también los archivos SQLite auxiliares `-wal`, `-shm` y
`-journal` antes de ejecutar `flask db upgrade`.

No se incluye una `instance/gia.db` preexistente en esta entrega: la base debe
ser reconstruida desde cero mediante Alembic.
