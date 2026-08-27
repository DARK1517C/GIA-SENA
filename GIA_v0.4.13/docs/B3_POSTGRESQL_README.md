# Bloque Arquitectura — Fase B.3

## Objetivo
Preparar el proyecto para que el mismo dominio y árbol de migraciones puedan ejecutarse en SQLite y PostgreSQL sin depender de detalles de SQLite.

## Cambios
- Los `Boolean` de los modelos ya no usan `server_default="0"/"1"`; usan `sqlalchemy.true()`/`sqlalchemy.false()`.
- Se mantienen separados los `PRAGMA` de SQLite: solo se registran cuando el backend es SQLite.
- Se mantienen expresiones específicas `sqlite_where`/`postgresql_where` únicamente donde son necesarias para índices parciales.
- Se conserva `DATABASE_URL` como punto de configuración del motor y se normalizan URLs PostgreSQL al driver `postgresql+psycopg`.
- Se elimina el `.bak` histórico de `migrations/versions/`; no forma parte del árbol Alembic activo.
- Se incluye `scripts/audit_database_portability.py` como guardrail de compilación DDL para PostgreSQL y detección de defaults booleanos no portables.

## Validación realizada en este paquete
- Compilación sintáctica de todo el proyecto: PASS.
- No quedan `server_default="0"/"1"` para campos Boolean de los modelos.
- El árbol activo de `migrations/versions/` no contiene copias `.bak`.
- Las migraciones mantienen condiciones explícitas para índices parciales SQLite/PostgreSQL.

## Validación PostgreSQL real
Este entorno de construcción no dispone de un servidor PostgreSQL ni del paquete `psycopg`, por lo que no se declara aquí una prueba real de conexión como PASS. La prueba definitiva debe ejecutarse en un PostgreSQL vacío con:

```powershell
$env:DATABASE_URL="postgresql+psycopg://usuario:clave@localhost:5432/gia"
python -m flask db upgrade
python -m flask db current
```

El resultado esperado es `4a6b9c1d2e30 (head)`. Después deben ejecutarse las comprobaciones de FK, índices y el smoke test funcional.
