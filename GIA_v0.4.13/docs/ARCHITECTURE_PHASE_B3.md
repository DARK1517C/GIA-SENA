# Bloque de Arquitectura — Fase B.3
## Portabilidad SQLite → PostgreSQL

### Objetivo

Mantener SQLite como motor local de desarrollo/pruebas y preparar el mismo
modelo SQLAlchemy + árbol Alembic para PostgreSQL sin introducir una segunda
arquitectura de persistencia.

### Cambios realizados

1. Los `Boolean` de los modelos ya no usan `server_default="0"` / `server_default="1"`.
   Se utilizan `false()` / `true()` de SQLAlchemy, que generan expresiones
   apropiadas para cada dialecto.
2. Se conserva la configuración por `DATABASE_URL` y la normalización de
   `postgres://` / `postgresql://` hacia `postgresql+psycopg://`.
3. La configuración SQLite específica continúa aislada en `extensions.py` y
   solo se ejecuta para conexiones SQLite.
4. Las migraciones de índices parciales mantienen predicados específicos para
   SQLite y PostgreSQL cuando el dialecto requiere una expresión distinta.
5. Se añade `scripts/audit_database_portability.py` como guardrail estático.

### No se hace todavía

- No se migra `instance/gia.db` a PostgreSQL.
- No se cambia el motor por defecto de desarrollo.
- No se elimina ninguna utilidad legítima de mantenimiento específica de SQLite.
- No se declara PostgreSQL validado funcionalmente hasta ejecutar Alembic y las
  pruebas de integración contra una instancia PostgreSQL real.

### Criterio de cierre B.3

SQLite:

```text
BD vacía → flask db upgrade → HEAD único → integridad OK
```

PostgreSQL:

```text
BD vacía → flask db upgrade → HEAD único → FK/índices OK
```

La segunda prueba requiere un servidor PostgreSQL real; la ausencia de servidor
no se debe ocultar mediante una simulación.
