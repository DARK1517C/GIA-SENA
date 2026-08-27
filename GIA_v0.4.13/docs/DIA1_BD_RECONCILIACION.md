# Día 1 — Reconciliación de SQLite/Alembic

## Problema detectado

La SQLite local de GIA ya contiene el esquema funcional, pero no contiene la tabla
`alembic_version`. Por eso `flask db upgrade` interpreta la base como vacía y trata
de ejecutar la migración inicial `b7e2c1a4f901`, provocando `table user already exists`.

## Solución

`python scripts/day1_db_reconcile.py`:

1. Verifica que se está operando sobre SQLite local.
2. Comprueba una firma mínima del esquema actual.
3. No elimina ni recrea la BD.
4. Crea una copia `.pre_stamp_YYYYMMDD_HHMMSS.bak`.
5. Registra el head actual `a5d8e7f4c2b1` mediante `flask db stamp`.
6. Verifica que `alembic_version` quedó en el head.

Después se debe ejecutar:

```powershell
python -m flask --app app:create_app db upgrade
```

La salida esperada es que no haya migraciones pendientes.

## Importante

Este procedimiento es para la **SQLite de desarrollo existente** que ya contiene el
esquema actual. Una BD realmente nueva debe seguir usando `flask db upgrade` desde cero.
