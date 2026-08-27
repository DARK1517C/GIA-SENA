# GIA v0.4.9.4 — Día 2: estados, perfil y zona horaria

## Correcciones

1. El helper `_profile_catalog_labels()` se movió al ámbito de módulo en `routes/users.py`. Antes quedaba anidado dentro de `users.create()`, por lo que `/users/profile` provocaba `NameError` al intentar renderizar el perfil.

2. El listado de evidencias ahora deriva **texto y color del mismo estado canónico** (`submission.status`). Esto evita que una evidencia `no_entregado` muestre color de `pendiente_revision` por lógica de plantilla heredada.

3. El detalle de evidencia usa `submission.status_color` para que el color del estado sea coherente con su estado real; se eliminó el color naranja fijo.

4. Las fechas `uploaded_at` se muestran mediante `format_datetime_local()` usando `DISPLAY_TIMEZONE`, por defecto `America/Bogota`. Los valores almacenados en UTC/naive se convierten de forma consistente al huso de presentación.

5. No se requiere nueva migración para estos cambios.

## Comprobación

Ejecutar:

```powershell
python scripts\day2_regression_status_profile.py
```

Resultado esperado: `DAY2_REGRESSION=PASS`.
