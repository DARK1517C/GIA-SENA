# v0.4.9.3 — Correcciones funcionales Día 2

- Corregido el renderizado de estados de evidencias: el estado canónico controla texto y color de forma independiente para cada actividad.
- Corregida la semántica de fechas: `Fecha límite de entrega` usa el rango de vencimiento y `Fecha de entrega` usa la fecha real de carga del archivo.
- Perfil de aprendiz: ProgramLevel, EpModality y SofiaStatus se presentan mediante etiquetas normalizadas de catálogo.
- Se conserva la BD y el flujo de migraciones; no requiere una nueva migración para estos cambios.
