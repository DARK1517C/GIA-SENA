# GIA v0.4.9.6 — Modalidades y auditoría de fechas

## Modalidades

Se alineó el catálogo `EpModality` con las modalidades que ya aparecen en la UI y en la normalización de Excel: se agregó `PASANTIA`, con etiqueta `Pasantías` y alias de importación. También se corrigieron los nombres de campos consumidos por el detalle de grupo para que los conteos de `Pasantías` y `Prácticas en la economía popular y/o campesina` lleguen correctamente al template.

## Fechas — decisión de esta versión

No se cambian todavía las fórmulas institucionales de fechas. El código actual contiene dos implementaciones distintas de `calculate_followup_ranges`: una en `services/utils.py` y otra en `services/followup_service.py`. Sus fórmulas no son equivalentes. Cambiar una por inferencia en este momento podría alterar datos de Excel, seguimiento y vistas existentes.

La recomendación es congelar primero la regla institucional de:
- inicio de etapa productiva
- fin de etapa productiva
- vigencia de ficha
- cuatro momentos de seguimiento
- fechas derivadas de grupo/aprendiz

Después centralizar la fórmula en una sola función y actualizar importación, creación/edición, seguimiento y presentación en una única modificación controlada.
