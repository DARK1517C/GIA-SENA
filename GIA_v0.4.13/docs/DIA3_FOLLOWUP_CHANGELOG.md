# GIA v0.4.12 — Seguimiento funcional (Día 2/3)

## Objetivo
Cerrar la primera capa funcional del seguimiento sin tocar todavía el diseño definitivo de comentarios/notificaciones.

## Cambios
- Estado operativo de M1/M2/M3 calculado desde fechas reales del aprendiz.
- El estado del momento refleja la evidencia FUP correspondiente: pendiente, en curso, vencido, en revisión, requiere corrección o aprobado.
- M4 permanece opcional/no configurado.
- Detalle del aprendiz muestra los tres momentos operativos y el enlace a su evidencia.
- Detalle del grupo muestra el estado y próximo momento de cada aprendiz.
- Dashboard muestra resumen y alertas de seguimiento dentro del alcance visible.
- Corregida la sincronización de vencimientos de actividades FUP: ya no usa `training_end_date` como fin de etapa productiva. Usa las fechas EP reales de los aprendices y evita inventar una ventana grupal cuando existe inconsistencia.
- Se mantienen reentrega y aprobación existentes; este bloque no redefine comentarios/notificaciones.

## Base de datos
No se agrega ninguna migración en esta versión.

## Validaciones
- `scripts/day3_followup_preflight.py` => PASS
- `scripts/day2_evidence_cycle_preflight.py` => PASS
- `scripts/day2_regression_stats.py` => PASS
- `scripts/date_engine_preflight.py` => PASS
- `python -m compileall -q .` => PASS
