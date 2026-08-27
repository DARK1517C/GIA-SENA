# GIA v0.4.12 — Seguimiento Día 2/3

## Objetivo

Conectar el motor de fechas ya validado con el ciclo operativo de seguimiento usando las actividades institucionales FUP-01, FUP-02 y FUP-03.

## Alcance de esta versión

- M1, M2 y M3 se calculan con el motor central `services/date_rules.py`.
- El estado operativo de cada momento se deriva de su evidencia más reciente.
- `Pendiente de revisión`, `Requiere corrección` y `Aprobado` se reflejan directamente en el seguimiento.
- Grupo y aprendiz muestran el mismo estado de seguimiento sin crear una segunda fuente de verdad.
- Dashboard muestra resumen y alertas para el alcance visible.
- Momento 4 permanece visible pero no operativo hasta definir su regla institucional.

## Reentrega y aprobación

Este bloque reutiliza el ciclo de evidencias ya existente:

`Pendiente de revisión -> Requiere corrección -> Pendiente de revisión -> Aprobado`

La versión no amplía ni redefine el diseño provisional de comentarios/notificaciones.

## Regla de fechas

Se mantienen las reglas ya aceptadas para GIA:

- Momento 1: primeros 15 días desde el inicio de EP.
- Momento 2: ventana de 15 días centrada en la mitad de la EP.
- Momento 3: últimos 15 días de EP.
- Momento 4: pendiente.

No se usa `training_end_date` como sustituto de `practice_end_date` del aprendiz para calcular el seguimiento.
