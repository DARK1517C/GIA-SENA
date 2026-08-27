# GIA v0.4.13 — revisión de consistencia y correcciones

Fecha: 2026-08-25

## Correcciones aplicadas

1. `services/followup_service.py`
   - Se normalizan códigos internos de actividades `FUP_01..FUP_04` al código lógico `FUP-01..FUP-04` usado por el motor de seguimiento.
   - Esto corrige la pérdida de asociación entre `EvidenceSubmission` aprobadas y los momentos de seguimiento.

2. `routes/groups.py`
   - Se eliminaron las referencias a atributos inexistentes `stats.en_lectiva` y `stats.en_productiva` durante el recálculo persistido de estadísticas.
   - Los campos heredados `apprentices_training` y `apprentices_practice` se mantienen en `0` para compatibilidad de esquema, pero ya no se calculan como indicadores funcionales.
   - El formulario de creación de grupos deja de solicitar esos dos campos derivados obsoletos.

3. `models/training_group.py`
   - La etiqueta de reporte del campo heredado `internship` deja de mostrarse como `PASANTÍA` y pasa a `VÍNCULO FORMATIVO`.
   - No se agrega ni se reintroduce ninguna modalidad `PASANTIA` en `EpModality`; el catálogo oficial sigue teniendo seis modalidades.

4. `scripts/day3_seed_certification_eligible.py`
   - Se corrigen los códigos reales de las actividades de seguimiento a `FUP_01..FUP_03`.
   - El script deja de afirmar `READY_FOR_CERTIFICATION=YES` de manera fija y valida la elegibilidad con `build_certification_checklist()`.
   - Devuelve código de salida distinto de cero si el caso no resulta realmente elegible.

## Validación estática

- `python -m compileall -q .` → PASS
- Se verificó que no quedan referencias a `stats.en_lectiva` ni `stats.en_productiva`.
- Se verificó que no queda la etiqueta de reporte `PASANTÍA`.
- Se verificó que el helper `_normalize_followup_code` existe en el servicio de seguimiento.

## No modificados deliberadamente

- La lógica de `services.date_rules.py`: las reglas institucionales de fechas siguen siendo la fuente de verdad y no se alteraron sin una nueva definición funcional.
- El catálogo `EpModality`: permanece limitado a las seis modalidades vigentes.
- La lógica de aprobación/rechazo de certificación: se conserva; la corrección se hizo en la asociación de seguimiento y en el seed de prueba.
