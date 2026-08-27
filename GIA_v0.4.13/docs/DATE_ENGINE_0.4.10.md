# GIA v0.4.10 — Motor de fechas (fundación segura)

## Qué se cambia ahora

1. Las fechas reales importadas del aprendiz (`practice_start_date`, `practice_end_date`) siguen siendo fuente de verdad.
2. El Excel de aprendices **no recibe columnas nuevas de fechas** en esta fase.
3. El grupo conserva las fechas institucionales disponibles (`group_start_date`, `ep_start_date`, `training_end_date`, `group_validity`).
4. `group_validity` solo se deriva si falta y existe `training_end_date`: fin de formación + 6 meses. Nunca se infiere el fin de formación desde el nivel.
5. La lógica de seguimiento se centraliza: M1, M2 y M3 usan el periodo real de EP del aprendiz; M4 permanece pendiente porque todavía no tenemos una regla institucional suficientemente respaldada.
6. Se elimina `Pasantía` del catálogo de las seis modalidades de EP de GIA.

## Qué NO se decide todavía

- Duración universal por nivel (Técnico/Tecnólogo/Operario/Auxiliar). La normativa vigente remite la duración concreta al diseño curricular y al programa.
- Una fórmula definitiva para M4.
- Sobrescribir fechas reales importadas con fechas calculadas.

## Siguiente fase

Cuando el Centro defina o aporte el catálogo de duración por programa, se incorpora una configuración de programa (`nivel`, `duración total`, `lectiva`, `productiva`) y el motor podrá derivar las fechas de un grupo cuando la fuente no las proporcione.
