# Día 2 — v0.4.9.5: estadísticas y estados operativos

## Cambios
- El dashboard deja de mostrar “En etapa lectiva” y “En etapa productiva” como KPIs.
- “Habilitados” y “Con alternativa” se derivan de la modalidad de etapa productiva mientras el aprendiz no esté certificado.
- Un aprendiz certificado se cuenta exclusivamente en “Certificados”.
- “Sin alternativa” excluye certificados y aprendices con modalidad.
- El dashboard del instructor reutiliza el mismo alcance por grupos que el módulo Grupos, evitando dashboards en cero cuando existen grupos asignados.
- El detalle de grupo muestra los mismos cuatro indicadores.
- Nivel de formación, modalidad y estado SOFIA del aprendiz se presentan con etiquetas de catálogo.
- La fecha “Última actualización” del detalle de grupo usa `DISPLAY_TIMEZONE` (por defecto `America/Bogota`).
- La observación del instructor mantiene su semántica actual: registrar comentario + pasar la evidencia a `Requiere corrección`.

## No requiere migración
Los cambios son de lógica de consulta y presentación y no alteran el esquema de base de datos.
