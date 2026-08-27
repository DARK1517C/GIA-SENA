# Corrección: estadísticas de modalidad EP para aprendices certificados

Se corrigieron las estadísticas derivadas para que un aprendiz con estado `CERTIFICADO` no siga contabilizándose en las modalidades de etapa productiva.

## Alcance

- Dashboard: `routes/dashboard.py` filtra certificados antes de agrupar por `ep_modality`.
- Detalle de grupo: `routes/groups.py` excluye certificados de la agregación de las seis modalidades oficiales.
- Las tarjetas `Certificados`, `Con alternativa` y `Sin alternativa` mantienen su lógica separada.
- Se añadió `scripts/day3_regression_certified_ep_modality.py` para verificar la regresión.

## Regla funcional

Un aprendiz certificado puede conservar históricamente su modalidad de etapa productiva en sus datos, pero esa modalidad ya no pertenece a la estadística operativa de EP. El aprendiz se contabiliza únicamente como `Certificado` en las estadísticas de estado.
