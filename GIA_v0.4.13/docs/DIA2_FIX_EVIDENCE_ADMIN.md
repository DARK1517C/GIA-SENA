# Día 2 — Corrección de Administración de Evidencias

## Hallazgo
`routes/evidence_admin.py` utilizaba `get_active_evidence_categories()` y `get_active_evidence_templates()` sin importarlos desde `services.evidence_service`, provocando `NameError` al abrir `/evidencias/admin/`.

## Corrección
Se añadieron ambos imports y una prueba de regresión de los helpers.

## Verificación
Ejecutar:

```powershell
python scripts\day2_evidence_admin_preflight.py
```

Resultado esperado: `IMPORT_FIX=PASS`.
