# GIA v0.4.8 — Mega refactor consolidado

Este paquete consolida los bloques estructural, arquitectura, seguridad, usuarios, evidencias, trazabilidad y limpieza.

## 1. BD + migraciones

- La fuente de verdad es SQLAlchemy + Alembic.
- Se incluye `instance/gia.db` como base SQLite inicial con categorías/plantillas canónicas.
- Se incluye `database_schema.sql` como esquema SQL reproducible.
- Nueva migración: `a5d8e7f4c2b1_evidence_attempt_history_and_traceability.py`.

Para una instalación existente:

```powershell
alembic upgrade head
```

## 2. Dominio oficial de evidencias

```text
EvidenceCategory
    -> EvidenceTemplate
        -> EvidenceActivity
            -> EvidenceSubmission
                -> EvidenceSubmissionAttempt

EvidenceComment = historial estructurado de revisión/retroalimentación
```

No hay consumo activo de `EVIDENCE_TYPES` ni `DEFAULT_EVIDENCES`.

## 3. Seguridad

Política consolidada:

| Rol | Visibilidad | Gestión | Aprobación | Firma | Usuarios |
|---|---|---|---|---|---|
| Aprendiz | propias evidencias | propia entrega | no | no | no |
| Instructor | grupos asignados | sí, dentro de alcance | sí | sí | no |
| Instructor líder | global | sí | sí | sí | no |
| Administrativo | global | no | no | no | no |
| Certificador | global | evidencia/revisión | sí | no | no |
| Soporte | global | técnica/institucional | sí | no | sí |

El alcance por registro sigue separado del permiso por rol.

## 4. User

El modelo actual no contiene `username`. El identificador de acceso es correo o número de documento mediante `login_identifier`.

Las referencias `username` que permanecen están únicamente en migraciones históricas de importación y no forman parte del runtime.

## 5. Políticas de archivo

`allowed_extensions` y `max_file_size_mb` se validan durante la carga.

`requires_signature` se aplica al aprobar: una evidencia que exige firma no puede aprobarse hasta disponer de PDF firmado.

## 6. Trazabilidad

`observations TEXT` fue retirado del modelo activo.

La migración convierte su contenido histórico a `EvidenceComment` y permite autor desconocido (`NULL`) cuando el legado no contiene identidad confiable.

Cada nueva entrega crea un `EvidenceSubmissionAttempt` con número de intento y versión.

## 7. Limpieza

- El aprendiz es redirigido desde el dashboard administrativo hacia Evidencias.
- El resumen visual antiguo `evidence-category-summary` fue sustituido por nomenclatura canónica.
- La auditoría de dominio rechaza dependencias activas con catálogos legacy.
- Se conserva código de migración histórica solo como historial de migración; no se importa en runtime.

## Validaciones incluidas

```powershell
python -m compileall .
python scripts/audit_evidence_domain.py
python -m pytest -q tests/test_security_3d_e2e.py
```

En el entorno de generación del paquete no estaban instaladas las dependencias Flask/SQLAlchemy, por lo que el último comando debe ejecutarse en el entorno Python del proyecto.
