# Mega ZIP — cambios consolidados

## 🔴 Bloque estructural
- Cadena Alembic activa consolidada a un único head.
- Migraciones legacy incompatibles retiradas del grafo activo y conservadas en `migrations/legacy_versions/`.
- Se incluye `instance/gia.db` y `database_schema.sql`.

## 🔴 Bloque arquitectura
- Dominio oficial: `EvidenceCategory -> EvidenceTemplate -> EvidenceActivity -> EvidenceSubmission`.
- Nuevo historial real: `EvidenceSubmissionAttempt`.
- `EvidenceComment` queda como trazabilidad estructurada.
- No existe dependencia runtime de `EVIDENCE_TYPES` / `DEFAULT_EVIDENCES`.

## 🔴 Bloque seguridad
- Política única por rol en `services/permissions.py` + alcance en `services/access_scope.py`.
- Administrativo: solo lectura global.
- Certificador: aprobación de evidencias, sin firma.
- Soporte: administración global, sin firma de PDF.
- Instructor normal: alcance por grupos asignados.
- Instructor líder: visión y gestión global.
- Aprendiz: únicamente sus propias evidencias.

## 🔴 Bloque usuarios
- `User` no tiene `username`.
- `login_identifier` usa email o documento.
- Referencias legacy de `username` solo sobreviven en migraciones históricas archivadas.

## 🟠 Bloque evidencias
- `allowed_extensions` y `max_file_size_mb` se aplican durante upload.
- `requires_signature` bloquea aprobación hasta que exista archivo firmado.

## 🟠 Bloque trazabilidad
- `observations TEXT` eliminado del ORM.
- Migración convierte observaciones existentes a `EvidenceComment`.
- Cada reentrega crea un `EvidenceSubmissionAttempt` nuevo.

## 🟠 Bloque limpieza
- Dashboard administrativo ya no es superficie para aprendices; redirige a Evidencias.
- Se eliminó `catalogs/aliases.py`, que no tenía consumidores runtime.
- Se retiraron migraciones legacy del directorio Alembic activo.
- Se sustituyó nomenclatura visual `evidence-category-summary` por `evidence-category-header`.

## Validación ejecutada
- `python -m compileall .` → OK.
- `python scripts/audit_evidence_domain.py` → OK.
- Base SQLite incluida verificada con las tablas canónicas y categorías/plantillas seed.

## Validación pendiente en el entorno del proyecto
`pytest` no pudo ejecutarse en el entorno de construcción porque no están instaladas las dependencias Flask/SQLAlchemy. Ejecutar en el entorno del proyecto:

```powershell
python -m pytest -q tests/test_security_3d_e2e.py
```
