# Bloque Seguridad 3.D — Auditoría de bypasses y enforcement

## Objetivo

Comprobar que la matriz 3.B y el enforcement 3.C no puedan saltarse mediante
peticiones directas, identificadores ajenos o operaciones masivas.

## Hallazgos y correcciones

### 1. Operaciones de evidencia por ID directo

`observe`, `approve` y `sign_submission` cargan la evidencia por identificador,
pero ahora comprueban explícitamente el alcance del registro antes de mutarlo.

**Resultado:** conocer el `submission_id` de otra ficha no concede acceso.

### 2. Importaciones masivas

El importador podía recibir un libro completo y, si solo se comprobaba el
permiso de importación, la operación podía convertirse en un bypass de alcance.

Se añadió `group_scope` a `import_reference_workbook()` y las rutas de grupos y
aprendices lo utilizan para impedir creación/actualización de fichas fuera del
alcance del instructor normal.

### 3. Creación de fichas por instructor

Un instructor normal no tiene un registro existente que sirva como alcance al
crear una ficha. La creación ahora exige que la ficha quede asignada a su propio
seguimiento; si se omite, se asigna automáticamente a su identidad. Líder y
Soporte conservan creación global.

### 4. Operaciones masivas existentes

Las eliminaciones masivas de grupos y aprendices ya filtran por alcance. Las
operaciones globales (`delete-all`) permanecen reservadas a roles globales.

## Separación de responsabilidades

- **Permiso:** qué acción puede ejecutar el rol.
- **Alcance:** sobre qué registro puede ejecutarla.
- **Interfaz:** solo presenta/oculta opciones; nunca se considera una barrera de seguridad.

## Validación

Se incorporó `scripts/audit_security_3d.py`, que no requiere Flask ni SQLAlchemy
y comprueba invariantes estáticas del enforcement. Resultado de esta fase:

`SECURITY_3D_AUDIT=PASS`

La ejecución funcional HTTP con los seis roles requiere un entorno Python con
las dependencias del proyecto y una BD de pruebas limpia; no se debe confundir
la auditoría estática con una prueba E2E.
