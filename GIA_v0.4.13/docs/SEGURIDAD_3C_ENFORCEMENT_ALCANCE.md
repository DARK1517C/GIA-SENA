# Bloque 3.C — Enforcement de alcance extremo a extremo

Base: GIA_v0.4.8.10 / Bloque 3.B.

## Objetivo

Evitar que una operación pueda saltarse el alcance de seguridad simplemente
conociendo o manipulando una URL/ID. La autorización por rol y el alcance por
registro deben comprobarse por separado en cada mutación sensible.

## Cambios

- `evidences.observe`: ahora exige alcance de la entrega antes de modificarla.
- `evidences.approve`: ahora exige alcance de la entrega antes de aprobarla.
- `evidences.sign`: ahora exige alcance de la entrega antes de firmarla.
- `evidences.upload` ya contaba con comprobación de alcance y se conserva.
- Descarga, preview y detalle ya estaban protegidos por alcance y se conservan.
- `utils/auth.py` queda como shim hacia `services.permissions`; no conserva roles
  legacy ni una matriz paralela.

## Regla de seguridad

Tener `evidences.approve`, `evidences.manage` o `evidences.sign` no implica
acceso a cualquier registro. Primero se resuelve `can_view_submission()` y
después la capacidad concreta.

## Roles globales

`LEAD_FOLLOW_UP_INSTRUCTOR`, `CENTER_STAFF`, `CERTIFIER` y `SUPPORT` conservan
su alcance global de lectura de grupos/aprendices/evidencias según la política
3.B. El Instructor normal permanece limitado a sus grupos asignados y el
Aprendiz a sus propias entregas.

## Fuera de alcance de 3.C

No se modifican migraciones, esquema ni datos de la BD. Tampoco se cambia la
matriz de roles definida en 3.B; esta fase hace cumplir el alcance ya decidido.
