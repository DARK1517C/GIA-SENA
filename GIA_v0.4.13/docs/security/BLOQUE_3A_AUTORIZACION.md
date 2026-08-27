# Bloque 3.A — Consolidación de autorización

## Objetivo

Centralizar la autorización basada en rol en `services/permissions.py` y aplicar esa política en las vistas administrativas/mutadoras mediante `services.auth_helpers.permission_required`.

El **permiso** responde a "¿puede ejecutar esta operación?". El **alcance** continúa en `services/access_scope.py` y responde a "¿sobre qué registros puede ejecutarla?".

## Roles canónicos

- `APPRENTICE` — Aprendiz
- `FOLLOW_UP_INSTRUCTOR` — Instructor de seguimiento
- `LEAD_FOLLOW_UP_INSTRUCTOR` — Instructor de seguimiento líder
- `CENTER_STAFF` — Administrativo del centro
- `CERTIFIER` — Certificador
- `SUPPORT` — Soporte

## Permisos consolidados

| Permiso | APPRENTICE | FOLLOW_UP_INSTRUCTOR | LEAD_FOLLOW_UP_INSTRUCTOR | CENTER_STAFF | CERTIFIER | SUPPORT |
|---|---:|---:|---:|---:|---:|---:|
| `users.manage` | — | — | — | — | — | ✓ |
| `apprentices.manage` | — | ✓ | ✓ | — | — | ✓ |
| `groups.manage` | — | ✓ | ✓ | — | — | ✓ |
| `evidences.manage` | — | ✓ | ✓ | — | ✓ | ✓ |
| `evidences.approve` | — | ✓ | ✓ | — | ✓ | ✓ |
| `evidences.upload` | ✓ | ✓ | ✓ | — | — | ✓ |
| `evidences.sign` | — | ✓ | ✓ | — | — | — |
| `evidences.catalog.manage` | — | — | — | — | — | ✓ |
| `evidences.activities.manage` | — | ✓ | ✓ | — | — | ✓ |
| `data.global_view` | — | — | ✓ | ✓ | ✓ | ✓ |

## Regla de alcance

La consolidación de autorización no elimina los controles por registro. Un permiso positivo no implica acceso global.

- Instructor de seguimiento: opera sobre grupos/evidencias dentro de su asignación.
- Instructor líder: dispone de visión global de consulta, pero la gestión de grupos sigue estando condicionada por el alcance de gestión existente.
- Administrativo y Certificador: tienen visión global de consulta donde el módulo lo contempla, pero no adquieren por ello permisos de modificación.
- Soporte: mantiene las capacidades administrativas/técnicas globales previstas.
- Aprendiz: queda limitado a sus propios datos/evidencias según los controles de alcance.

## Cambios de esta fase

1. `groups.manage` se incorpora al catálogo canónico de permisos.
2. CRUD/importación administrativa de grupos usa `@permission_required("groups.manage")`.
3. CRUD/importación administrativa de aprendices usa `@permission_required("apprentices.manage")`.
4. CRUD/sincronización de categorías y plantillas usa `@permission_required("evidences.catalog.manage")`.
5. CRUD de actividades usa `@permission_required("evidences.activities.manage")`.
6. Administración de usuarios usa `@permission_required("users.manage")`.
7. `utils/auth.py` deja de contener una política paralela: queda como shim de compatibilidad que reexporta los helpers canónicos.
8. Los helpers de gestión de grupos/aprendices consultan el permiso canónico y dejan el alcance a `access_scope.py`.

## No realizado todavía

Esta fase no redefine todavía el alcance fino de cada operación de todos los módulos ni introduce un sistema ABAC/ACL. Eso corresponde a las fases posteriores del Bloque 3.
