# Bloque 3.B — Matriz definitiva de permisos y alcance

Base: **GIA_v0.4.8.10**. Esta fase no modifica modelos ni migraciones.

## Regla fundamental

La autorización tiene dos capas y ambas son obligatorias:

1. **Permiso**: el rol puede ejecutar la operación.
2. **Alcance**: el registro concreto pertenece al conjunto que ese usuario puede operar.

Tener permiso no concede por sí solo acceso global.

## Matriz definitiva

| Rol | Ver grupos/aprendices | Gestionar grupos | Gestionar aprendices | Evidencias | Catálogo categorías/plantillas | Actividades | Usuarios | Alcance |
|---|---|---|---|---|---|---|---|---|
| Aprendiz | Solo su información operativa | No | No | Subir propias | No | No | No | Solo su aprendiz y sus evidencias |
| Instructor | Sus grupos/aprendices | Sí | Sí | Revisar, observar, aprobar, subir y firmar | No | Sí | No | Solo grupos asignados |
| Instructor líder | Global | Sí | Sí | Revisar, observar, aprobar, subir y firmar | No | Sí | No | Global |
| Administrativo | Global, solo consulta | No | No | Solo consulta | No | No | No | Global lectura |
| Certificador | Global, solo consulta administrativa | No | No | Revisar, observar y aprobar | No | No | No | Global sobre evidencias |
| Soporte | Global | Sí | Sí | Administración completa, aprobación, subida y firma | Sí | Sí | Sí | Global |

## Decisiones importantes

- **Instructor líder** tiene alcance global. No queda limitado al `followup_instructor` del grupo.
- **Instructor normal** queda estrictamente limitado a grupos donde sea el instructor de seguimiento asignado.
- **Administrativo** es un rol de consulta; no puede modificar grupos, aprendices ni evidencias.
- **Certificador** tiene alcance global sobre evidencias y puede revisarlas/observarlas/aprobarlas, pero no administra grupos, aprendices, catálogo ni usuarios.
- **Soporte** conserva la capacidad técnica/administrativa global.
- **Aprendiz** no entra en los módulos administrativos de grupos/aprendices; opera sus propias evidencias.

## Aplicación técnica

- `services/permissions.py`: matriz de autorización por rol.
- `services/access_scope.py`: alcance por registro.
- `services/auth_helpers.py`: protección de vistas mediante `permission_required()`.
- CRUD de usuarios: `users.manage`.
- CRUD de grupos: `groups.manage` + comprobación de alcance.
- CRUD de aprendices: `apprentices.manage` + comprobación de alcance.
- Categorías/plantillas: `evidences.catalog.manage` (Soporte).
- Actividades: `evidences.activities.manage` + grupo gestionable.

No se añade ninguna migración en esta fase.
