# Bloque 3.B — Matriz definitiva de permisos y alcance

## Objetivo
Separar de forma explícita **autorización** (qué acción puede ejecutar un rol)
de **alcance** (sobre qué registros puede ejecutarla).

## Matriz canónica

| Rol | Grupos | Aprendices | Evidencias | Certificación | Usuarios/catálogo |
|---|---|---|---|---|---|
| Aprendiz | sin administración | solo consulta de su ficha | cargar y consultar las propias | no | no |
| Instructor | gestionar grupos asignados | gestionar aprendices de sus grupos | revisar/corregir y consultar sus grupos | no | no |
| Instructor líder | visión global y gestión de grupos | gestión global | revisión/corrección global | no | no |
| Administrativo | visión y gestión institucional de grupos/aprendices | gestión institucional | consulta institucional | no | no |
| Certificador | consulta de grupos necesaria para contexto | consulta necesaria para contexto | consulta global | aprobar y firmar | no |
| Soporte | administración técnica global | administración técnica global | consulta técnica global | no | gestionar usuarios y catálogo de evidencias |

## Reglas de alcance

- Instructor: únicamente registros de grupos cuyo `followup_instructor` coincide con su identidad.
- Instructor líder: alcance global de grupos/aprendices/evidencias.
- Administrativo: alcance global administrativo.
- Certificador: alcance global para evidencias; no obtiene administración de grupos.
- Soporte: alcance global para soporte técnico.
- Aprendiz: únicamente su propio registro y sus entregas.

## Regla de seguridad
Una autorización válida **no sustituye** la comprobación de alcance. Las rutas de mutación de evidencias deben validar ambos componentes.

## Decisiones relevantes de 3.B

1. Soporte no es certificador: no aprueba ni firma evidencias.
2. Certificador no administra usuarios, grupos ni catálogos.
3. Administrativo puede gestionar grupos y aprendices, pero no certifica evidencias.
4. Instructor líder puede operar globalmente sobre el dominio académico, pero no administra cuentas de usuario ni certifica.
5. Instructor queda limitado a sus grupos asignados.
6. Aprendiz nunca recibe permisos administrativos.

Estas decisiones no modifican el esquema de BD ni las migraciones.
