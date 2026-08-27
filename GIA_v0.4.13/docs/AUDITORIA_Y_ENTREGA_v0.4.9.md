# GIA v0.4.9 — Auditoría y entrega de estabilización

## Alcance de esta entrega

Esta versión parte de GIA v0.4.8 y se concentra en **estabilización, seguridad transversal, integración y reproducibilidad**, sin rehacer la arquitectura existente.

## Correcciones realizadas

### Seguridad
- Activación real de CSRF mediante Flask-WTF/`CSRFProtect`.
- Todas las formas HTML POST detectadas incluyen token CSRF.
- Cierre de sesión cambiado de GET a POST protegido por CSRF.
- Cambio de contraseña corregido: exige contraseña actual, nueva contraseña y confirmación.
- Mínimo de 8 caracteres para contraseñas nuevas.
- El rol de un usuario ya no puede modificarse desde su propio perfil.
- `SECRET_KEY` es obligatoria fuera de modo debug.
- Cookies de sesión endurecidas (`HttpOnly`, `SameSite=Lax` y `Secure` configurable).
- Se eliminó el registro de la URI completa de base de datos para no exponer credenciales en logs.
- Validación básica por firmas binarias para uploads PDF, imágenes y contenedores Office.

### Integración
- Se reparó la URL de descarga de evidencias desde el detalle del aprendiz.
- Se añadieron las vistas de importación de aprendices y grupos que los endpoints ya esperaban.
- Se implementó el endpoint `groups.recalculate_stats` que la interfaz ya tenía referenciado.
- Se permitió explícitamente `users.profile` para aprendices en la restricción central de endpoints.
- Se sincronizan datos de contacto del aprendiz con su perfil cuando existe un usuario aprendiz vinculado.

### Migraciones
- Se dejó una única cadena activa en `migrations/versions/`:

  `b7e2c1a4f901 -> 9c4d2e7a1b60 -> 3f8a7c2d1e90 -> 4a6b9c1d2e30 -> a5d8e7f4c2b1`

- Las migraciones históricas incompatibles permanecen exclusivamente en `migrations/legacy_versions/`.
- README actualizado para reconocer `a5d8e7f4c2b1` como head activo.

### Higiene del proyecto
- `.gitignore` añadido.
- `.env.example` añadido.
- Se eliminan bytecode, caches, `.bak` y artefactos WAL/SHM del entregable.
- Se añade una batería de contratos estáticos ejecutable sin Flask (`tests/test_static_contracts.py`).

## Verificaciones ejecutadas aquí

- `python -m compileall -q .` → **OK**.
- `pytest -q tests/test_static_contracts.py` → **5 passed**.
- Auditoría de autorización 3A → **PASS**.
- Auditoría de alcance 3D → **PASS**.
- Auditoría de dominio de evidencias → **PASS**.
- Auditoría de integración del dominio de evidencias → **PASS**.
- Comprobación del grafo activo de migraciones → **1 raíz / 1 head**.
- Comprobación estática de endpoints usados por `url_for()` → **sin endpoints inexistentes**.
- Comprobación estática de formularios POST → **sin formularios HTML POST sin CSRF**.

## Lo que NO pude certificar en este entorno

El entorno de auditoría no tiene instaladas las dependencias Flask/Flask-WTF/Flask-SQLAlchemy/Flask-Migrate y no tiene acceso al índice de paquetes para instalarlas. Por ello **no se debe interpretar esta entrega como una certificación de pruebas HTTP/E2E ejecutadas**.

Los tests E2E, migraciones contra PostgreSQL y pruebas CRUD completas deben ejecutarse en el entorno real de desarrollo/despliegue del proyecto.

## Pendientes prioritarios para los siguientes 5 días

1. Ejecutar instalación real de dependencias y levantar la aplicación desde cero.
2. Ejecutar `alembic upgrade head` sobre una base limpia y verificar PostgreSQL.
3. Ejecutar la suite E2E de seguridad y crear pruebas CRUD para usuarios, grupos, aprendices y evidencias.
4. Probar importación/exportación Excel con libros institucionales reales de prueba.
5. Probar carga, revisión, corrección, aprobación, firma y descarga de evidencias con archivos válidos y corruptos.
6. Validar el flujo completo por cada rol.
7. Completar certificación, notificaciones y reportes según los requisitos institucionales que sigan pendientes.
8. Hacer un ensayo de despliegue y respaldo/restauración.

## Recomendación

No continuar añadiendo arquitectura general salvo que aparezca un bloqueo. El siguiente esfuerzo debe ser **integración → pruebas → corrección → seguridad → despliegue**.
