# Cambios aplicados — Seguridad 3.D E2E

## Problema corregido

Los 7 tests de `tests/test_security_3d_e2e.py` fallaban antes de llegar a la autenticación por `sqlalchemy.orm.exc.DetachedInstanceError`.

La causa es que el fixture `e2e_app` hace `db.session.commit()` y después cierra el `app_context` antes de devolver `users`. SQLAlchemy puede dejar los atributos escalares expirados; al intentar leer `user.login_identifier` sobre la instancia detached, SQLAlchemy intenta refrescarla y no existe una sesión activa.

## Cambio realizado

Archivo:

- `tests/test_security_3d_e2e.py`

La función `login(client, user)` ahora:

1. abre un `app_context` activo;
2. vuelve a cargar `User` con `db.session.get(User, user.id)`;
3. obtiene `login_identifier` de la instancia persistida;
4. cierra el contexto;
5. realiza el POST de login con el identificador ya materializado.

Esto evita tocar la lógica de autenticación de producción y elimina la dependencia de una instancia ORM detached.

## No se cambia

- `models/user.py`
- `extensions.py`
- permisos/roles
- rutas de autorización
- esquema de base de datos
- lógica de autenticación de producción

## Validación

En el entorno de construcción disponible aquí no se pudo ejecutar pytest porque no está instalado Flask (`ModuleNotFoundError: No module named 'flask'`).

En Windows, desde la raíz del proyecto:

```powershell
python -m pytest -q tests/test_security_3d_e2e.py
```
