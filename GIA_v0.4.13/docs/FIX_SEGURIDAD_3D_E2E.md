# Fix Seguridad 3D E2E

El fallo observado en `tests/test_security_3d_e2e.py` era un `DetachedInstanceError`: el fixture hace `commit()` y luego cierra el application context antes de que `login()` intente leer `user.login_identifier`.

Cambio aplicado: `login()` vuelve a cargar el `User` por su `id` dentro de un application context activo y obtiene allí `login_identifier`. Esto evita depender de atributos expirados de una instancia ORM detached y no modifica la lógica de autenticación de producción.

Validación en este entorno: no fue posible ejecutar pytest porque el entorno de ejecución no tiene Flask instalado (`ModuleNotFoundError: No module named 'flask'`). En el entorno del proyecto, ejecutar:

```powershell
python -m pytest -q tests/test_security_3d_e2e.py
```
