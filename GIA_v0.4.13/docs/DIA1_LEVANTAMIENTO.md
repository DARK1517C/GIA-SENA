# GIA v0.4.9 — Día 1: levantamiento y validación

## Objetivo

Conseguir una instalación reproducible de GIA y verificar que la aplicación arranca con la cadena de migraciones actual.

## Windows

Desde la raíz del proyecto:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\setup_windows_day1.ps1
```

Después edita `.env` y establece al menos:

```env
SECRET_KEY=una-clave-larga-y-aleatoria
```

Para desarrollo local puede mantenerse SQLite. Para la prueba institucional se recomienda PostgreSQL.

## Smoke test

```powershell
.\.venv\Scripts\python.exe scripts\day1_smoke.py
```

Debe terminar con:

```text
DAY1_SMOKE=PASS
```

## Migraciones

```powershell
.\.venv\Scripts\python.exe -m flask --app app:create_app db upgrade
```

La instalación limpia debe finalizar en:

```text
a5d8e7f4c2b1
```

## Arranque

```powershell
.\.venv\Scripts\python.exe app.py
```

## Prueba manual mínima

1. Abrir la pantalla de login.
2. Confirmar que carga CSS y recursos.
3. Iniciar sesión con un usuario de prueba.
4. Abrir perfil.
5. Probar cambio de contraseña con contraseña actual incorrecta y confirmar rechazo.
6. Probar cambio de contraseña correcto.
7. Entrar a Evidencias.
8. Comprobar que una operación POST sin CSRF recibe HTTP 400.

## Qué debes devolverme

No hace falta enviar capturas de todo. Devuélveme:

- el resultado completo de `day1_preflight.py`;
- el resultado completo de `day1_smoke.py`;
- el resultado de `flask ... db upgrade`;
- el error completo de cualquier traceback;
- si arranca, una captura del login y otra del dashboard.

Con esos datos se puede separar inmediatamente un problema de entorno de un bug de GIA.
