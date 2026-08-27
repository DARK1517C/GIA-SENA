# Día 1 — Usuario inicial de soporte

La base de datos incluida puede arrancar correctamente, pero no contiene usuarios de forma intencional. Por eso unas credenciales de ejemplo como `soporte@gia.local` no pueden funcionar hasta crear la cuenta.

Para crear el primer usuario de soporte en una instalación local:

```powershell
python scripts\create_initial_support_user.py
```

El script solicita de manera interactiva:

- correo (por defecto `soporte@gia.local`),
- documento,
- nombres y apellidos,
- contraseña y confirmación.

La contraseña no se guarda en el código, no se imprime y no se escribe en `.env`.

El script es idempotente por protección: si ya existe un usuario con el correo o documento indicado, no modifica ningún registro.

Tras obtener `USER_CREATED=PASS`, iniciar sesión desde:

`http://127.0.0.1:5000/auth/login`


### Compatibilidad con bases SQLite preexistentes
La entidad base asigna ahora valores de fecha desde Python además del server_default. Esto permite crear usuarios en bases heredadas cuyos campos created_at/updated_at sean NOT NULL pero no tengan DEFAULT a nivel de SQLite.
