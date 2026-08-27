# GIA — Gestión Integral de Aprendices

GIA es una aplicación web Flask para la gestión académica, operativa y de seguimiento de aprendices SENA.

## Estado actual

La versión actual utiliza SQLAlchemy 2.x + Alembic y mantiene una estructura de persistencia reproducible desde una base de datos vacía.

### Roles canónicos

- `APPRENTICE`
- `FOLLOW_UP_INSTRUCTOR`
- `LEAD_FOLLOW_UP_INSTRUCTOR`
- `CENTER_STAFF`
- `CERTIFIER`
- `SUPPORT`

La autorización se centraliza en `services/permissions.py`.

## Módulos principales

- Autenticación y perfil.
- Aprendices.
- Grupos de formación.
- Evidencias: categorías, plantillas, actividades, entregas, revisión, aprobación, firma y comentarios.
- Usuarios.
- Estadísticas administrativas.

El aprendiz no utiliza un dashboard administrativo independiente: su flujo principal se concentra en Evidencias y las funciones permitidas por su rol.

## Persistencia

SQLite se utiliza como base de desarrollo local. El esquema se crea mediante Alembic y puede reconstruirse desde cero.

La cadena de migraciones parte de:

```text
b7e2c1a4f901
```

y actualmente continúa con la normalización del dominio de evidencias:

```text
9c4d2e7a1b60 -> 3f8a7c2d1e90 -> 4a6b9c1d2e30 -> a5d8e7f4c2b1
```

Una base limpia debe ejecutar toda la cadena y quedar en la revisión `a5d8e7f4c2b1` (head).

Para despliegue, la configuración acepta `DATABASE_URL`. Las URLs PostgreSQL `postgres://` y `postgresql://` se normalizan al driver `postgresql+psycopg://`.

### Desarrollo local

```bash
pip install -r requirements.txt
flask --app app:create_app db upgrade
python app.py
```

### PostgreSQL

Configurar, por ejemplo:

```env
DATABASE_URL=postgresql+psycopg://usuario:clave@servidor:5432/gia
SECRET_KEY=clave-segura
```

y ejecutar:

```bash
flask --app app:create_app db upgrade
```

No se requiere una segunda cadena de migraciones para PostgreSQL.

## Importación/exportación

Los servicios de Excel conservan la estructura institucional utilizada por el proyecto y deben considerarse parte del contrato funcional del módulo de aprendices y grupos.

## Estructura

- `models/` — modelos SQLAlchemy.
- `routes/` — blueprints Flask.
- `services/` — lógica de aplicación.
- `catalogs/` — catálogos y valores canónicos.
- `migrations/` — migraciones Alembic.
- `templates/` — interfaz Jinja.
- `static/` — CSS, imágenes y recursos estáticos.
- `instance/` — base SQLite local.
- `uploads/` — archivos de evidencias.
