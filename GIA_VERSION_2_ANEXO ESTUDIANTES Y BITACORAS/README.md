# GIA - Gestión Integral de Aprendices

Aplicación web desarrollada con Flask, HTML, CSS y base de datos MySQL configurable para la gestión integral de aprendices del SENA.

## Análisis de requerimientos

### Objetivo general
Desarrollar una plataforma web llamada **GIA** orientada al registro, consulta, actualización, seguimiento y análisis de información académica y de etapa productiva de los aprendices SENA.

### Roles del sistema
- `Docente`: administra fichas y aprendices a cargo, consulta estadísticas, exporta/importa XLSX y gestiona bitácoras.
- `Visualizador`: consulta estadísticas globales, gestiona docentes/usuarios y descarga información consolidada.
- `Super Admin`: administra toda la plataforma, usuarios, aprendices, fichas y configuraciones operativas.
- `Aprendiz`: accede con su documento, actualiza su información personal y sube bitácoras.

### Módulos funcionales
- `Inicio de sesión`: autenticación por usuario o número de documento.
- `Panel de control`: indicadores, estadísticas y tarjetas de seguimiento etapa 1, 2 y 3.
- `Aprendices`: CRUD completo, búsqueda, filtros, vista detallada y exportación/importación XLSX.
- `Fichas`: CRUD completo, búsqueda, filtros, vista detallada y exportación/importación XLSX.
- `Usuarios`: administración de roles y cuentas.
- `Bitácoras`: carga de archivos, notas, histórico y gestión por docente o aprendiz.
- `Perfil`: edición de datos personales del aprendiz o usuario autenticado.

### Reglas clave del negocio
- Cada aprendiz debe tener un usuario automático con su documento como usuario y contraseña inicial.
- El orden de campos de aprendices y fichas debe conservarse en formularios y archivos XLSX.
- La importación XLSX debe permitir volver a exportar la misma estructura.
- La importación XLSX soporta libros con dos hojas: una hoja de aprendices y otra de fichas, similares al formato institucional recibido.
- Los aprendices solo pueden modificar información personal básica y sus bitácoras.
- Los docentes solo gestionan la información creada o asignada a su cuenta.

### Requerimientos no funcionales
- Diseño institucional inspirado en SENA, con interfaz clara, moderna y responsive.
- Arquitectura simple, mantenible y sin archivos innecesarios.
- Compatibilidad para despliegue en hosting Linux con Python, Gunicorn o servidor WSGI similar.
- Base de datos preparada para MySQL mediante variable de entorno `MYSQL_URL` o `DATABASE_URL`.

## Host, dominio y despliegue

### Recomendación de hosting
- VPS o hosting Python compatible con Flask.
- Sistema recomendado: Ubuntu 24.04 LTS.
- Servidor web: Nginx como proxy inverso.
- Ejecución de aplicación: Gunicorn.
- Base de datos: MySQL 8.

### Dominio sugerido
- `gia.sena.edu.co` o un subdominio interno institucional similar.

### Variables necesarias
- `SECRET_KEY`
- `DATABASE_URL` o `MYSQL_URL`

### Ejemplo de conexión MySQL
```env
MYSQL_URL=mysql+pymysql://usuario:clave@localhost/gia_sena
SECRET_KEY=clave-segura
```

## Inicio rápido

```bash
pip install -r requirements.txt
python app.py
```

## Usuarios demo
- `admin / admin123`
- `docente1 / docente123`
- `visualizador1 / visual123`
- `1003456789 / 1003456789`

## Formato XLSX compatible
- El sistema detecta automáticamente una hoja de aprendices con columnas como `N° DE DOCUMENTO DEL APRENDIZ`, `MODALIDAD ETAPA PRODUCTIVA` y `GESTIÓN INDIVIDUAL DEL APRENDIZ EN EP`.
- También detecta una hoja de fichas con columnas como `N° DE FICHA`, `APRENDICES EN FORMACIÓN`, `APRENDICES EN PRÁCTICA` y subcolumnas de modalidad.
- La exportación genera un libro `gia_gestion_integral.xlsx` con dos hojas y una estructura visual similar al archivo institucional de referencia.
