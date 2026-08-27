# Bloque Seguridad 3.D — E2E real de los seis roles

## Objetivo

Validar por HTTP, con `Flask.test_client()`, la matriz de autorización y el
alcance de registros sin tocar la BD de desarrollo.

## Cobertura

Se preparan seis usuarios reales en una BD SQLite temporal:

- Aprendiz
- Instructor de seguimiento
- Instructor de seguimiento líder
- Administrativo del centro
- Certificador
- Soporte

También se crean dos fichas (una propia del instructor y otra ajena), dos
aprendices y dos entregas de evidencia para probar aislamiento por registro.

### Casos

1. Login HTTP de los seis roles.
2. Acceso a creación de fichas.
3. Acceso a administración de usuarios.
4. Acceso al catálogo de categorías de evidencias.
5. Instructor: acceso a recursos propios y rechazo de recursos ajenos.
6. Aprendiz: aislamiento de sus propios recursos.
7. Roles globales: lectura de evidencia ajena.
8. Instructor: observar evidencia propia, rechazo de ajena.
9. Administrativo: sin modificación de evidencias.
10. Certificador: aprobación global y rechazo de firma.
11. Soporte: gestión global y rechazo de firma.
12. Administración de usuarios exclusivamente para Soporte.

## Resultado del entorno de auditoría

La suite E2E fue preparada y validada sintácticamente, pero **no pudo
 ejecutarse en este entorno de análisis** porque no dispone de Flask/SQLAlchemy
y no tiene acceso de red para instalar `requirements.txt`.

La ejecución intentada produjo:

`ModuleNotFoundError: No module named 'flask'`

Y la instalación falló por ausencia de acceso al índice de paquetes.

Por tanto, **NO se declara PASS E2E**. El único PASS confirmado aquí es la
auditoría estática 3.D (`SECURITY_3D_AUDIT=PASS`) y la compilación Python.

## Ejecución en el proyecto local

Desde la raíz de GIA:

```powershell
python -m pytest -q tests/test_security_3d_e2e.py
```

La prueba usa una BD SQLite temporal y no utiliza `instance/gia.db`.

## Criterio de cierre

3.D E2E debe considerarse cerrado únicamente cuando la orden anterior termine
con todos los tests en verde. Si aparece un fallo, se corrige antes de pasar
al Bloque 4.
