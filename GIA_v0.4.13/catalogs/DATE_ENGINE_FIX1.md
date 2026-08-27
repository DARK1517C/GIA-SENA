# v0.4.10 FIX1

Corrección puntual del arranque de v0.4.10:
- Eliminada referencia obsoleta `TrainingRelationship.PASANTIA` de `catalogs/aliases.py`.
- `TrainingRelationship` conserva únicamente los valores definidos en `catalogs/apprentice.py`.
- No se modifica la BD ni se requiere migración.
