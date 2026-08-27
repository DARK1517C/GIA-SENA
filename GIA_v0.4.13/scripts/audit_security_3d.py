"""Auditoría estática del Bloque Seguridad 3.D.

No requiere Flask/SQLAlchemy: comprueba invariantes de enforcement para que
las rutas sensibles no dependan únicamente de la interfaz.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

def text(rel):
    return (ROOT / rel).read_text(encoding="utf-8")

def require(condition, message):
    if not condition:
        raise SystemExit(f"FAIL: {message}")
    print(f"OK: {message}")

evid = text("routes/evidences.py")
groups = text("routes/groups.py")
appr = text("routes/apprentices.py")
imp = text("services/excel_import.py")
scope = text("services/access_scope.py")
perms = text("services/permissions.py")

# Recursos individuales sensibles deben comprobar alcance explícito.
for fn in ("observe", "approve", "sign_submission"):
    start = evid.index(f"def {fn}(")
    end = evid.find("\ndef ", start + 5)
    block = evid[start:end if end != -1 else None]
    require("_check_submission_access(" in block,
            f"evidencias.{fn} aplica alcance por registro")

# CRUD de grupos/aprendices debe filtrar el registro antes de mutarlo.
for fn in ("edit", "delete"):
    start = groups.index(f"def {fn}(")
    end = groups.find("\ndef ", start + 5)
    block = groups[start:end if end != -1 else None]
    require("_can_manage_group(" in block,
            f"groups.{fn} valida alcance del grupo")

for fn in ("edit", "delete"):
    start = appr.index(f"def {fn}(")
    end = appr.find("\ndef ", start + 5)
    block = appr[start:end if end != -1 else None]
    require("_get_visible_apprentice_or_404(" in block,
            f"apprentices.{fn} carga únicamente registros visibles")

# Operaciones masivas de importación no pueden saltarse el alcance.
require("group_scope=None" in imp,
        "el importador acepta un callback de alcance")
require("group_scope is not None and not group_scope" in imp,
        "el importador bloquea filas fuera de alcance")
require("group_scope=" in groups,
        "importación de grupos pasa alcance")
require("group_scope=" in appr,
        "importación de aprendices pasa alcance")

# Creación de fichas por instructor normal debe quedar asignada a su propio seguimiento.
start = groups.index("def create(")
end = groups.find("\ndef ", start + 5)
create_block = groups[start:end if end != -1 else None]
require("_is_followup_instructor()" in create_block and "_scope_group_identity()" in create_block,
        "creación de grupos restringida al alcance del instructor")

# El alcance canónico continúa separado de los permisos.
require("visible_groups_query" in scope and "PERMISSIONS" not in scope,
        "alcance permanece separado del catálogo de permisos")
require('"groups.manage"' in perms and '"evidences.approve"' in perms,
        "permisos canónicos siguen centralizados")

print("SECURITY_3D_AUDIT=PASS")
