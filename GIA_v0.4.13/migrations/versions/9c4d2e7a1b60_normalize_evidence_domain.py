"""Normalize evidence domain around category/template/activity/submission.

Revision ID: 9c4d2e7a1b60
Revises: b7e2c1a4f901
"""
from alembic import op
import sqlalchemy as sa

revision = "9c4d2e7a1b60"
down_revision = "b7e2c1a4f901"
branch_labels = None
depends_on = None

CATEGORIES = [
    {"code": "initial_requirements", "name": "Requisitos Iniciales", "color": "#39a900", "sort_order": 10},
    {"code": "logs", "name": "Bitácoras", "color": "#ff6b00", "sort_order": 20},
    {"code": "followup_moments", "name": "Momentos de Seguimiento", "color": "#00a9c7", "sort_order": 30},
    {"code": "certification_requirements", "name": "Requisitos de Certificación", "color": "#e3267f", "sort_order": 40},
]

TEMPLATES = [
    ("initial_requirements", "RI-F165", "F-165 Formato seleccion modificacion alternativa etapa productiva"),
    ("initial_requirements", "RI-ARL", "Certificado de afiliacion ARL"),
    ("logs", "LOG-01", "Bitacora 1"),
    ("logs", "LOG-02", "Bitacora 2"),
    ("logs", "LOG-03", "Bitacora 3"),
    ("logs", "LOG-04", "Bitacora 4"),
    ("logs", "LOG-05", "Bitacora 5"),
    ("logs", "LOG-06", "Bitacora 6"),
    ("followup_moments", "FUP-01", "Momento 1: planeacion de la etapa productiva"),
    ("followup_moments", "FUP-02", "Momento 2: seguimiento de la etapa productiva"),
    ("followup_moments", "FUP-03", "Momento 3: evaluacion de la etapa productiva"),
    ("followup_moments", "FUP-04", "Momento 4: adicional (opcional)"),
    ("certification_requirements", "CERT-ID", "Copia de documento de Identidad"),
    ("certification_requirements", "CERT-ICFES", "Certificado de presentacion Pruebas ICFES TyT"),
    ("certification_requirements", "CERT-APE", "Certificado de la APE"),
    ("certification_requirements", "CERT-CARNET", "Carnet Destruido"),
    ("certification_requirements", "CERT-COFORMADOR", "Certificado de ente coformador aprobando finalizacion de practicas"),
]

LEGACY_TO_CODE = {
    "Requisitos Iniciales": "initial_requirements",
    "Requisitos iniciales": "initial_requirements",
    "Bitacoras": "logs",
    "Bitácoras": "logs",
    "Momentos de Seguimiento": "followup_moments",
    "Momentos de seguimiento": "followup_moments",
    "Requisitos de Certificacion": "certification_requirements",
    "Requisitos de certificación": "certification_requirements",
}


def upgrade():
    conn = op.get_bind()

    # 1. Catálogo institucional persistente.
    for item in CATEGORIES:
        conn.execute(sa.text("""
            INSERT INTO evidence_category (code, name, description, icon, color, sort_order, is_active, created_at, updated_at)
            SELECT :code, :name, NULL, 'document', :color, :sort_order, 1, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
            WHERE NOT EXISTS (SELECT 1 FROM evidence_category WHERE code = :code)
        """), item)

    # 2. Backfill de actividades existentes usando únicamente el campo legacy como puente.
    for legacy_name, code in LEGACY_TO_CODE.items():
        conn.execute(sa.text("""
            UPDATE evidence_activity
            SET category_id = (SELECT id FROM evidence_category WHERE code = :code)
            WHERE category_id IS NULL AND evidence_type = :legacy_name
        """), {"legacy_name": legacy_name, "code": code})

    unresolved = conn.execute(sa.text(
        "SELECT COUNT(*) FROM evidence_activity WHERE category_id IS NULL"
    )).scalar_one()
    if unresolved:
        raise RuntimeError(
            f"No se pudo normalizar {unresolved} actividad(es) de evidencia: "
            "su evidence_type legacy no tiene equivalencia. Corrige esos datos antes de migrar."
        )

    # 3. Plantillas oficiales. Se crean en BD, no en Python.
    category_ids = {
        row.code: row.id
        for row in conn.execute(sa.text("SELECT id, code FROM evidence_category")).mappings()
    }
    for code, template_code, title in TEMPLATES:
        conn.execute(sa.text("""
            INSERT INTO evidence_template
                (category_id, code, title, description, allowed_extensions, max_file_size_mb,
                 requires_signature, is_required, sort_order, is_active, created_by_id, created_at, updated_at)
            SELECT :category_id, :template_code, :title, NULL, NULL, NULL,
                   0, 1, :sort_order, 1, NULL, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
            WHERE NOT EXISTS (SELECT 1 FROM evidence_template WHERE code = :template_code)
        """), {
            "category_id": category_ids[code],
            "template_code": template_code,
            "title": title,
            "sort_order": TEMPLATES.index((code, template_code, title)) + 1,
        })

    # 4. Vincular actividades legacy que coinciden con una plantilla oficial.
    conn.execute(sa.text("""
        UPDATE evidence_activity
        SET template_id = (
            SELECT et.id
            FROM evidence_template et
            WHERE et.category_id = evidence_activity.category_id
              AND et.title = evidence_activity.title
            LIMIT 1
        )
        WHERE template_id IS NULL
          AND EXISTS (
            SELECT 1 FROM evidence_template et
            WHERE et.category_id = evidence_activity.category_id
              AND et.title = evidence_activity.title
          )
    """))

    # 5. category_id pasa a ser obligatorio. evidence_type queda nullable como puente
    #    de una sola generación; la siguiente fase podrá eliminarlo físicamente.
    with op.batch_alter_table("evidence_activity") as batch:
        batch.alter_column("category_id", existing_type=sa.Integer(), nullable=False)
        batch.alter_column("evidence_type", existing_type=sa.String(length=80), nullable=True)


def downgrade():
    # No eliminamos datos institucionales en downgrade; solo restauramos la nulabilidad.
    with op.batch_alter_table("evidence_activity") as batch:
        batch.alter_column("category_id", existing_type=sa.Integer(), nullable=True)
