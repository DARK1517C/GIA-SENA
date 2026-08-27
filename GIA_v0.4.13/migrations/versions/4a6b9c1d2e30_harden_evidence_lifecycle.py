"""Harden canonical evidence lifecycle and template activity uniqueness.

Revision ID: 4a6b9c1d2e30
Revises: 3f8a7c2d1e90
"""
from alembic import op
import sqlalchemy as sa

revision = "4a6b9c1d2e30"
down_revision = "3f8a7c2d1e90"
branch_labels = None
depends_on = None


def upgrade():
    conn = op.get_bind()

    # A template-derived activity must be unique per training group.
    # Custom activities remain free to coexist. Before creating the index,
    # refuse to silently choose a survivor when bad legacy data exists.
    duplicates = conn.execute(sa.text("""
        SELECT group_id, template_id, COUNT(*) AS total
        FROM evidence_activity
        WHERE template_id IS NOT NULL
        GROUP BY group_id, template_id
        HAVING COUNT(*) > 1
    """)).fetchall()

    if duplicates:
        details = ", ".join(
            f"grupo={row[0]}, plantilla={row[1]}, filas={row[2]}"
            for row in duplicates
        )
        raise RuntimeError(
            "No se puede endurecer el dominio de evidencias: existen "
            "actividades duplicadas para la misma plantilla y ficha. "
            f"Resolver antes de migrar: {details}"
        )

    op.create_index(
        "uq_evidence_activity_group_template",
        "evidence_activity",
        ["group_id", "template_id"],
        unique=True,
        sqlite_where=sa.text("template_id IS NOT NULL"),
        postgresql_where=sa.text("template_id IS NOT NULL"),
    )


def downgrade():
    op.drop_index(
        "uq_evidence_activity_group_template",
        table_name="evidence_activity",
    )
