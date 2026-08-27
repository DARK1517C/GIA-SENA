"""Finalize canonical evidence domain and remove legacy activity type.

Revision ID: 3f8a7c2d1e90
Revises: 9c4d2e7a1b60
"""
from alembic import op
import sqlalchemy as sa

revision = "3f8a7c2d1e90"
down_revision = "9c4d2e7a1b60"
branch_labels = None
depends_on = None


def upgrade():
    conn = op.get_bind()

    # The Phase A migration made category_id mandatory. Refuse to cut over if
    # any inconsistent activity survived, rather than silently assigning data.
    unresolved = conn.execute(sa.text(
        "SELECT COUNT(*) FROM evidence_activity WHERE category_id IS NULL"
    )).scalar_one()
    if unresolved:
        raise RuntimeError(
            f"No se puede finalizar el dominio de evidencias: {unresolved} "
            "actividad(es) no tienen category_id."
        )

    # The legacy index exists in the historical schema. Drop it before the
    # column so SQLite/PostgreSQL both accept the transition.
    inspector = sa.inspect(conn)
    indexes = {
        item["name"]
        for item in inspector.get_indexes("evidence_activity")
    }
    if "ix_evidence_activity_evidence_type" in indexes:
        op.drop_index(
            "ix_evidence_activity_evidence_type",
            table_name="evidence_activity",
        )

    # evidence_type is no longer part of the canonical model.
    with op.batch_alter_table("evidence_activity") as batch:
        batch.drop_column("evidence_type")


def downgrade():
    # Reintroducing a legacy column does not restore historical values. It is
    # intentionally nullable so rollback never invents category information.
    with op.batch_alter_table("evidence_activity") as batch:
        batch.add_column(
            sa.Column(
                "evidence_type",
                sa.String(length=80),
                nullable=True,
            )
        )

    op.create_index(
        "ix_evidence_activity_evidence_type",
        "evidence_activity",
        ["evidence_type"],
        unique=False,
    )
