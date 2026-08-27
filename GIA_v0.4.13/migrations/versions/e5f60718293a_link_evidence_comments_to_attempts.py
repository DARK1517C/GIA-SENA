"""Link evidence comments to the delivery attempt they reference.

Revision ID: e5f60718293a
Revises: d4e5f6071829
"""
from alembic import op
import sqlalchemy as sa

revision = "e5f60718293a"
down_revision = "d4e5f6071829"
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table("evidence_comment", schema=None) as batch:
        batch.add_column(sa.Column("attempt_id", sa.Integer(), nullable=True))
        batch.create_index("ix_evidence_comment_attempt_id", ["attempt_id"], unique=False)
        batch.create_foreign_key(
            "fk_evidence_comment_attempt_id",
            "evidence_submission_attempt",
            ["attempt_id"],
            ["id"],
        )


def downgrade():
    with op.batch_alter_table("evidence_comment", schema=None) as batch:
        batch.drop_constraint("fk_evidence_comment_attempt_id", type_="foreignkey")
        batch.drop_index("ix_evidence_comment_attempt_id")
        batch.drop_column("attempt_id")
