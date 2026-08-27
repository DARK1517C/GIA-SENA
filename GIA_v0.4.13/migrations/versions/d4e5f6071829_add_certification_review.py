"""Add certification review audit trail."""
from alembic import op
import sqlalchemy as sa

revision = "d4e5f6071829"
down_revision = "c3d4e5f60718"
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        "certification_review",
        sa.Column("apprentice_id", sa.Integer(), nullable=False),
        sa.Column("reviewer_id", sa.Integer(), nullable=False),
        sa.Column("status", sa.String(length=20), server_default="PENDING", nullable=False),
        sa.Column("notes", sa.Text(), nullable=True),
        sa.Column("reviewed_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("id", sa.Integer(), autoincrement=True, nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.current_timestamp(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.current_timestamp(), nullable=False),
        sa.ForeignKeyConstraint(["apprentice_id"], ["apprentice.id"]),
        sa.ForeignKeyConstraint(["reviewer_id"], ["user.id"]),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_certification_review_apprentice", "certification_review", ["apprentice_id"], unique=False)
    op.create_index("ix_certification_review_reviewer_id", "certification_review", ["reviewer_id"], unique=False)
    op.create_index("ix_certification_review_status", "certification_review", ["status"], unique=False)


def downgrade():
    op.drop_index("ix_certification_review_status", table_name="certification_review")
    op.drop_index("ix_certification_review_reviewer_id", table_name="certification_review")
    op.drop_index("ix_certification_review_apprentice", table_name="certification_review")
    op.drop_table("certification_review")
