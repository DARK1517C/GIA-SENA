"""Add correction-request flag to evidence comments.

Revision ID: c3d4e5f60718
Revises: b2c3d4e5f607
"""
from alembic import op
import sqlalchemy as sa

revision = "c3d4e5f60718"
down_revision = "b2c3d4e5f607"
branch_labels = None
depends_on = None

def upgrade():
    op.add_column("evidence_comment", sa.Column("is_correction_request", sa.Boolean(), server_default=sa.false(), nullable=False))
    op.create_index("ix_evidence_comment_is_correction_request", "evidence_comment", ["is_correction_request"], unique=False)

def downgrade():
    op.drop_index("ix_evidence_comment_is_correction_request", table_name="evidence_comment")
    op.drop_column("evidence_comment", "is_correction_request")
