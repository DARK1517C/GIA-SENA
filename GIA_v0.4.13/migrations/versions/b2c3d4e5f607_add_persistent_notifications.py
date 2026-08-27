"""Add persistent notifications for evidence workflow.

Revision ID: b2c3d4e5f607
Revises: a5d8e7f4c2b1
"""
from alembic import op
import sqlalchemy as sa

revision = "b2c3d4e5f607"
down_revision = "a5d8e7f4c2b1"
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        "notification",
        sa.Column("user_id", sa.Integer(), nullable=False),
        sa.Column("notification_type", sa.String(length=60), nullable=False),
        sa.Column("title", sa.String(length=160), nullable=False),
        sa.Column("message", sa.Text(), nullable=False),
        sa.Column("url", sa.String(length=500), nullable=True),
        sa.Column("is_read", sa.Boolean(), server_default=sa.false(), nullable=False),
        sa.Column("read_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), server_default=sa.func.now(), nullable=False),
        sa.ForeignKeyConstraint(["user_id"], ["user.id"]),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_notification_user_id", "notification", ["user_id"], unique=False)
    op.create_index("ix_notification_notification_type", "notification", ["notification_type"], unique=False)
    op.create_index("ix_notification_is_read", "notification", ["is_read"], unique=False)


def downgrade():
    op.drop_index("ix_notification_is_read", table_name="notification")
    op.drop_index("ix_notification_notification_type", table_name="notification")
    op.drop_index("ix_notification_user_id", table_name="notification")
    op.drop_table("notification")
