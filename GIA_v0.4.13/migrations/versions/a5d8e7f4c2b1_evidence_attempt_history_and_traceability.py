"""Create real evidence attempts and remove observations text history.

Revision ID: a5d8e7f4c2b1
Revises: 4a6b9c1d2e30
"""
from alembic import op
import sqlalchemy as sa

revision = "a5d8e7f4c2b1"
down_revision = "4a6b9c1d2e30"
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        "evidence_submission_attempt",
        sa.Column("submission_id", sa.Integer(), nullable=False),
        sa.Column("attempt_number", sa.Integer(), nullable=False),
        sa.Column("version_number", sa.Integer(), nullable=False),
        sa.Column("status", sa.String(length=40), nullable=False, server_default="pendiente_revision"),
        sa.Column("file_name", sa.String(length=255), nullable=True),
        sa.Column("file_path", sa.String(length=255), nullable=True),
        sa.Column("mime_type", sa.String(length=120), nullable=True),
        sa.Column("file_size_bytes", sa.Integer(), nullable=True),
        sa.Column("uploaded_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("reviewed_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("reviewed_by", sa.Integer(), nullable=True),
        sa.Column("approved_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("approved_by_id", sa.Integer(), nullable=True),
        sa.Column("signed_file_name", sa.String(length=255), nullable=True),
        sa.Column("signed_file_path", sa.String(length=255), nullable=True),
        sa.Column("signed_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=sa.text("CURRENT_TIMESTAMP")),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=sa.text("CURRENT_TIMESTAMP")),
        sa.ForeignKeyConstraint(["submission_id"], ["evidence_submission.id"]),
        sa.ForeignKeyConstraint(["reviewed_by"], ["user.id"]),
        sa.ForeignKeyConstraint(["approved_by_id"], ["user.id"]),
    )
    op.create_index("ix_evidence_submission_attempt_submission_id", "evidence_submission_attempt", ["submission_id"])
    op.create_index("ix_evidence_submission_attempt_status", "evidence_submission_attempt", ["status"])

    conn = op.get_bind()
    conn.execute(sa.text("""
        INSERT INTO evidence_submission_attempt
            (submission_id, attempt_number, version_number, status, file_name, file_path, mime_type, file_size_bytes, uploaded_at, reviewed_at, reviewed_by, approved_at, approved_by_id, signed_file_name, signed_file_path, signed_at, created_at, updated_at)
        SELECT id, CASE WHEN attempt_number < 1 THEN 1 ELSE attempt_number END,
               CASE WHEN version_number < 1 THEN 1 ELSE version_number END,
               CASE WHEN file_path IS NULL THEN 'no_entregado' ELSE status END,
               file_name, file_path, mime_type, file_size_bytes, uploaded_at, reviewed_at, reviewed_by, approved_at, approved_by_id, signed_file_name, signed_file_path, signed_at, created_at, updated_at
        FROM evidence_submission
    """))

    # Legacy observations become structured comments. Author may be unknown;
    # null is preferable to inventing an identity.
    conn.execute(sa.text("""
        INSERT INTO evidence_comment (submission_id, author_id, comment, is_internal, created_at, updated_at)
        SELECT id, COALESCE(reviewed_by, approved_by_id), observations, 0, COALESCE(updated_at, CURRENT_TIMESTAMP), COALESCE(updated_at, CURRENT_TIMESTAMP)
        FROM evidence_submission
        WHERE observations IS NOT NULL AND TRIM(observations) <> ''
    """))

    with op.batch_alter_table("evidence_comment") as batch:
        batch.alter_column("author_id", existing_type=sa.Integer(), nullable=True)

    with op.batch_alter_table("evidence_submission") as batch:
        batch.drop_column("observations")


def downgrade():
    with op.batch_alter_table("evidence_submission") as batch:
        batch.add_column(sa.Column("observations", sa.Text(), nullable=True))
    op.drop_index("ix_evidence_submission_attempt_status", table_name="evidence_submission_attempt")
    op.drop_index("ix_evidence_submission_attempt_submission_id", table_name="evidence_submission_attempt")
    op.drop_table("evidence_submission_attempt")
