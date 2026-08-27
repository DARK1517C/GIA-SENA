"""Create the current GIA schema from an empty database.

Revision ID: b7e2c1a4f901
Revises:
Create Date: 2026-08-19

This is the clean baseline for the current SQLAlchemy models.  It intentionally
replaces the historical migration chain, which could not reliably bootstrap an
empty database because it assumed legacy tables/data.
"""

from alembic import op
import sqlalchemy as sa


revision = "b7e2c1a4f901"
down_revision = None
branch_labels = None
depends_on = None


TRUE = sa.text("true")
FALSE = sa.text("false")
CURRENT_TIMESTAMP = sa.text("CURRENT_TIMESTAMP")


def upgrade():
    op.create_table(
        "user",
        sa.Column("document_type", sa.String(30), nullable=False),
        sa.Column("document_number", sa.String(30), nullable=False),
        sa.Column("first_names", sa.String(120), nullable=False),
        sa.Column("last_names", sa.String(120), nullable=False),
        sa.Column("email", sa.String(255), nullable=True),
        sa.Column("phone", sa.String(30), nullable=True),
        sa.Column("role", sa.String(60), nullable=False),
        sa.Column("status", sa.String(30), nullable=False, server_default=sa.text("'ACTIVE'")),
        sa.Column("password_hash", sa.String(255), nullable=False),
        sa.Column("signature_file_name", sa.String(255), nullable=True),
        sa.Column("signature_file_path", sa.String(255), nullable=True),
        sa.Column("signature_updated_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("last_login_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.UniqueConstraint("document_number", name="uq_user_document_number"),
        sa.UniqueConstraint("email", name="uq_user_email"),
    )

    op.create_table(
        "evidence_category",
        sa.Column("code", sa.String(80), nullable=False),
        sa.Column("name", sa.String(120), nullable=False),
        sa.Column("description", sa.Text(), nullable=True),
        sa.Column("icon", sa.String(80), nullable=True),
        sa.Column("color", sa.String(20), nullable=True),
        sa.Column("sort_order", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("is_active", sa.Boolean(), nullable=False, server_default=TRUE),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.UniqueConstraint("code", name="uq_evidence_category_code"),
        sa.UniqueConstraint("name", name="uq_evidence_category_name"),
    )

    op.create_table(
        "training_group",
        sa.Column("created_by", sa.Integer(), nullable=False),
        sa.Column("group_number", sa.String(30), nullable=False),
        sa.Column("program_name", sa.String(150), nullable=False),
        sa.Column("lead_instructor", sa.String(150), nullable=True),
        sa.Column("followup_instructor", sa.String(150), nullable=True),
        sa.Column("municipality", sa.String(120), nullable=True),
        sa.Column("program_level", sa.String(80), nullable=True),
        sa.Column("modality", sa.String(80), nullable=True),
        sa.Column("sofia_group_status", sa.String(80), nullable=True),
        sa.Column("group_validity", sa.String(80), nullable=True),
        sa.Column("group_start_date", sa.String(40), nullable=True),
        sa.Column("training_end_date", sa.String(40), nullable=True),
        sa.Column("ep_start_date", sa.String(40), nullable=True),
        sa.Column("apprentices_statistics", sa.String(120), nullable=True),
        sa.Column("apprentices_training", sa.String(30), nullable=True),
        sa.Column("apprentices_enabled", sa.String(30), nullable=True),
        sa.Column("apprentices_rap_pending", sa.String(30), nullable=True),
        sa.Column("apprentices_practice", sa.String(30), nullable=True),
        sa.Column("apprentices_without_alternative", sa.String(30), nullable=True),
        sa.Column("apprentices_certified", sa.String(30), nullable=True),
        sa.Column("productive_modalities", sa.String(120), nullable=True),
        sa.Column("learning_contract", sa.String(30), nullable=True),
        sa.Column("internship", sa.String(30), nullable=True),
        sa.Column("productive_project", sa.String(30), nullable=True),
        sa.Column("employment_link", sa.String(30), nullable=True),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.ForeignKeyConstraint(["created_by"], ["user.id"], name="fk_training_group_created_by_user"),
        sa.UniqueConstraint("group_number", name="uq_training_group_group_number"),
    )

    op.create_table(
        "apprentice",
        sa.Column("created_by", sa.Integer(), nullable=False),
        sa.Column("student_user_id", sa.Integer(), nullable=True),
        sa.Column("group_id", sa.Integer(), nullable=True),
        sa.Column("group_number", sa.String(30), nullable=False),
        sa.Column("document_type", sa.String(30), nullable=False),
        sa.Column("document_number", sa.String(30), nullable=False),
        sa.Column("first_names", sa.String(120), nullable=False),
        sa.Column("last_names", sa.String(120), nullable=False),
        sa.Column("gender", sa.String(20), nullable=True),
        sa.Column("phone", sa.String(30), nullable=True),
        sa.Column("email", sa.String(150), nullable=True),
        sa.Column("municipality_origin", sa.String(120), nullable=True),
        sa.Column("program_name", sa.String(150), nullable=True),
        sa.Column("program_level", sa.String(80), nullable=True),
        sa.Column("group_validity", sa.String(80), nullable=True),
        sa.Column("lead_instructor", sa.String(150), nullable=True),
        sa.Column("followup_instructor", sa.String(150), nullable=True),
        sa.Column("followup_instructor_email", sa.String(150), nullable=True),
        sa.Column("ep_modality", sa.String(120), nullable=True),
        sa.Column("sofia_status", sa.String(80), nullable=True),
        sa.Column("practice_start_date", sa.String(40), nullable=True),
        sa.Column("practice_end_date", sa.String(40), nullable=True),
        sa.Column("followup_moment1_start", sa.String(40), nullable=True),
        sa.Column("followup_moment1_end", sa.String(40), nullable=True),
        sa.Column("followup_moment2_start", sa.String(40), nullable=True),
        sa.Column("followup_moment2_end", sa.String(40), nullable=True),
        sa.Column("followup_moment3_start", sa.String(40), nullable=True),
        sa.Column("followup_moment3_end", sa.String(40), nullable=True),
        sa.Column("followup_moment4_start", sa.String(40), nullable=True),
        sa.Column("followup_moment4_end", sa.String(40), nullable=True),
        sa.Column("company_name", sa.String(150), nullable=True),
        sa.Column("company_municipality", sa.String(120), nullable=True),
        sa.Column("company_address", sa.String(180), nullable=True),
        sa.Column("coformador_name", sa.String(150), nullable=True),
        sa.Column("coformador_email", sa.String(150), nullable=True),
        sa.Column("coformador_phone", sa.String(30), nullable=True),
        sa.Column("arl_responsible", sa.String(150), nullable=True),
        sa.Column("continues_company", sa.String(30), nullable=True),
        sa.Column("individual_management", sa.Text(), nullable=True),
        sa.Column("followup_moments", sa.String(200), nullable=True),
        sa.Column("evaluation_date", sa.String(40), nullable=True),
        sa.Column("english_results", sa.String(120), nullable=True),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.ForeignKeyConstraint(["created_by"], ["user.id"], name="fk_apprentice_created_by_user"),
        sa.ForeignKeyConstraint(["student_user_id"], ["user.id"], name="fk_apprentice_student_user_id_user"),
        sa.ForeignKeyConstraint(["group_id"], ["training_group.id"], name="fk_apprentice_group_id_training_group"),
        sa.UniqueConstraint("document_number", name="uq_apprentice_document_number"),
    )

    op.create_table(
        "evidence_template",
        sa.Column("category_id", sa.Integer(), nullable=False),
        sa.Column("code", sa.String(120), nullable=False),
        sa.Column("title", sa.String(180), nullable=False),
        sa.Column("description", sa.Text(), nullable=True),
        sa.Column("allowed_extensions", sa.Text(), nullable=True),
        sa.Column("max_file_size_mb", sa.Integer(), nullable=True),
        sa.Column("requires_signature", sa.Boolean(), nullable=False, server_default=FALSE),
        sa.Column("is_required", sa.Boolean(), nullable=False, server_default=TRUE),
        sa.Column("sort_order", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("is_active", sa.Boolean(), nullable=False, server_default=TRUE),
        sa.Column("created_by_id", sa.Integer(), nullable=True),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.ForeignKeyConstraint(["category_id"], ["evidence_category.id"], name="fk_evidence_template_category_id_evidence_category"),
        sa.ForeignKeyConstraint(["created_by_id"], ["user.id"], name="fk_evidence_template_created_by_id_user"),
        sa.UniqueConstraint("code", name="uq_evidence_template_code"),
    )

    op.create_table(
        "evidence_activity",
        sa.Column("group_id", sa.Integer(), nullable=False),
        sa.Column("template_id", sa.Integer(), nullable=True),
        sa.Column("category_id", sa.Integer(), nullable=True),
        sa.Column("evidence_type", sa.String(80), nullable=False),
        sa.Column("code", sa.String(120), nullable=True),
        sa.Column("title", sa.String(180), nullable=False),
        sa.Column("description", sa.Text(), nullable=True),
        sa.Column("due_start", sa.String(40), nullable=True),
        sa.Column("due_end", sa.String(40), nullable=True),
        sa.Column("allowed_extensions", sa.Text(), nullable=True),
        sa.Column("max_file_size_mb", sa.Integer(), nullable=True),
        sa.Column("requires_signature", sa.Boolean(), nullable=False, server_default=FALSE),
        sa.Column("is_required", sa.Boolean(), nullable=False, server_default=TRUE),
        sa.Column("is_visible", sa.Boolean(), nullable=False, server_default=TRUE),
        sa.Column("is_default", sa.Boolean(), nullable=False, server_default=TRUE),
        sa.Column("origin", sa.String(20), nullable=False, server_default=sa.text("'template'")),
        sa.Column("sort_order", sa.Integer(), nullable=False, server_default="0"),
        sa.Column("created_by_id", sa.Integer(), nullable=True),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.ForeignKeyConstraint(["group_id"], ["training_group.id"], name="fk_evidence_activity_group_id_training_group"),
        sa.ForeignKeyConstraint(["template_id"], ["evidence_template.id"], name="fk_evidence_activity_template_id_evidence_template"),
        sa.ForeignKeyConstraint(["category_id"], ["evidence_category.id"], name="fk_evidence_activity_category_id_evidence_category"),
        sa.ForeignKeyConstraint(["created_by_id"], ["user.id"], name="fk_evidence_activity_created_by_id_user"),
    )

    op.create_table(
        "evidence_submission",
        sa.Column("activity_id", sa.Integer(), nullable=False),
        sa.Column("apprentice_id", sa.Integer(), nullable=False),
        sa.Column("status", sa.String(40), nullable=False, server_default=sa.text("'no_entregado'")),
        sa.Column("observations", sa.Text(), nullable=True),
        sa.Column("file_name", sa.String(255), nullable=True),
        sa.Column("file_path", sa.String(255), nullable=True),
        sa.Column("mime_type", sa.String(120), nullable=True),
        sa.Column("file_size_bytes", sa.Integer(), nullable=True),
        sa.Column("uploaded_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("reviewed_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("reviewed_by", sa.Integer(), nullable=True),
        sa.Column("approved_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("approved_by_id", sa.Integer(), nullable=True),
        sa.Column("signed_file_name", sa.String(255), nullable=True),
        sa.Column("signed_file_path", sa.String(255), nullable=True),
        sa.Column("signed_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("version_number", sa.Integer(), nullable=False, server_default="1"),
        sa.Column("attempt_number", sa.Integer(), nullable=False, server_default="1"),
        sa.Column("is_latest", sa.Boolean(), nullable=False, server_default=TRUE),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.ForeignKeyConstraint(["activity_id"], ["evidence_activity.id"], name="fk_evidence_submission_activity_id_evidence_activity"),
        sa.ForeignKeyConstraint(["apprentice_id"], ["apprentice.id"], name="fk_evidence_submission_apprentice_id_apprentice"),
        sa.ForeignKeyConstraint(["reviewed_by"], ["user.id"], name="fk_evidence_submission_reviewed_by_user"),
        sa.ForeignKeyConstraint(["approved_by_id"], ["user.id"], name="fk_evidence_submission_approved_by_id_user"),
    )

    op.create_table(
        "evidence_comment",
        sa.Column("submission_id", sa.Integer(), nullable=False),
        sa.Column("author_id", sa.Integer(), nullable=False),
        sa.Column("comment", sa.Text(), nullable=False),
        sa.Column("is_internal", sa.Boolean(), nullable=False, server_default=FALSE),
        sa.Column("id", sa.Integer(), primary_key=True, autoincrement=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False, server_default=CURRENT_TIMESTAMP),
        sa.ForeignKeyConstraint(["submission_id"], ["evidence_submission.id"], name="fk_evidence_comment_submission_id_evidence_submission"),
        sa.ForeignKeyConstraint(["author_id"], ["user.id"], name="fk_evidence_comment_author_id_user"),
    )

    # Column indexes declared by the current models.
    indexes = {
        "user": [
            ("ix_user_document_type", ["document_type"], False),
            ("ix_user_first_names", ["first_names"], False),
            ("ix_user_last_names", ["last_names"], False),
            ("ix_user_role", ["role"], False),
            ("ix_user_status", ["status"], False),
            ("ix_user_document_number", ["document_number"], True),
            ("ix_user_email", ["email"], True),
        ],
        "evidence_category": [
            ("ix_evidence_category_code", ["code"], True),
            ("ix_evidence_category_name", ["name"], True),
            ("ix_evidence_category_is_active", ["is_active"], False),
        ],
        "training_group": [
            ("ix_training_group_created_by", ["created_by"], False),
            ("ix_training_group_group_number", ["group_number"], True),
            ("ix_training_group_program_name", ["program_name"], False),
            ("ix_training_group_municipality", ["municipality"], False),
            ("ix_training_group_program_level", ["program_level"], False),
            ("ix_training_group_modality", ["modality"], False),
            ("ix_training_group_sofia_group_status", ["sofia_group_status"], False),
        ],
        "apprentice": [
            ("ix_apprentice_created_by", ["created_by"], False),
            ("ix_apprentice_student_user_id", ["student_user_id"], False),
            ("ix_apprentice_group_id", ["group_id"], False),
            ("ix_apprentice_group_number", ["group_number"], False),
            ("ix_apprentice_document_type", ["document_type"], False),
            ("ix_apprentice_document_number", ["document_number"], True),
            ("ix_apprentice_gender", ["gender"], False),
            ("ix_apprentice_email", ["email"], False),
            ("ix_apprentice_municipality_origin", ["municipality_origin"], False),
            ("ix_apprentice_program_name", ["program_name"], False),
            ("ix_apprentice_program_level", ["program_level"], False),
            ("ix_apprentice_ep_modality", ["ep_modality"], False),
            ("ix_apprentice_sofia_status", ["sofia_status"], False),
        ],
        "evidence_template": [
            ("ix_evidence_template_category_id", ["category_id"], False),
            ("ix_evidence_template_code", ["code"], True),
            ("ix_evidence_template_title", ["title"], False),
            ("ix_evidence_template_is_active", ["is_active"], False),
            ("ix_evidence_template_created_by_id", ["created_by_id"], False),
        ],
        "evidence_activity": [
            ("ix_evidence_activity_group_id", ["group_id"], False),
            ("ix_evidence_activity_template_id", ["template_id"], False),
            ("ix_evidence_activity_category_id", ["category_id"], False),
            ("ix_evidence_activity_evidence_type", ["evidence_type"], False),
            ("ix_evidence_activity_code", ["code"], False),
            ("ix_evidence_activity_title", ["title"], False),
            ("ix_evidence_activity_origin", ["origin"], False),
            ("ix_evidence_activity_created_by_id", ["created_by_id"], False),
        ],
        "evidence_submission": [
            ("ix_evidence_submission_activity_id", ["activity_id"], False),
            ("ix_evidence_submission_apprentice_id", ["apprentice_id"], False),
            ("ix_evidence_submission_status", ["status"], False),
            ("ix_evidence_submission_mime_type", ["mime_type"], False),
            ("ix_evidence_submission_reviewed_by", ["reviewed_by"], False),
            ("ix_evidence_submission_approved_by_id", ["approved_by_id"], False),
            ("ix_evidence_submission_is_latest", ["is_latest"], False),
        ],
        "evidence_comment": [
            ("ix_evidence_comment_submission_id", ["submission_id"], False),
            ("ix_evidence_comment_author_id", ["author_id"], False),
            ("ix_evidence_comment_is_internal", ["is_internal"], False),
        ],
    }

    for table_name, table_indexes in indexes.items():
        for name, columns, unique in table_indexes:
            op.create_index(name, table_name, columns, unique=unique)

    # A submission marked as latest is unique for each activity/apprentice pair.
    op.create_index(
        "uq_evidence_submission_latest_per_activity_apprentice",
        "evidence_submission",
        ["activity_id", "apprentice_id"],
        unique=True,
        sqlite_where=sa.text("is_latest = 1"),
        postgresql_where=sa.text("is_latest = true"),
    )


def downgrade():
    op.drop_index("uq_evidence_submission_latest_per_activity_apprentice", table_name="evidence_submission")

    for table_name, table_indexes in {
        "evidence_comment": [
            "ix_evidence_comment_is_internal", "ix_evidence_comment_author_id", "ix_evidence_comment_submission_id"],
        "evidence_submission": [
            "ix_evidence_submission_is_latest", "ix_evidence_submission_approved_by_id", "ix_evidence_submission_reviewed_by",
            "ix_evidence_submission_mime_type", "ix_evidence_submission_status", "ix_evidence_submission_apprentice_id", "ix_evidence_submission_activity_id"],
        "evidence_activity": [
            "ix_evidence_activity_created_by_id", "ix_evidence_activity_origin", "ix_evidence_activity_title", "ix_evidence_activity_code",
            "ix_evidence_activity_evidence_type", "ix_evidence_activity_category_id", "ix_evidence_activity_template_id", "ix_evidence_activity_group_id"],
        "evidence_template": [
            "ix_evidence_template_created_by_id", "ix_evidence_template_is_active", "ix_evidence_template_title", "ix_evidence_template_code", "ix_evidence_template_category_id"],
        "apprentice": [
            "ix_apprentice_sofia_status", "ix_apprentice_ep_modality", "ix_apprentice_program_level", "ix_apprentice_program_name",
            "ix_apprentice_municipality_origin", "ix_apprentice_email", "ix_apprentice_gender", "ix_apprentice_document_number",
            "ix_apprentice_document_type", "ix_apprentice_group_number", "ix_apprentice_group_id", "ix_apprentice_student_user_id", "ix_apprentice_created_by"],
        "training_group": [
            "ix_training_group_sofia_group_status", "ix_training_group_modality", "ix_training_group_program_level", "ix_training_group_municipality",
            "ix_training_group_program_name", "ix_training_group_group_number", "ix_training_group_created_by"],
        "evidence_category": ["ix_evidence_category_is_active", "ix_evidence_category_name", "ix_evidence_category_code"],
        "user": ["ix_user_email", "ix_user_document_number", "ix_user_status", "ix_user_role", "ix_user_last_names", "ix_user_first_names", "ix_user_document_type"],
    }.items():
        for name in table_indexes:
            op.drop_index(name, table_name=table_name)

    for table_name in [
        "evidence_comment", "evidence_submission", "evidence_activity", "evidence_template",
        "apprentice", "training_group", "evidence_category", "user",
    ]:
        op.drop_table(table_name)
