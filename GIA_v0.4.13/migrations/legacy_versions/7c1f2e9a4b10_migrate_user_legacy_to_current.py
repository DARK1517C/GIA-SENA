"""Migrate legacy user table to current User model.

Revision ID: 7c1f2e9a4b10
Revises: 93671fcbae56
"""

from alembic import op
import sqlalchemy as sa


revision = "7c1f2e9a4b10"
down_revision = "93671fcbae56"
branch_labels = None
depends_on = None


def upgrade():
    conn = op.get_bind()

    # ============================================================
    # 1. Eliminar tabla temporal dejada por la migración anterior
    # ============================================================

    conn.execute(
        sa.text("DROP TABLE IF EXISTS _alembic_tmp_apprentice")
    )

    # ============================================================
    # 2. Crear tabla user nueva
    # ============================================================

    op.create_table(
        "user_new",

        sa.Column(
            "id",
            sa.Integer(),
            primary_key=True,
            nullable=False,
        ),

        sa.Column(
            "document_type",
            sa.String(length=30),
            nullable=False,
        ),

        sa.Column(
            "document_number",
            sa.String(length=30),
            nullable=False,
        ),

        sa.Column(
            "first_names",
            sa.String(length=120),
            nullable=False,
        ),

        sa.Column(
            "last_names",
            sa.String(length=120),
            nullable=False,
        ),

        sa.Column(
            "email",
            sa.String(length=255),
            nullable=True,
        ),

        sa.Column(
            "phone",
            sa.String(length=30),
            nullable=True,
        ),

        sa.Column(
            "role",
            sa.String(length=60),
            nullable=False,
        ),

        sa.Column(
            "status",
            sa.String(length=30),
            nullable=False,
        ),

        sa.Column(
            "password_hash",
            sa.String(length=255),
            nullable=False,
        ),

        sa.Column(
            "signature_file_name",
            sa.String(length=255),
            nullable=True,
        ),

        sa.Column(
            "signature_file_path",
            sa.String(length=255),
            nullable=True,
        ),

        sa.Column(
            "signature_updated_at",
            sa.DateTime(timezone=True),
            nullable=True,
        ),

        sa.Column(
            "last_login_at",
            sa.DateTime(timezone=True),
            nullable=True,
        ),
    )

    # ============================================================
    # 3. Leer usuarios existentes
    # ============================================================

    users = conn.execute(
        sa.text(
            """
            SELECT
                id,
                username,
                password_hash,
                role,
                full_name,
                email,
                document_type,
                document_number,
                active,
                created_at
            FROM user
            ORDER BY id
            """
        )
    ).mappings().all()

    used_document_numbers = set()
    used_emails = set()

    # ============================================================
    # 4. Migrar usuarios
    # ============================================================

    for user in users:

        user_id = user["id"]

        username = (
            str(user["username"]).strip()
            if user["username"] is not None
            else ""
        )

        full_name = (
            str(user["full_name"]).strip()
            if user["full_name"] is not None
            else ""
        )

        # --------------------------------------------------------
        # Email
        # --------------------------------------------------------

        email = (
            str(user["email"]).strip().lower()
            if user["email"]
            else None
        )

        # La columna nueva es UNIQUE.
        # Si hubiera duplicados, conservamos el primero y
        # dejamos NULL en los siguientes.
        if email:
            if email in used_emails:
                email = None
            else:
                used_emails.add(email)

        # --------------------------------------------------------
        # Documento
        # --------------------------------------------------------

        document_type = (
            str(user["document_type"]).strip().upper()
            if user["document_type"]
            else "NATIONAL_ID"
        )

        document_number = (
            str(user["document_number"]).strip().upper()
            if user["document_number"]
            else ""
        )

        document_number = document_number.replace(" ", "")

        # --------------------------------------------------------
        # La tabla vieja permitía NULL.
        # La nueva NO.
        #
        # Si no existe documento, usamos el username.
        # Si tampoco existe, usamos LEGACY-ID.
        # --------------------------------------------------------

        if not document_number:
            document_number = username or f"LEGACY-{user_id}"

        original_document_number = document_number
        counter = 1

        while document_number in used_document_numbers:
            document_number = (
                f"{original_document_number}-{counter}"
            )
            counter += 1

        used_document_numbers.add(document_number)

        # --------------------------------------------------------
        # Nombre
        # --------------------------------------------------------

        parts = full_name.split()

        if not parts:
            first_names = username or "Usuario"
            last_names = "GIA"

        elif len(parts) == 1:
            first_names = parts[0]
            last_names = "GIA"

        else:
            first_names = " ".join(parts[:-1])
            last_names = parts[-1]

        # --------------------------------------------------------
        # ROLE LEGACY -> ROLE CANÓNICO
        # --------------------------------------------------------

        old_role = (
            str(user["role"]).strip().lower()
            if user["role"]
            else ""
        )

        role_map = {

            # Aprendiz
            "aprendiz":
                "APPRENTICE",

            "apprentice":
                "APPRENTICE",

            # Instructor
            "docente":
                "FOLLOW_UP_INSTRUCTOR",

            "instructor":
                "FOLLOW_UP_INSTRUCTOR",

            "instructor_seguimiento":
                "FOLLOW_UP_INSTRUCTOR",

            # Instructor líder
            "instructor_lider":
                "LEAD_FOLLOW_UP_INSTRUCTOR",

            "instructor_seguimiento_lider":
                "LEAD_FOLLOW_UP_INSTRUCTOR",

            # Certificador
            "certificador":
                "CERTIFIER",

            # Administrativo
            "administrativo":
                "CENTER_STAFF",

            "administrativo_centro":
                "CENTER_STAFF",

            # Soporte
            "visualizador":
                "SUPPORT",

            "super_admin":
                "SUPPORT",

            "admin":
                "SUPPORT",

            "soporte":
                "SUPPORT",
        }

        role = role_map.get(
            old_role,
            "SUPPORT",
        )

        # --------------------------------------------------------
        # STATUS LEGACY -> STATUS CANÓNICO
        # --------------------------------------------------------

        status = (
            "ACTIVE"
            if user["active"]
            else "INACTIVE"
        )

        # --------------------------------------------------------
        # INSERT
        # --------------------------------------------------------

        conn.execute(
            sa.text(
                """
                INSERT INTO user_new (
                    id,
                    document_type,
                    document_number,
                    first_names,
                    last_names,
                    email,
                    phone,
                    role,
                    status,
                    password_hash,
                    signature_file_name,
                    signature_file_path,
                    signature_updated_at,
                    last_login_at
                )
                VALUES (
                    :id,
                    :document_type,
                    :document_number,
                    :first_names,
                    :last_names,
                    :email,
                    NULL,
                    :role,
                    :status,
                    :password_hash,
                    NULL,
                    NULL,
                    NULL,
                    NULL
                )
                """
            ),
            {
                "id": user_id,
                "document_type": document_type,
                "document_number": document_number,
                "first_names": first_names,
                "last_names": last_names,
                "email": email,
                "role": role,
                "status": status,
                "password_hash": user["password_hash"],
            },
        )

    # ============================================================
    # 5. Eliminar tabla vieja
    # ============================================================

    op.drop_table("user")

    # ============================================================
    # 6. Renombrar tabla nueva
    # ============================================================

    op.rename_table(
        "user_new",
        "user",
    )

    # ============================================================
    # 7. Índices
    # ============================================================

    op.create_index(
        "ix_user_document_type",
        "user",
        ["document_type"],
        unique=False,
    )

    op.create_index(
        "ix_user_document_number",
        "user",
        ["document_number"],
        unique=True,
    )

    op.create_index(
        "ix_user_first_names",
        "user",
        ["first_names"],
        unique=False,
    )

    op.create_index(
        "ix_user_last_names",
        "user",
        ["last_names"],
        unique=False,
    )

    op.create_index(
        "ix_user_email",
        "user",
        ["email"],
        unique=True,
    )

    op.create_index(
        "ix_user_role",
        "user",
        ["role"],
        unique=False,
    )

    op.create_index(
        "ix_user_status",
        "user",
        ["status"],
        unique=False,
    )


def downgrade():
    raise RuntimeError(
        "Downgrade no soportado para esta migración de User."
    )