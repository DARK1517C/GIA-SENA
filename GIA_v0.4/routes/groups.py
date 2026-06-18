# routes/groups.py
from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
from io import BytesIO
from models import TrainingGroup, Apprentice
from extensions import db
from services.excel_import import import_reference_workbook
from services.excel_export import export_reference_workbook, format_value

groups_bp = Blueprint("groups", __name__, url_prefix="/groups")


@groups_bp.route("/")
@login_required
def index():
    groups = TrainingGroup.query.order_by(TrainingGroup.group_number).all()
    # obtener definición de columnas desde el modelo (para export / detalle)
    group_record_fields = getattr(TrainingGroup, "RECORD_FIELDS", [])

    # Campos que deben mostrarse en cada card (solo estos)
    CARD_FIELDS = [
        "group_number",
        "program_name",
        "program_level",
        "municipality",
        "modality",
        "lead_instructor",
        "followup_instructor",
        "group_start_date",
        "training_end_date",
        "ep_start_date",
        "sofia_group_status",
    ]

    # construir una lista de "cards" que contiene el objeto y un dict con los valores necesarios
    cards = []
    for idx, g in enumerate(groups, start=1):
        row = {}
        for key in CARD_FIELDS:
            try:
                value = getattr(g, key)
            except Exception:
                value = None
            # formatear valores (fechas, None) para mostrar en la card
            row[key] = format_value(value)
        # incluir índice si lo necesitas en la card (opcional)
        row["consecutive"] = idx
        cards.append({"group": g, "row": row})

    return render_template(
        "groups/index.html",
        groups=groups,
        cards=cards,
        group_record_fields=group_record_fields
    )


@groups_bp.route("/<int:id>")
@login_required
def detail(id):
    group = TrainingGroup.query.get_or_404(id)
    record_fields = getattr(TrainingGroup, "RECORD_FIELDS", [])

    # Construir un dict 'row' con todos los campos definidos en RECORD_FIELDS,
    # formateando valores (fechas, None, etc.) para que el template muestre exactamente lo exportado.
    row = {}
    for key, _label in record_fields:
        if key == "consecutive":
            # 'consecutive' no aplica en detalle; omitir
            continue
        try:
            value = getattr(group, key)
        except Exception:
            value = None
        row[key] = format_value(value)

    return render_template(
        "groups/detail.html",
        group=group,
        group_record_fields=record_fields,
        row=row
    )


@groups_bp.route("/create", methods=["GET", "POST"])
@login_required
def create():
    if request.method == "POST":
        data = {
            "group_number": request.form.get("group_number", "").strip(),
            "program_name": request.form.get("program_name", "").strip(),
            "lead_instructor": request.form.get("lead_instructor", "").strip(),
            "followup_instructor": request.form.get("followup_instructor", "").strip(),
        }

        if not data["group_number"] or not data["program_name"]:
            flash("Número de ficha y nombre del programa son obligatorios", "warning")
            return render_template("groups/create.html", form=request.form)

        group = TrainingGroup()
        for key, value in data.items():
            if hasattr(group, key):
                setattr(group, key, value)

        if hasattr(group, "created_by"):
            try:
                group.created_by = current_user.id
            except Exception:
                pass

        try:
            db.session.add(group)
            db.session.commit()
            flash("Ficha creada", "success")
            return redirect(url_for("groups.index"))
        except Exception:
            db.session.rollback()
            flash("Ocurrió un error al crear la ficha. Intenta de nuevo.", "danger")
            return render_template("groups/create.html", form=request.form)

    return render_template("groups/create.html")


@groups_bp.route("/<int:id>/edit", methods=["GET", "POST"])
@login_required
def edit(id):
    group = TrainingGroup.query.get_or_404(id)

    if request.method == "POST":
        data = {
            "group_number": request.form.get("group_number", "").strip(),
            "program_name": request.form.get("program_name", "").strip(),
            "lead_instructor": request.form.get("lead_instructor", "").strip(),
            "followup_instructor": request.form.get("followup_instructor", "").strip(),
        }

        if not data["group_number"] or not data["program_name"]:
            flash("Número de ficha y nombre del programa son obligatorios", "warning")
            return render_template("groups/create.html", form=request.form, editing=True, group=group)

        for key, value in data.items():
            if hasattr(group, key):
                setattr(group, key, value)

        if hasattr(group, "updated_by"):
            try:
                group.updated_by = current_user.id
            except Exception:
                pass

        try:
            db.session.commit()
            flash("Ficha actualizada correctamente.", "success")
            return redirect(url_for("groups.index"))
        except Exception:
            db.session.rollback()
            flash("Ocurrió un error al actualizar la ficha. Intenta de nuevo.", "danger")
            return render_template("groups/create.html", form=request.form, editing=True, group=group)

    return render_template("groups/create.html", editing=True, group=group)


@groups_bp.route("/<int:id>/delete", methods=["POST"])
@login_required
def delete(id):
    group = TrainingGroup.query.get_or_404(id)

    try:
        db.session.delete(group)
        db.session.commit()
        flash("Ficha eliminada correctamente.", "success")
    except Exception:
        db.session.rollback()
        flash("No se pudo eliminar la ficha. Asegúrate de que no tenga aprendices asociados.", "danger")

    return redirect(url_for("groups.index"))


@groups_bp.route("/import", methods=["GET", "POST"])
@login_required
def import_groups():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            flash("No se subió ningún archivo", "warning")
            return redirect(url_for("groups.import_groups"))

        filename = secure_filename(file.filename)
        try:
            apprentice_count, group_count, has_apprentice, has_group = import_reference_workbook(
                file, owner_id=current_user.id, mode="groups"
            )
            flash(f"Importados {group_count} grupos y {apprentice_count} aprendices.", "success")
        except Exception:
            flash("Error al procesar el archivo. Verifica el formato y vuelve a intentarlo.", "danger")

        return redirect(url_for("groups.index"))

    return render_template("groups/import.html")


@groups_bp.route("/export")
@login_required
def export_groups():
    try:
        groups = TrainingGroup.query.all()
        apprentices = Apprentice.query.all()
        output = export_reference_workbook(apprentices, groups)

        if isinstance(output, bytes):
            output = BytesIO(output)
            output.seek(0)
        elif hasattr(output, "seek"):
            try:
                output.seek(0)
            except Exception:
                pass
        else:
            try:
                output = BytesIO(output)
                output.seek(0)
            except Exception:
                flash("Error al generar el archivo de exportación. Intenta de nuevo.", "danger")
                return redirect(url_for("groups.index"))

        return send_file(
            output,
            as_attachment=True,
            download_name="grupos_y_aprendices.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception:
        flash("Error al generar el archivo de exportación. Intenta de nuevo.", "danger")
        return redirect(url_for("groups.index"))
