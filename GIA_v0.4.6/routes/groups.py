# routes/groups.py
from flask import Blueprint, render_template, request, redirect, url_for, flash, send_file, current_app
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename
from io import BytesIO
from sqlalchemy import or_
from models import TrainingGroup, Apprentice
from extensions import db
from services.excel_import import import_reference_workbook
from services.excel_export import export_reference_workbook, format_value
from services.evidence_service import ensure_group_submissions, seed_default_evidences_for_group

groups_bp = Blueprint("groups", __name__, url_prefix="/groups")

GROUP_FORM_FIELDS = [
    "group_number",
    "program_name",
    "program_level",
    "modality",
    "municipality",
    "group_validity",
    "sofia_group_status",
    "group_start_date",
    "training_end_date",
    "ep_start_date",
    "lead_instructor",
    "followup_instructor",
    "apprentices_training",
    "apprentices_enabled",
    "apprentices_practice",
    "apprentices_certified",
]


def _group_form_data():
    return {field: request.form.get(field, "").strip() for field in GROUP_FORM_FIELDS}


@groups_bp.route("/")
@login_required
def index():
    """
    Lista de grupos (index).
    - Restaura la construcción de `cards` tal como estaba antes.
    - Mantiene además las listas únicas para selects de filtro:
      municipalities, program_levels, sofia_statuses.
    - Devuelve groups (lista completa) y cards (lista con row formateado).
    """
    # Aplicar filtros y obtener todos los grupos (igual que antes)
    groups = _filtered_groups_query().order_by(TrainingGroup.group_number).all()

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

    # total para toolbar (mantener compatibilidad)
    try:
        total_groups = TrainingGroup.query.count()
    except Exception:
        total_groups = len(cards)

    # Obtener opciones predefinidas para selects (valores únicos existentes en la DB)
    # Estas listas se usan en el formulario de filtros (municipality, program_level, sofia_group_status)
    try:
        municipalities_q = (
            db.session.query(TrainingGroup.municipality)
            .filter(TrainingGroup.municipality.isnot(None))
            .distinct()
            .order_by(TrainingGroup.municipality)
            .all()
        )
        municipalities = [m[0] for m in municipalities_q if m[0]]
    except Exception:
        current_app.logger.debug("Error fetching municipalities", exc_info=True)
        municipalities = []

    try:
        program_levels_q = (
            db.session.query(TrainingGroup.program_level)
            .filter(TrainingGroup.program_level.isnot(None))
            .distinct()
            .order_by(TrainingGroup.program_level)
            .all()
        )
        program_levels = [p[0] for p in program_levels_q if p[0]]
    except Exception:
        current_app.logger.debug("Error fetching program_levels", exc_info=True)
        program_levels = []

    try:
        sofia_statuses_q = (
            db.session.query(TrainingGroup.sofia_group_status)
            .filter(TrainingGroup.sofia_group_status.isnot(None))
            .distinct()
            .order_by(TrainingGroup.sofia_group_status)
            .all()
        )
        sofia_statuses = [s[0] for s in sofia_statuses_q if s[0]]
    except Exception:
        current_app.logger.debug("Error fetching sofia statuses", exc_info=True)
        sofia_statuses = []

    return render_template(
        "groups/index.html",
        groups=groups,
        cards=cards,
        group_record_fields=group_record_fields,
        total_groups=total_groups,
        municipalities=municipalities,
        program_levels=program_levels,
        sofia_statuses=sofia_statuses,
    )


def _filtered_groups_query():
    query = TrainingGroup.query

    search = (request.args.get("search") or "").strip()
    municipality = (request.args.get("municipality") or "").strip()
    program_level = (request.args.get("program_level") or "").strip()
    sofia_group_status = (request.args.get("sofia_group_status") or "").strip()

    # Búsqueda general
    if search:
        if len(search) > 300:
            search = search[:300]
        pattern = f"%{search}%"
        query = query.filter(or_(
            TrainingGroup.group_number.ilike(pattern),
            TrainingGroup.program_name.ilike(pattern),
            TrainingGroup.lead_instructor.ilike(pattern),
            TrainingGroup.followup_instructor.ilike(pattern),
        ))

    # Municipio: como viene de select, preferible igualdad exacta (pero tolerante con ilike si lo deseas)
    if municipality:
        query = query.filter(TrainingGroup.municipality == municipality)

    # Nivel del programa (program_level)
    if program_level:
        query = query.filter(TrainingGroup.program_level == program_level)

    # Estado (sofia_group_status)
    if sofia_group_status:
        query = query.filter(TrainingGroup.sofia_group_status == sofia_group_status)

    return query


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

    ensure_group_submissions(group)
    db.session.commit()
    apprentices = Apprentice.query.filter_by(group_id=group.id).order_by(Apprentice.last_names, Apprentice.first_names).all()

    return render_template(
        "groups/detail.html",
        group=group,
        group_record_fields=record_fields,
        row=row,
        apprentices=apprentices,
    )


@groups_bp.route("/create", methods=["GET", "POST"])
@login_required
def create():
    if request.method == "POST":
        data = _group_form_data()

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
            db.session.flush()
            seed_default_evidences_for_group(group)
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
        data = _group_form_data()

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
    if getattr(current_user, "role", None) not in ["docente", "super_admin"]:
        flash("No tienes permisos para eliminar fichas.", "warning")
        return redirect(url_for("groups.index"))
    group = TrainingGroup.query.get_or_404(id)
    group_number = group.group_number
    associated = Apprentice.query.filter_by(group_id=group.id).count()

    try:
        Apprentice.query.filter_by(group_id=group.id).update({"group_id": None}, synchronize_session=False)
        db.session.delete(group)
        db.session.commit()
        current_app.logger.info("Ficha eliminada: id=%s numero=%s aprendices_desasociados=%s usuario=%s", id, group_number, associated, current_user.id)
        flash(f"Ficha eliminada correctamente. Aprendices desasociados: {associated}.", "success")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("No se pudo eliminar la ficha %s", id)
        flash("No se pudo eliminar la ficha.", "danger")

    return redirect(url_for("groups.index"))


@groups_bp.route("/bulk-delete", methods=["POST"])
@login_required
def bulk_delete():
    if getattr(current_user, "role", None) != "super_admin":
        flash("No tienes permisos para eliminar fichas en lote.", "warning")
        return redirect(url_for("groups.index"))

    try:
        ids = [int(item) for item in request.form.getlist("selected_ids") if str(item).strip()]
    except Exception:
        flash("Seleccion invalida.", "warning")
        return redirect(url_for("groups.index"))

    if not ids:
        flash("Selecciona al menos una ficha para eliminar.", "warning")
        return redirect(url_for("groups.index"))

    try:
        items = TrainingGroup.query.filter(TrainingGroup.id.in_(ids)).all()
        deleted = len(items)
        affected_apprentices = 0
        for item in items:
            affected_apprentices += Apprentice.query.filter_by(group_id=item.id).count()
            Apprentice.query.filter_by(group_id=item.id).update({"group_id": None}, synchronize_session=False)
            db.session.delete(item)
        db.session.commit()
        current_app.logger.info("Fichas eliminadas en lote: total=%s aprendices_desasociados=%s usuario=%s", deleted, affected_apprentices, current_user.id)
        flash(f"Eliminadas {deleted} fichas. Aprendices desasociados: {affected_apprentices}.", "success")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("No se pudieron eliminar fichas en lote")
        flash("No se pudieron eliminar las fichas seleccionadas.", "danger")
    return redirect(url_for("groups.index"))


@groups_bp.route("/delete-all", methods=["POST"])
@login_required
def delete_all():
    if getattr(current_user, "role", None) != "super_admin":
        flash("No tienes permisos para eliminar todas las fichas.", "warning")
        return redirect(url_for("groups.index"))
    try:
        total = TrainingGroup.query.count()
        affected_apprentices = Apprentice.query.filter(Apprentice.group_id.isnot(None)).count()
        Apprentice.query.filter(Apprentice.group_id.isnot(None)).update({"group_id": None}, synchronize_session=False)
        for item in TrainingGroup.query.all():
            db.session.delete(item)
        db.session.commit()
        current_app.logger.info("Todas las fichas eliminadas: total=%s aprendices_desasociados=%s usuario=%s", total, affected_apprentices, current_user.id)
        flash(f"Eliminadas todas las fichas ({total}). Aprendices desasociados: {affected_apprentices}.", "success")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("No se pudieron eliminar todas las fichas")
        flash("No se pudieron eliminar todas las fichas.", "danger")
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
            result = import_reference_workbook(file, owner_id=current_user.id, mode="both")
            if not result.has_apprentice_sheet and not result.has_group_sheet:
                flash("El archivo no contiene las hojas oficiales Record Fichas y Aprendices.", "warning")
            else:
                flash(
                    f"Importacion completada: {result.group_count} fichas y {result.apprentice_count} aprendices. "
                    f"Omitidos: {result.skipped_apprentices}.",
                    "success" if not result.errors else "warning",
                )
                for message in result.errors[:5]:
                    flash(message, "warning")
        except Exception:
            flash("Error al procesar el archivo. Verifica el formato y vuelve a intentarlo.", "danger")

        return redirect(url_for("groups.index"))

    return render_template("groups/import.html")


@groups_bp.route("/export")
@login_required
def export_groups():
    try:
        groups = _filtered_groups_query().order_by(TrainingGroup.group_number).all()
        group_numbers = {item.group_number for item in groups if item.group_number}
        apprentices = Apprentice.query.filter(Apprentice.group_id.in_([item.id for item in groups])).all() if groups else []
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
