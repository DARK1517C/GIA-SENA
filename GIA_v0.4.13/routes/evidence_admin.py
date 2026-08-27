"""Administración del catálogo canónico de evidencias.

Fase C.5 del Bloque Arquitectura:
- CRUD de EvidenceCategory.
- CRUD de EvidenceTemplate.
- CRUD de EvidenceActivity.

Las categorías y plantillas son configuración institucional global y solo
Soporte puede modificarlas. Las actividades son operativas por ficha y se
limitan por alcance al instructor responsable, al líder o a Soporte.
"""
from __future__ import annotations

from flask import Blueprint, current_app, flash, redirect, render_template, request, url_for
from flask_login import current_user, login_required
from sqlalchemy.exc import IntegrityError

from extensions import db
from services.auth_helpers import permission_required
from models import EvidenceActivity, EvidenceCategory, EvidenceSubmission, EvidenceTemplate, TrainingGroup
from services.access_scope import can_manage_group, visible_groups_query
from services.evidence_service import (
    ensure_group_submissions,
    get_active_evidence_categories,
    get_active_evidence_templates,
    project_template_to_all_groups,
)
from services.permissions import has_permission


evidence_admin_bp = Blueprint(
    "evidence_admin",
    __name__,
    url_prefix="/evidencias/admin",
)


def _require(permission: str, redirect_endpoint: str = "evidences.index") -> bool:
    if has_permission(permission):
        return True
    flash("No tienes permisos para realizar esta operación.", "warning")
    return False


def _bool(name: str, default: bool = False) -> bool:
    value = request.form.get(name)
    if value is None:
        return default
    return value.lower() in {"1", "true", "on", "yes", "si", "sí"}


def _int(name: str, default: int = 0) -> int:
    raw = (request.form.get(name) or "").strip()
    if not raw:
        return default
    return int(raw)


def _extensions() -> str | None:
    raw = (request.form.get("allowed_extensions") or "").strip()
    return raw or None


@evidence_admin_bp.route("/")
@login_required
def index():
    can_catalog = has_permission("evidences.catalog.manage")
    can_activities = has_permission("evidences.activities.manage")
    if not (can_catalog or can_activities):
        flash("No tienes permisos para administrar evidencias.", "warning")
        return redirect(url_for("evidences.index"))

    categories = get_active_evidence_categories()
    templates = get_active_evidence_templates()

    if can_activities:
        groups = visible_groups_query().order_by(TrainingGroup.group_number).all()
        activities = (
            EvidenceActivity.query
            .filter(EvidenceActivity.group_id.in_([g.id for g in groups]))
            .order_by(EvidenceActivity.group_id, EvidenceActivity.sort_order, EvidenceActivity.id)
            .all()
            if groups else []
        )
    else:
        groups, activities = [], []

    return render_template(
        "evidences/admin/index.html",
        categories=categories,
        templates=templates,
        groups=groups,
        activities=activities,
        can_catalog=can_catalog,
        can_activities=can_activities,
    )


# -----------------------------------------------------------------------------
# CATEGORÍAS
# -----------------------------------------------------------------------------

@evidence_admin_bp.route("/categories/create", methods=["GET", "POST"])
@login_required
@permission_required("evidences.catalog.manage")
def category_create():
    if not _require("evidences.catalog.manage"):
        return redirect(url_for("evidence_admin.index"))
    if request.method == "POST":
        try:
            category = EvidenceCategory(
                code=request.form.get("code"),
                name=request.form.get("name"),
                description=request.form.get("description") or None,
                icon=request.form.get("icon") or None,
                color=request.form.get("color") or None,
                sort_order=_int("sort_order"),
                is_active=_bool("is_active", True),
            )
            db.session.add(category)
            db.session.commit()
            flash("Categoría creada correctamente.", "success")
            return redirect(url_for("evidence_admin.index"))
        except (ValueError, IntegrityError):
            db.session.rollback()
            current_app.logger.exception("Error creando categoría de evidencia")
            flash("No fue posible crear la categoría. Verifica código y nombre únicos.", "danger")
    return render_template("evidences/admin/category_form.html", category=None)


@evidence_admin_bp.route("/categories/<int:category_id>/edit", methods=["GET", "POST"])
@login_required
@permission_required("evidences.catalog.manage")
def category_edit(category_id):
    if not _require("evidences.catalog.manage"):
        return redirect(url_for("evidence_admin.index"))
    category = EvidenceCategory.query.get_or_404(category_id)
    if request.method == "POST":
        try:
            category.code = request.form.get("code")
            category.name = request.form.get("name")
            category.description = request.form.get("description") or None
            category.icon = request.form.get("icon") or None
            category.color = request.form.get("color") or None
            category.sort_order = _int("sort_order")
            category.is_active = _bool("is_active", True)
            db.session.commit()
            flash("Categoría actualizada correctamente.", "success")
            return redirect(url_for("evidence_admin.index"))
        except (ValueError, IntegrityError):
            db.session.rollback()
            current_app.logger.exception("Error actualizando categoría %s", category_id)
            flash("No fue posible actualizar la categoría.", "danger")
    return render_template("evidences/admin/category_form.html", category=category)


@evidence_admin_bp.route("/categories/<int:category_id>/delete", methods=["POST"])
@login_required
@permission_required("evidences.catalog.manage")
def category_delete(category_id):
    if not _require("evidences.catalog.manage"):
        return redirect(url_for("evidence_admin.index"))
    category = EvidenceCategory.query.get_or_404(category_id)
    if category.templates or category.activities:
        flash("La categoría no se puede eliminar porque tiene plantillas o actividades asociadas. Desactívala en su lugar.", "warning")
        return redirect(url_for("evidence_admin.index"))
    try:
        db.session.delete(category)
        db.session.commit()
        flash("Categoría eliminada correctamente.", "success")
    except IntegrityError:
        db.session.rollback()
        flash("No fue posible eliminar la categoría por restricciones de integridad.", "danger")
    return redirect(url_for("evidence_admin.index"))


# -----------------------------------------------------------------------------
# PLANTILLAS
# -----------------------------------------------------------------------------

@evidence_admin_bp.route("/templates/create", methods=["GET", "POST"])
@login_required
@permission_required("evidences.catalog.manage")
def template_create():
    if not _require("evidences.catalog.manage"):
        return redirect(url_for("evidence_admin.index"))
    categories = get_active_evidence_categories()
    if request.method == "POST":
        try:
            template = EvidenceTemplate(
                category_id=_int("category_id"),
                code=request.form.get("code"),
                title=request.form.get("title"),
                description=request.form.get("description") or None,
                allowed_extensions=_extensions(),
                max_file_size_mb=_int("max_file_size_mb", 0) or None,
                requires_signature=_bool("requires_signature"),
                is_required=_bool("is_required", True),
                sort_order=_int("sort_order"),
                is_active=_bool("is_active", True),
                created_by_id=current_user.id,
            )
            db.session.add(template)
            db.session.commit()
            flash("Plantilla creada correctamente.", "success")
            return redirect(url_for("evidence_admin.index"))
        except (ValueError, IntegrityError):
            db.session.rollback()
            current_app.logger.exception("Error creando plantilla de evidencia")
            flash("No fue posible crear la plantilla. Verifica los datos únicos y obligatorios.", "danger")
    return render_template("evidences/admin/template_form.html", template=None, categories=categories)


@evidence_admin_bp.route("/templates/<int:template_id>/edit", methods=["GET", "POST"])
@login_required
@permission_required("evidences.catalog.manage")
def template_edit(template_id):
    if not _require("evidences.catalog.manage"):
        return redirect(url_for("evidence_admin.index"))
    template = EvidenceTemplate.query.get_or_404(template_id)
    categories = get_active_evidence_categories()
    if request.method == "POST":
        new_category_id = _int("category_id")
        if template.activities and new_category_id != template.category_id:
            flash("No puedes cambiar la categoría de una plantilla que ya tiene actividades proyectadas.", "warning")
            return render_template("evidences/admin/template_form.html", template=template, categories=categories)
        try:
            template.category_id = new_category_id
            template.code = request.form.get("code")
            template.title = request.form.get("title")
            template.description = request.form.get("description") or None
            template.allowed_extensions = _extensions()
            template.max_file_size_mb = _int("max_file_size_mb", 0) or None
            template.requires_signature = _bool("requires_signature")
            template.is_required = _bool("is_required", True)
            template.sort_order = _int("sort_order")
            template.is_active = _bool("is_active", True)
            db.session.commit()
            flash("Plantilla actualizada correctamente.", "success")
            return redirect(url_for("evidence_admin.index"))
        except (ValueError, IntegrityError):
            db.session.rollback()
            current_app.logger.exception("Error actualizando plantilla %s", template_id)
            flash("No fue posible actualizar la plantilla.", "danger")
    return render_template("evidences/admin/template_form.html", template=template, categories=categories)


@evidence_admin_bp.route("/templates/<int:template_id>/delete", methods=["POST"])
@login_required
@permission_required("evidences.catalog.manage")
def template_delete(template_id):
    if not _require("evidences.catalog.manage"):
        return redirect(url_for("evidence_admin.index"))
    template = EvidenceTemplate.query.get_or_404(template_id)
    if template.activities:
        flash("La plantilla no se puede eliminar porque tiene actividades proyectadas. Desactívala en su lugar.", "warning")
        return redirect(url_for("evidence_admin.index"))
    try:
        db.session.delete(template)
        db.session.commit()
        flash("Plantilla eliminada correctamente.", "success")
    except IntegrityError:
        db.session.rollback()
        flash("No fue posible eliminar la plantilla por restricciones de integridad.", "danger")
    return redirect(url_for("evidence_admin.index"))


@evidence_admin_bp.route("/templates/<int:template_id>/sync", methods=["POST"])
@login_required
@permission_required("evidences.catalog.manage")
def template_sync(template_id):
    if not _require("evidences.catalog.manage"):
        return redirect(url_for("evidence_admin.index"))
    template = EvidenceTemplate.query.get_or_404(template_id)
    if not template.is_active:
        flash("Solo se pueden sincronizar plantillas activas.", "warning")
        return redirect(url_for("evidence_admin.index"))
    try:
        created = project_template_to_all_groups(template)
        db.session.commit()
        flash(f"Plantilla sincronizada. Actividades nuevas: {created}.", "success")
    except Exception:
        db.session.rollback()
        current_app.logger.exception("Error sincronizando plantilla %s", template_id)
        flash("No fue posible sincronizar la plantilla.", "danger")
    return redirect(url_for("evidence_admin.index"))


# -----------------------------------------------------------------------------
# ACTIVIDADES
# -----------------------------------------------------------------------------

@evidence_admin_bp.route("/activities/create", methods=["GET", "POST"])
@login_required
@permission_required("evidences.activities.manage")
def activity_create():
    if not _require("evidences.activities.manage"):
        return redirect(url_for("evidence_admin.index"))
    groups = visible_groups_query().order_by(TrainingGroup.group_number).all()
    categories = get_active_evidence_categories()
    templates = get_active_evidence_templates()
    if request.method == "POST":
        try:
            group = TrainingGroup.query.get_or_404(_int("group_id"))
            if not can_manage_group(group):
                flash("No puedes gestionar actividades de esa ficha.", "danger")
                return redirect(url_for("evidence_admin.index"))
            origin = (request.form.get("origin") or "custom").strip().lower()
            template_id = _int("template_id", 0) or None
            if origin == "template":
                template = EvidenceTemplate.query.get_or_404(template_id) if template_id else None
                if template is None or not template.is_active:
                    raise ValueError("Selecciona una plantilla activa.")
                category_id = template.category_id
            else:
                template = None
                template_id = None
                category_id = _int("category_id")
            activity = EvidenceActivity(
                group_id=group.id,
                template_id=template_id,
                category_id=category_id,
                code=request.form.get("code") or None,
                title=request.form.get("title"),
                description=request.form.get("description") or None,
                due_start=request.form.get("due_start") or None,
                due_end=request.form.get("due_end") or None,
                allowed_extensions=_extensions(),
                max_file_size_mb=_int("max_file_size_mb", 0) or None,
                requires_signature=_bool("requires_signature"),
                is_required=_bool("is_required", True),
                is_visible=_bool("is_visible", True),
                is_default=_bool("is_default", True),
                origin=origin,
                sort_order=_int("sort_order"),
                created_by_id=current_user.id,
            )
            activity.validate_domain_consistency(template=template)
            db.session.add(activity)
            db.session.flush()
            ensure_group_submissions(group)
            db.session.commit()
            flash("Actividad creada correctamente.", "success")
            return redirect(url_for("evidence_admin.index"))
        except (ValueError, IntegrityError):
            db.session.rollback()
            current_app.logger.exception("Error creando actividad de evidencia")
            flash("No fue posible crear la actividad. Verifica la ficha, categoría, plantilla y datos obligatorios.", "danger")
    return render_template("evidences/admin/activity_form.html", activity=None, groups=groups, categories=categories, templates=templates)


@evidence_admin_bp.route("/activities/<int:activity_id>/edit", methods=["GET", "POST"])
@login_required
@permission_required("evidences.activities.manage")
def activity_edit(activity_id):
    if not _require("evidences.activities.manage"):
        return redirect(url_for("evidence_admin.index"))
    activity = EvidenceActivity.query.get_or_404(activity_id)
    if not can_manage_group(activity.group):
        flash("No puedes editar esta actividad.", "warning")
        return redirect(url_for("evidence_admin.index"))
    groups = visible_groups_query().order_by(TrainingGroup.group_number).all()
    categories = get_active_evidence_categories()
    templates = get_active_evidence_templates()
    if request.method == "POST":
        try:
            new_group = TrainingGroup.query.get_or_404(_int("group_id"))
            if not can_manage_group(new_group):
                raise ValueError("No puedes gestionar actividades de esa ficha.")

            if new_group.id != activity.group_id:
                existing_submissions = EvidenceSubmission.query.filter_by(
                    activity_id=activity.id
                ).count()
                if existing_submissions:
                    flash(
                        "No puedes mover una actividad que ya tiene entregas. "
                        "Crea una actividad nueva en la ficha destino para conservar "
                        "el historial y la trazabilidad.",
                        "warning",
                    )
                    return render_template(
                        "evidences/admin/activity_form.html",
                        activity=activity,
                        groups=groups,
                        categories=categories,
                        templates=templates,
                    )

            origin = (request.form.get("origin") or activity.origin).strip().lower()
            template_id = _int("template_id", 0) or None
            template = None
            if origin == "template":
                template = EvidenceTemplate.query.get_or_404(template_id) if template_id else None
                if template is None or not template.is_active:
                    raise ValueError("Selecciona una plantilla activa.")
                category_id = template.category_id
            else:
                template_id = None
                category_id = _int("category_id")
            activity.group_id = new_group.id
            activity.template_id = template_id
            activity.category_id = category_id
            activity.code = request.form.get("code") or None
            activity.title = request.form.get("title")
            activity.description = request.form.get("description") or None
            activity.due_start = request.form.get("due_start") or None
            activity.due_end = request.form.get("due_end") or None
            activity.allowed_extensions = _extensions()
            activity.max_file_size_mb = _int("max_file_size_mb", 0) or None
            activity.requires_signature = _bool("requires_signature")
            activity.is_required = _bool("is_required", True)
            activity.is_visible = _bool("is_visible", True)
            activity.is_default = _bool("is_default", True)
            activity.origin = origin
            activity.sort_order = _int("sort_order")
            activity.validate_domain_consistency(template=template)
            ensure_group_submissions(new_group)
            db.session.commit()
            flash("Actividad actualizada correctamente.", "success")
            return redirect(url_for("evidence_admin.index"))
        except (ValueError, IntegrityError):
            db.session.rollback()
            current_app.logger.exception("Error actualizando actividad %s", activity_id)
            flash("No fue posible actualizar la actividad.", "danger")
    return render_template("evidences/admin/activity_form.html", activity=activity, groups=groups, categories=categories, templates=templates)


@evidence_admin_bp.route("/activities/<int:activity_id>/delete", methods=["POST"])
@login_required
@permission_required("evidences.activities.manage")
def activity_delete(activity_id):
    if not _require("evidences.activities.manage"):
        return redirect(url_for("evidence_admin.index"))
    activity = EvidenceActivity.query.get_or_404(activity_id)
    if not can_manage_group(activity.group):
        flash("No puedes eliminar esta actividad.", "warning")
        return redirect(url_for("evidence_admin.index"))
    submissions = EvidenceSubmission.query.filter_by(activity_id=activity.id).count()
    if submissions:
        flash("La actividad no se puede eliminar porque ya tiene entregas. Ocúltala o desactiva su proyección para conservar el historial.", "warning")
        return redirect(url_for("evidence_admin.index"))
    try:
        db.session.delete(activity)
        db.session.commit()
        flash("Actividad eliminada correctamente.", "success")
    except IntegrityError:
        db.session.rollback()
        flash("No fue posible eliminar la actividad por restricciones de integridad.", "danger")
    return redirect(url_for("evidence_admin.index"))
