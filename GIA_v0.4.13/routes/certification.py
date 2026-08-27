from __future__ import annotations

from flask import Blueprint, flash, redirect, render_template, request, url_for
from flask_login import current_user, login_required
from sqlalchemy import or_

from extensions import db
from models import Apprentice
from services.auth_helpers import permission_required
from services.permissions import ROLE_CERTIFIER, ROLE_SUPPORT
from services.access_scope import visible_apprentices_query
from services.certification_service import build_certification_checklist, approve_certification, reject_certification
from services.notification_service import (
    notify_apprentice_certification_approved,
    notify_apprentice_certification_rejected,
)

certification_bp = Blueprint("certification", __name__, url_prefix="/certificacion")


def _global_allowed():
    return current_user.role in {ROLE_CERTIFIER, ROLE_SUPPORT}


@certification_bp.get("/")
@login_required
@permission_required("data.global_view")
def index():
    if not _global_allowed():
        flash("No tiene permisos para consultar certificaciones.", "danger")
        return redirect(url_for("dashboard.index"))

    apprentices = (
        visible_apprentices_query()
        .order_by(Apprentice.last_names.asc(), Apprentice.first_names.asc())
        .all()
    )

    rows = []
    for apprentice in apprentices:
        checklist = build_certification_checklist(apprentice)
        rows.append({
            "apprentice": apprentice,
            "ready": checklist["ready"],
            "certified": checklist["is_certified"],
            "requirements_ok": checklist["requirements_ok"],
            "followup_ok": checklist["followup_ok"],
        })

    return render_template("certification/index.html", rows=rows)


@certification_bp.get("/<int:apprentice_id>")
@login_required
@permission_required("data.global_view")
def detail(apprentice_id: int):
    if not _global_allowed():
        flash("No tiene permisos para consultar certificaciones.", "danger")
        return redirect(url_for("dashboard.index"))

    apprentice = Apprentice.query.get_or_404(apprentice_id)
    checklist = build_certification_checklist(apprentice)
    return render_template("certification/detail.html", apprentice=apprentice, **checklist)


@certification_bp.post("/<int:apprentice_id>/approve")
@login_required
@permission_required("data.global_view")
def approve(apprentice_id: int):
    if current_user.role not in {ROLE_CERTIFIER, ROLE_SUPPORT}:
        flash("Solo un certificador autorizado puede aprobar la certificación.", "danger")
        return redirect(url_for("certification.detail", apprentice_id=apprentice_id))

    apprentice = Apprentice.query.get_or_404(apprentice_id)
    try:
        approve_certification(apprentice, current_user, request.form.get("notes"))
        db.session.commit()
        notify_apprentice_certification_approved(apprentice)
        db.session.commit()
        flash("Certificación aprobada correctamente.", "success")
    except ValueError as exc:
        db.session.rollback()
        flash(str(exc), "warning")
    except Exception:
        db.session.rollback()
        flash("No fue posible registrar la aprobación de certificación.", "danger")

    return redirect(url_for("certification.detail", apprentice_id=apprentice_id))


@certification_bp.post("/<int:apprentice_id>/reject")
@login_required
@permission_required("data.global_view")
def reject(apprentice_id: int):
    if current_user.role not in {ROLE_CERTIFIER, ROLE_SUPPORT}:
        flash("Solo un certificador autorizado puede registrar la revisión.", "danger")
        return redirect(url_for("certification.detail", apprentice_id=apprentice_id))

    apprentice = Apprentice.query.get_or_404(apprentice_id)
    try:
        reject_certification(apprentice, current_user, request.form.get("notes", ""))
        db.session.commit()
        notify_apprentice_certification_rejected(apprentice)
        db.session.commit()
        flash("La revisión de certificación quedó registrada como no aprobada.", "success")
    except ValueError as exc:
        db.session.rollback()
        flash(str(exc), "warning")
    except Exception:
        db.session.rollback()
        flash("No fue posible registrar la revisión de certificación.", "danger")

    return redirect(url_for("certification.detail", apprentice_id=apprentice_id))
