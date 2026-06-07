from io import BytesIO
from flask import Blueprint, flash, redirect, render_template, request, send_file, url_for
from flask_login import current_user, login_required
from openpyxl import Workbook

from models import Apprentice, EvidenceActivity, EvidenceSubmission, TrainingGroup, EVIDENCE_STATUS_LABELS, EVIDENCE_TYPES
from services.evidence_service import global_evidence_stats, group_compliance_rows

reports_bp = Blueprint("reports", __name__, url_prefix="/reports")


def _can_view_reports():
    return current_user.role in ["docente", "visualizador", "super_admin"]


def _filtered_query():
    query = EvidenceSubmission.query.join(Apprentice).join(EvidenceActivity)
    filters = {
        "group_number": request.args.get("group_number", "").strip(),
        "program_name": request.args.get("program_name", "").strip(),
        "lead_instructor": request.args.get("lead_instructor", "").strip(),
        "followup_instructor": request.args.get("followup_instructor", "").strip(),
        "municipality": request.args.get("municipality", "").strip(),
        "status": request.args.get("status", "").strip(),
        "date_from": request.args.get("date_from", "").strip(),
        "date_to": request.args.get("date_to", "").strip(),
    }
    if filters["group_number"]:
        query = query.filter(Apprentice.group_number == filters["group_number"])
    if filters["program_name"]:
        query = query.filter(Apprentice.program_name == filters["program_name"])
    if filters["lead_instructor"]:
        query = query.filter(Apprentice.lead_instructor == filters["lead_instructor"])
    if filters["followup_instructor"]:
        query = query.filter(Apprentice.followup_instructor == filters["followup_instructor"])
    if filters["municipality"]:
        query = query.filter(Apprentice.municipality_origin == filters["municipality"])
    if filters["status"]:
        query = query.filter(EvidenceSubmission.status == filters["status"])
    if filters["date_from"]:
        query = query.filter(EvidenceSubmission.uploaded_at >= filters["date_from"])
    if filters["date_to"]:
        query = query.filter(EvidenceSubmission.uploaded_at <= filters["date_to"])
    return query, filters


@reports_bp.route("/")
@login_required
def index():
    if not _can_view_reports():
        flash("No tienes permisos para ver reportes.", "warning")
        return redirect(url_for("dashboard.index"))
    query, filters = _filtered_query()
    submissions = query.all()
    apprentices = Apprentice.query.all()
    groups = TrainingGroup.query.order_by(TrainingGroup.group_number).all()
    stats = global_evidence_stats(query)
    group_rows = group_compliance_rows(submissions)
    filter_options = {
        "groups": sorted({item.group_number for item in apprentices if item.group_number}),
        "programs": sorted({item.program_name for item in apprentices if item.program_name}),
        "leaders": sorted({item.lead_instructor for item in apprentices if item.lead_instructor}),
        "followups": sorted({item.followup_instructor for item in apprentices if item.followup_instructor}),
        "municipalities": sorted({item.municipality_origin for item in apprentices if item.municipality_origin}),
    }
    type_rows = []
    for evidence_type in EVIDENCE_TYPES:
        items = [item for item in submissions if item.activity.evidence_type == evidence_type]
        approved = sum(1 for item in items if item.status == "aprobado")
        type_rows.append({
            "type": evidence_type,
            "approved": approved,
            "total": len(items),
            "percent": round((approved / len(items)) * 100, 1) if items else 0,
        })
    return render_template(
        "reports/index.html",
        filters=filters,
        filter_options=filter_options,
        stats=stats,
        group_rows=group_rows,
        type_rows=type_rows,
        submissions=submissions[:200],
        total_apprentices=len(apprentices),
        total_groups=len(groups),
        status_labels=EVIDENCE_STATUS_LABELS,
    )


def _report_rows(submissions):
    rows = []
    for item in submissions:
        rows.append([
            item.apprentice.group_number,
            item.apprentice.full_name,
            item.apprentice.document_number,
            item.apprentice.program_name or "",
            item.apprentice.lead_instructor or "",
            item.apprentice.followup_instructor or "",
            item.apprentice.municipality_origin or "",
            item.activity.evidence_type,
            item.activity.title,
            EVIDENCE_STATUS_LABELS.get(item.status, item.status),
            item.observations or "",
        ])
    return rows


@reports_bp.route("/export.xlsx")
@login_required
def export_xlsx():
    if not _can_view_reports():
        flash("No tienes permisos para exportar reportes.", "warning")
        return redirect(url_for("dashboard.index"))
    query, _filters = _filtered_query()
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"
    headers = ["Ficha", "Aprendiz", "Documento", "Programa", "Instructor lider", "Instructor seguimiento", "Municipio", "Tipo", "Evidencia", "Estado", "Observaciones"]
    ws.append(headers)
    for row in _report_rows(query.all()):
        ws.append(row)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="reporte_evidencias.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@reports_bp.route("/export.pdf")
@login_required
def export_pdf():
    if not _can_view_reports():
        flash("No tienes permisos para exportar reportes.", "warning")
        return redirect(url_for("dashboard.index"))
    query, _filters = _filtered_query()
    stats = global_evidence_stats(query)
    lines = [
        "Reporte GIA - Evidencias",
        f"Total evidencias: {stats['total']}",
        f"Entregadas: {stats['delivered']}",
        f"No entregadas: {stats['not_submitted']}",
        f"Pendientes de aprobacion: {stats['pending']}",
        f"Aprobadas: {stats['approved']}",
        f"Cumplimiento global: {stats['global_percent']}%",
        "",
    ]
    for row in _report_rows(query.limit(80).all()):
        lines.append(" | ".join(str(value)[:70] for value in row))
    text = "\n".join(lines).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    stream = f"""%PDF-1.4
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >> endobj
4 0 obj << /Length {len(text) + 80} >> stream
BT /F1 9 Tf 40 760 Td 12 TL ({text}) Tj ET
endstream endobj
5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
xref
0 6
0000000000 65535 f 
trailer << /Root 1 0 R /Size 6 >>
startxref
0
%%EOF"""
    output = BytesIO(stream.encode("latin-1", errors="ignore"))
    return send_file(output, as_attachment=True, download_name="reporte_evidencias.pdf", mimetype="application/pdf")
