from collections import OrderedDict, defaultdict
from datetime import datetime

from flask import current_app, url_for

from extensions import db

from models import (
    Apprentice,
    TrainingGroup,
    EvidenceCategory,
    EvidenceTemplate,
    EvidenceActivity,
    EvidenceSubmission,
    EVIDENCE_STATUS_NOT_SUBMITTED,
    EVIDENCE_STATUS_PENDING_REVIEW,
    EVIDENCE_STATUS_REQUIRES_CORRECTION,
    EVIDENCE_STATUS_APPROVED,
)

from services.date_rules import calculate_followup_ranges_from_ep

# =============================================================================
# CONSTANTES
# =============================================================================

FOLLOWUP_MOMENT_RANGES = {
    "Momento 1": (
        "followup_moment1_start",
        "followup_moment1_end",
    ),
    "Momento 2": (
        "followup_moment2_start",
        "followup_moment2_end",
    ),
    "Momento 3": (
        "followup_moment3_start",
        "followup_moment3_end",
    ),
    "Momento 4": (
        "followup_moment4_start",
        "followup_moment4_end",
    ),
}


# =============================================================================
# CATÁLOGO CANÓNICO
# =============================================================================

def get_active_evidence_categories():
    """Devuelve las categorías activas del catálogo canónico.

    Las rutas no deben reconstruir categorías desde constantes Python: la BD
    es la única fuente de verdad del catálogo de evidencias.
    """
    return (
        EvidenceCategory.query
        .filter_by(is_active=True)
        .order_by(EvidenceCategory.sort_order, EvidenceCategory.id)
        .all()
    )


def get_active_evidence_templates():
    """Devuelve las plantillas institucionales activas en orden canónico."""
    return (
        EvidenceTemplate.query
        .filter_by(is_active=True)
        .order_by(
            EvidenceTemplate.category_id,
            EvidenceTemplate.sort_order,
            EvidenceTemplate.id,
        )
        .all()
    )


def get_active_evidence_catalog():
    """Devuelve el catálogo institucional agrupado por categoría.

    No crea categorías artificiales ni mantiene un catálogo paralelo en Python.
    """
    catalog = OrderedDict()
    for category in get_active_evidence_categories():
        catalog[category.code] = {
            "category": category,
            "templates": [],
        }

    for template in get_active_evidence_templates():
        bucket = catalog.get(template.category.code)
        if bucket is not None:
            bucket["templates"].append(template)

    return catalog


# =============================================================================
# CREACIÓN DE EVIDENCIAS PREDETERMINADAS
# =============================================================================

def ensure_template_activities_for_group(group):
    """
    Proyecta las plantillas institucionales activas sobre una ficha.

    La BD es la única fuente de verdad del catálogo:
    EvidenceCategory -> EvidenceTemplate -> EvidenceActivity.
    No existe un catálogo paralelo en Python.
    """
    if not group:
        return []

    templates = get_active_evidence_templates()

    existing_template_ids = {
        activity.template_id
        for activity in EvidenceActivity.query.filter_by(group_id=group.id).all()
        if activity.template_id is not None
    }

    created = []
    for template in templates:
        if template.id in existing_template_ids:
            continue

        activity = EvidenceActivity.from_template(
            template=template,
            group_id=group.id,
            created_by_id=None,
        )
        activity.validate_domain_consistency(template=template)
        db.session.add(activity)
        created.append(activity)

    if created:
        db.session.flush()

    sync_group_followup_dates(group)
    return created


# =============================================================================
# SINCRONIZACIÓN DE FECHAS
# =============================================================================

def sync_group_followup_dates(group):
    """
    Actualiza automáticamente las fechas de vencimiento de las evidencias
    correspondientes a los cuatro momentos de seguimiento.
    """

    if not group:
        return

    # Las fechas de seguimiento usan la misma fuente de verdad del motor de fechas.
    # La fecha final de EP proviene de los aprendices importados; nunca se
    # sustituye por training_end_date (fin de formación).
    apprentice_dates = (
        Apprentice.query
        .filter_by(group_id=group.id)
        .with_entities(Apprentice.practice_start_date, Apprentice.practice_end_date)
        .all()
    )
    if not apprentice_dates:
        return

    starts = [row[0] for row in apprentice_dates if row[0]]
    ends = [row[1] for row in apprentice_dates if row[1]]
    if not starts or not ends:
        return

    unique_starts = {str(value).strip() for value in starts}
    unique_ends = {str(value).strip() for value in ends}
    if len(unique_starts) != 1 or len(unique_ends) != 1:
        # Si las fechas reales de los aprendices difieren, no inventamos una
        # fecha grupal: el detalle individual conserva el calendario real.
        current_app.logger.warning(
            "No se actualizan vencimientos FUP del grupo %s porque sus aprendices "
            "tienen fechas EP inconsistentes.",
            group.group_number,
        )
        return

    ranges = calculate_followup_ranges_from_ep(
        starts[0],
        ends[0],
    )

    activities = (
        EvidenceActivity.query
        .join(EvidenceCategory, EvidenceActivity.category_id == EvidenceCategory.id)
        .filter(
            EvidenceActivity.group_id == group.id,
            EvidenceCategory.code == "followup_moments",
        )
        .all()
    )

    for activity in activities:

        for label, (
            start_key,
            end_key,
        ) in FOLLOWUP_MOMENT_RANGES.items():

            if activity.title.startswith(label):

                activity.due_start = (
                    ranges.get(start_key)
                    or activity.due_start
                )

                activity.due_end = (
                    ranges.get(end_key)
                    or activity.due_end
                )

                break


# =============================================================================
# CREACIÓN DE ENTREGAS
# =============================================================================

def ensure_submissions_for_apprentice(apprentice):
    """
    Garantiza que un aprendiz tenga una entrega asociada para cada
    evidencia existente en su ficha.
    """

    if not apprentice or not apprentice.group_number:
        return []

    group = TrainingGroup.query.filter_by(
        group_number=apprentice.group_number
    ).first()

    if not group:
        return []

    ensure_template_activities_for_group(group)

    activities = (
        EvidenceActivity.query
        .filter_by(group_id=group.id)
        .all()
    )

    existing_activity_ids = {

        submission.activity_id

        for submission in (
            EvidenceSubmission.query
            .filter_by(
                apprentice_id=apprentice.id
            )
            .all()
        )

    }

    created = []

    for activity in activities:

        # La actividad pertenece a la misma ficha del aprendiz. Esta
        # comprobación protege la integridad de dominio que una FK simple
        # (activity_id, apprentice_id) no puede expresar.
        if activity.group_id != group.id:
            continue

        if activity.id in existing_activity_ids:
            continue

        submission = EvidenceSubmission(
            activity_id=activity.id,
            apprentice_id=apprentice.id,
            status=EVIDENCE_STATUS_NOT_SUBMITTED,
        )

        db.session.add(submission)

        created.append(submission)

    if created:
        db.session.flush()

    return created


def ensure_submissions_for_apprentice_group(group):
    """Crea entregas para las actividades ya existentes de una ficha.

    A diferencia de ``ensure_submissions_for_apprentice``, esta variante no
    proyecta plantillas nuevas. Es la operación adecuada después de crear o
    sincronizar una actividad concreta.
    """
    if not group:
        return 0

    activities = (
        EvidenceActivity.query
        .filter_by(group_id=group.id)
        .all()
    )
    if not activities:
        return 0

    apprentices = (
        Apprentice.query
        .filter_by(group_number=group.group_number)
        .all()
    )

    created_count = 0
    for apprentice in apprentices:
        existing_ids = {
            submission.activity_id
            for submission in EvidenceSubmission.query.filter_by(
                apprentice_id=apprentice.id
            ).all()
        }
        for activity in activities:
            if activity.id in existing_ids:
                continue
            db.session.add(
                EvidenceSubmission(
                    activity_id=activity.id,
                    apprentice_id=apprentice.id,
                    status=EVIDENCE_STATUS_NOT_SUBMITTED,
                )
            )
            created_count += 1

    if created_count:
        db.session.flush()

    return created_count


def ensure_group_submissions(group):
    """
    Garantiza que todos los aprendices de una ficha tengan creadas
    sus entregas correspondientes.
    """

    if not group:
        return 0

    ensure_template_activities_for_group(group)

    count = 0

    apprentices = (
        Apprentice.query
        .filter_by(
            group_number=group.group_number
        )
        .all()
    )

    for apprentice in apprentices:

        count += len(
            ensure_submissions_for_apprentice(
                apprentice
            )
        )

    return count


# =============================================================================
# ESTADÍSTICAS
# =============================================================================

def summarize_submissions(submissions):
    """
    Resume las entregas agrupadas por tipo de evidencia.
    """

    summary = OrderedDict()

    for submission in submissions:

        category = getattr(submission.activity, "category", None)
        category_key = getattr(category, "code", None) or "uncategorized"
        category_name = getattr(category, "name", None) or "Sin categoría"

        item = summary.setdefault(
            category_key,
            {
                "total": 0,
                "approved": 0,
                "pending_review": 0,
                "requires_correction": 0,
                "not_submitted": 0,
                "code": category_key,
                "name": category_name,
            },
        )

        item["total"] += 1

        if submission.status == EVIDENCE_STATUS_APPROVED:

            item["approved"] += 1

        elif submission.status == EVIDENCE_STATUS_PENDING_REVIEW:

            item["pending_review"] += 1

        elif submission.status == EVIDENCE_STATUS_REQUIRES_CORRECTION:

            item["requires_correction"] += 1

        else:

            item["not_submitted"] += 1

    return summary


def global_evidence_stats(query=None):
    """
    Calcula las estadísticas generales de las evidencias.
    """

    submissions = (
        query.all()
        if query is not None
        else EvidenceSubmission.query.all()
    )

    total = len(submissions)

    approved = sum(
        1
        for submission in submissions
        if submission.status == EVIDENCE_STATUS_APPROVED
    )

    pending_review = sum(
        1
        for submission in submissions
        if submission.status == EVIDENCE_STATUS_PENDING_REVIEW
    )

    requires_correction = sum(
        1
        for submission in submissions
        if submission.status == EVIDENCE_STATUS_REQUIRES_CORRECTION
    )

    not_submitted = sum(
        1
        for submission in submissions
        if submission.status == EVIDENCE_STATUS_NOT_SUBMITTED
    )

    return {
        "total": total,
        "approved": approved,
        "pending_review": pending_review,
        "requires_correction": requires_correction,
        "not_submitted": not_submitted,
        "delivered": (
            approved
            + pending_review
            + requires_correction
        ),
        "global_percent": (
            round((approved / total) * 100, 1)
            if total
            else 0
        ),
    }


def group_compliance_rows(submissions=None):
    """
    Calcula el porcentaje de cumplimiento por ficha.
    """

    submissions = (
        submissions
        or EvidenceSubmission.query.all()
    )

    grouped = defaultdict(list)

    for submission in submissions:

        grouped[
            submission.apprentice.group_number
        ].append(submission)

    rows = []

    for group_number, items in sorted(grouped.items()):

        approved = sum(
            1
            for item in items
            if item.status == EVIDENCE_STATUS_APPROVED
        )

        total = len(items)

        rows.append({
            "group_number": group_number,
            "approved": approved,
            "total": total,
            "percent": (
                round((approved / total) * 100, 1)
                if total
                else 0
            ),
        })

    return rows


# =============================================================================
# AGRUPACIÓN DE EVIDENCIAS
# =============================================================================

def build_evidence_groups(submissions):
    """Agrupa entregas por EvidenceCategory, usando su código como clave."""
    groups = OrderedDict()

    for submission in submissions:
        activity = submission.activity
        category = getattr(activity, "category", None)
        category_code = getattr(category, "code", None) or "uncategorized"
        category_name = getattr(category, "name", None) or "Sin categoría"

        if category_code not in groups:
            groups[category_code] = {
                "code": category_code,
                "name": category_name,
                "total_files": 0,
                "approved": 0,
                "pending_review": 0,
                "requires_correction": 0,
                "submissions": [],
            }

        info = groups[category_code]
        info["submissions"].append(submission)
        if getattr(submission, "file_name", None):
            info["total_files"] += 1
        if submission.status == EVIDENCE_STATUS_APPROVED:
            info["approved"] += 1
        elif submission.status == EVIDENCE_STATUS_PENDING_REVIEW:
            info["pending_review"] += 1
        elif submission.status == EVIDENCE_STATUS_REQUIRES_CORRECTION:
            info["requires_correction"] += 1

    return groups


def project_template_to_all_groups(template):
    """
    Proyecta una plantilla activa sobre todas las fichas existentes.

    La operación es idempotente y solo crea la actividad correspondiente a
    la plantilla indicada; no modifica actividades ya proyectadas.
    """
    if not template or not template.is_active:
        return 0

    total_created = 0
    groups = TrainingGroup.query.order_by(TrainingGroup.id).all()

    for group in groups:
        exists = EvidenceActivity.query.filter_by(
            group_id=group.id,
            template_id=template.id,
        ).first()
        if exists:
            # La sincronización es idempotente: no toca una actividad ya
            # proyectada ni sus entregas existentes.
            ensure_submissions_for_apprentice_group(group)
            continue

        activity = EvidenceActivity.from_template(
            template=template,
            group_id=group.id,
            created_by_id=None,
        )
        db.session.add(activity)
        db.session.flush()
        ensure_submissions_for_apprentice_group(group)
        total_created += 1

    return total_created
