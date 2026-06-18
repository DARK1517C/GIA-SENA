from collections import defaultdict
from datetime import datetime
from services.utils import calculate_followup_ranges
from models import (
    Apprentice,
    TrainingGroup,
    EvidenceActivity,
    EvidenceSubmission,
    DEFAULT_EVIDENCES,
    EVIDENCE_STATUS_APPROVED,
    EVIDENCE_STATUS_NOT_SUBMITTED,
    EVIDENCE_TYPES,
)
from extensions import db


def seed_default_evidences_for_group(group):
    if not group:
        return []

    existing_titles = {
        item.title
        for item in EvidenceActivity.query.filter_by(group_id=group.id).all()
    }
    created = []
    for evidence_type, title in DEFAULT_EVIDENCES:
        if title in existing_titles:
            continue
        activity = EvidenceActivity(
            group_id=group.id,
            evidence_type=evidence_type,
            title=title,
            description=f"Evidencia predefinida para {evidence_type}.",
            is_default=True,
        )
        db.session.add(activity)
        created.append(activity)
    if created:
        db.session.flush()
    sync_group_followup_dates(group)
    return created


def sync_group_followup_dates(group):
    """Populate default group-level due ranges for the four follow-up moments."""
    if not group:
        return
    ranges = calculate_followup_ranges(group.ep_start_date, group.training_end_date)
    moment_ranges = {
        "Momento 1": ("followup_moment1_start", "followup_moment1_end"),
        "Momento 2": ("followup_moment2_start", "followup_moment2_end"),
        "Momento 3": ("followup_moment3_start", "followup_moment3_end"),
        "Momento 4": ("followup_moment4_start", "followup_moment4_end"),
    }
    activities = EvidenceActivity.query.filter_by(group_id=group.id, evidence_type="Momentos de Seguimiento").all()
    for activity in activities:
        for label, (start_key, end_key) in moment_ranges.items():
            if activity.title.startswith(label):
                activity.due_start = ranges.get(start_key) or activity.due_start
                activity.due_end = ranges.get(end_key) or activity.due_end
                break


def ensure_submissions_for_apprentice(apprentice):
    if not apprentice or not apprentice.group_number:
        return []
    group = TrainingGroup.query.filter_by(group_number=apprentice.group_number).first()
    if not group:
        return []

    seed_default_evidences_for_group(group)
    activities = EvidenceActivity.query.filter_by(group_id=group.id).all()
    existing_activity_ids = {
        item.activity_id
        for item in EvidenceSubmission.query.filter_by(apprentice_id=apprentice.id).all()
    }
    created = []
    for activity in activities:
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


def ensure_group_submissions(group):
    if not group:
        return 0
    seed_default_evidences_for_group(group)
    count = 0
    apprentices = Apprentice.query.filter_by(group_number=group.group_number).all()
    for apprentice in apprentices:
        count += len(ensure_submissions_for_apprentice(apprentice))
    return count


def summarize_submissions(submissions):
    summary = {
        evidence_type: {"total": 0, "approved": 0, "pending": 0, "not_submitted": 0}
        for evidence_type in EVIDENCE_TYPES
    }
    for submission in submissions:
        evidence_type = submission.activity.evidence_type
        item = summary.setdefault(evidence_type, {"total": 0, "approved": 0, "pending": 0, "not_submitted": 0})
        item["total"] += 1
        if submission.status == EVIDENCE_STATUS_APPROVED:
            item["approved"] += 1
        elif submission.status == "pendiente_aprobacion":
            item["pending"] += 1
        else:
            item["not_submitted"] += 1
    return summary


def global_evidence_stats(query=None):
    submissions = query.all() if query is not None else EvidenceSubmission.query.all()
    total = len(submissions)
    approved = sum(1 for item in submissions if item.status == EVIDENCE_STATUS_APPROVED)
    pending = sum(1 for item in submissions if item.status == "pendiente_aprobacion")
    not_submitted = sum(1 for item in submissions if item.status == EVIDENCE_STATUS_NOT_SUBMITTED)
    return {
        "total": total,
        "approved": approved,
        "pending": pending,
        "not_submitted": not_submitted,
        "delivered": approved + pending,
        "global_percent": round((approved / total) * 100, 1) if total else 0,
    }


def group_compliance_rows(submissions=None):
    submissions = submissions or EvidenceSubmission.query.all()
    grouped = defaultdict(list)
    for submission in submissions:
        grouped[submission.apprentice.group_number].append(submission)
    rows = []
    for group_number, items in sorted(grouped.items()):
        approved = sum(1 for item in items if item.status == EVIDENCE_STATUS_APPROVED)
        total = len(items)
        rows.append({
            "group_number": group_number,
            "approved": approved,
            "total": total,
            "percent": round((approved / total) * 100, 1) if total else 0,
        })
    return rows
