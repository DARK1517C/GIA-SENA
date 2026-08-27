from services.evidence_service import get_active_evidence_categories, get_active_evidence_templates


def test_evidence_admin_catalog_helpers_import():
    assert callable(get_active_evidence_categories)
    assert callable(get_active_evidence_templates)
