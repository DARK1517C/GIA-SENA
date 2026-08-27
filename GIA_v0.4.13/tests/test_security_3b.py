from services.permissions import (
    ROLE_APPRENTICE, ROLE_FOLLOWUP_INSTRUCTOR, ROLE_LEAD_FOLLOWUP_INSTRUCTOR,
    ROLE_CENTER_STAFF, ROLE_CERTIFIER, ROLE_SUPPORT, has_permission,
    PERMISSIONS,
)

# Smoke-test estructural de la política 3.B. No requiere BD.
def test_security_3b_matrix():
    expected = {
        ROLE_APPRENTICE: {"evidences.upload"},
        ROLE_FOLLOWUP_INSTRUCTOR: {
            "groups.manage", "apprentices.manage", "evidences.manage",
            "evidences.approve", "evidences.upload", "evidences.sign",
            "evidences.activities.manage",
        },
        ROLE_LEAD_FOLLOWUP_INSTRUCTOR: {
            "groups.manage", "apprentices.manage", "evidences.manage",
            "evidences.approve", "evidences.upload", "evidences.sign",
            "evidences.activities.manage",
        },
        ROLE_CENTER_STAFF: {"data.global_view"},
        ROLE_CERTIFIER: {"data.global_view", "evidences.manage", "evidences.approve"},
        ROLE_SUPPORT: set(PERMISSIONS),
    }
    for role, permissions in expected.items():
        actual = {name for name, roles in PERMISSIONS.items() if role in roles}
        assert actual == permissions
