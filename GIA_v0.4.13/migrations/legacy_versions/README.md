# Legacy Alembic migrations

These revisions are preserved for historical reference only.

The current GIA database bootstrap uses the canonical clean baseline:

    b7e2c1a4f901
        -> 9c4d2e7a1b60
        -> 3f8a7c2d1e90
        -> 4a6b9c1d2e30

The previous migration chain (`93671fcbae56` / `7c1f2e9a4b10`) was a
separate legacy branch and could not be applied safely to a clean database
alongside the canonical baseline. It is therefore kept outside
`migrations/versions/` so Alembic no longer treats it as an active branch.

Do not restore these files into `migrations/versions/` unless a deliberate
historical migration strategy is designed.
