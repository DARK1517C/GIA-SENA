from __future__ import annotations

from datetime import datetime

from sqlalchemy import ForeignKey, String, Text, Index
from sqlalchemy.orm import Mapped, mapped_column, relationship

from extensions import db
from models.base import BaseModel


CERTIFICATION_REVIEW_PENDING = "PENDING"
CERTIFICATION_REVIEW_APPROVED = "APPROVED"
CERTIFICATION_REVIEW_REJECTED = "REJECTED"


class CertificationReview(BaseModel):
    """Auditoría mínima de la decisión del certificador sobre un aprendiz."""

    __tablename__ = "certification_review"
    __table_args__ = (
        Index("ix_certification_review_apprentice", "apprentice_id"),
    )

    apprentice_id: Mapped[int] = mapped_column(
        ForeignKey("apprentice.id"), nullable=False, index=True
    )
    reviewer_id: Mapped[int] = mapped_column(
        ForeignKey("user.id"), nullable=False, index=True
    )
    status: Mapped[str] = mapped_column(
        String(20), nullable=False, default=CERTIFICATION_REVIEW_PENDING,
        server_default=CERTIFICATION_REVIEW_PENDING, index=True,
    )
    notes: Mapped[str | None] = mapped_column(Text, nullable=True)
    reviewed_at: Mapped[datetime | None] = mapped_column(nullable=True)

    apprentice = relationship("Apprentice", backref="certification_reviews")
    reviewer = relationship("User", foreign_keys=[reviewer_id], backref="certification_reviews")
