from datetime import datetime
from extensions import db

class Apprentice(db.Model):
    __tablename__ = "apprentice"

    id = db.Column(db.Integer, primary_key=True)
    created_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    student_user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    group_number = db.Column(db.String(30), nullable=False)
    document_type = db.Column(db.String(30), nullable=False)
    document_number = db.Column(db.String(30), unique=True, nullable=False)
    first_names = db.Column(db.String(120), nullable=False)
    last_names = db.Column(db.String(120), nullable=False)
    gender = db.Column(db.String(20), nullable=True)
    phone = db.Column(db.String(30), nullable=True)
    email = db.Column(db.String(150), nullable=True)
    municipality_origin = db.Column(db.String(120), nullable=True)
    program_name = db.Column(db.String(150), nullable=True)
    program_level = db.Column(db.String(80), nullable=True)
    group_validity = db.Column(db.String(80), nullable=True)
    lead_instructor = db.Column(db.String(150), nullable=True)
    followup_instructor = db.Column(db.String(150), nullable=True)
    followup_instructor_email = db.Column(db.String(150), nullable=True)
    ep_modality = db.Column(db.String(120), nullable=True)
    sofia_status = db.Column(db.String(80), nullable=True)
    practice_start_date = db.Column(db.String(40), nullable=True)
    practice_end_date = db.Column(db.String(40), nullable=True)
    company_name = db.Column(db.String(150), nullable=True)
    company_municipality = db.Column(db.String(120), nullable=True)
    company_address = db.Column(db.String(180), nullable=True)
    coformador_name = db.Column(db.String(150), nullable=True)
    coformador_email = db.Column(db.String(150), nullable=True)
    coformador_phone = db.Column(db.String(30), nullable=True)
    arl_responsible = db.Column(db.String(150), nullable=True)
    individual_management = db.Column(db.Text, nullable=True)
    followup_moments = db.Column(db.String(120), nullable=True)
    evaluation_date = db.Column(db.String(40), nullable=True)
    english_results = db.Column(db.String(120), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Relaciones
    bitacoras = db.relationship(
        "Bitacora",
        backref="apprentice",
        lazy="select",
        cascade="all, delete-orphan",
        order_by="desc(Bitacora.created_at)",
    )

    student_user = db.relationship("User", foreign_keys=[student_user_id], lazy="select")
    evidence_submissions = db.relationship(
        "EvidenceSubmission",
        back_populates="apprentice",
        lazy="select",
        cascade="all, delete-orphan",
    )

    @property
    def full_name(self):
        return f"{self.first_names} {self.last_names}".strip()
