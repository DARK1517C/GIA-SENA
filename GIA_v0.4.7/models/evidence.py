from datetime import datetime
from extensions import db


EVIDENCE_STATUS_NOT_SUBMITTED = "no_entregado"
EVIDENCE_STATUS_PENDING = "pendiente_aprobacion"
EVIDENCE_STATUS_APPROVED = "aprobado"

EVIDENCE_STATUS_LABELS = {
    EVIDENCE_STATUS_NOT_SUBMITTED: "No entregado",
    EVIDENCE_STATUS_PENDING: "Pendiente de aprobación",
    EVIDENCE_STATUS_APPROVED: "Aprobado",
}

EVIDENCE_STATUS_COLORS = {
    EVIDENCE_STATUS_NOT_SUBMITTED: "#d13438",
    EVIDENCE_STATUS_PENDING: "#fdc300",
    EVIDENCE_STATUS_APPROVED: "#39a900",
}

EVIDENCE_TYPES = [
    "Requisitos Iniciales",
    "Bitacoras",
    "Momentos de Seguimiento",
    "Requisitos de Certificacion",
]

DEFAULT_EVIDENCES = [
    ("Requisitos Iniciales", "F-023 Formato de Planeacion, Seguimiento y Evaluacion de Etapa Productiva"),
    ("Requisitos Iniciales", "F-165 Formato seleccion modificacion alternativa etapa productiva (individual)"),
    ("Requisitos Iniciales", "Certificado de afiliacion ARL"),
    ("Bitacoras", "Bitacora 1"),
    ("Bitacoras", "Bitacora 2"),
    ("Bitacoras", "Bitacora 3"),
    ("Bitacoras", "Bitacora 4"),
    ("Bitacoras", "Bitacora 5"),
    ("Bitacoras", "Bitacora 6"),
    ("Momentos de Seguimiento", "Momento 1: planeacion de la etapa productiva"),
    ("Momentos de Seguimiento", "Momento 2: seguimiento de la etapa productiva"),
    ("Momentos de Seguimiento", "Momento 3: evaluacion de la etapa productiva"),
    ("Momentos de Seguimiento", "Momento 4: adicional (opcional)"),
    ("Requisitos de Certificacion", "Copia de documento de Identidad"),
    ("Requisitos de Certificacion", "Certificado de presentacion Pruebas ICFES TyT"),
    ("Requisitos de Certificacion", "Certificado de la APE"),
    ("Requisitos de Certificacion", "Carnet Destruido"),
    ("Requisitos de Certificacion", "Certificado de ente coformador aprobando finalizacion de practicas"),
]


class EvidenceActivity(db.Model):
    __tablename__ = "evidence_activity"

    id = db.Column(db.Integer, primary_key=True)
    group_id = db.Column(db.Integer, db.ForeignKey("training_group.id"), nullable=False)
    evidence_type = db.Column(db.String(80), nullable=False)
    title = db.Column(db.String(180), nullable=False)
    description = db.Column(db.Text, nullable=True)
    due_start = db.Column(db.String(40), nullable=True)
    due_end = db.Column(db.String(40), nullable=True)
    is_default = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    group = db.relationship("TrainingGroup", back_populates="evidence_activities")
    submissions = db.relationship(
        "EvidenceSubmission",
        back_populates="activity",
        lazy="select",
        cascade="all, delete-orphan",
    )


class EvidenceSubmission(db.Model):
    __tablename__ = "evidence_submission"

    id = db.Column(db.Integer, primary_key=True)
    activity_id = db.Column(db.Integer, db.ForeignKey("evidence_activity.id"), nullable=False)
    apprentice_id = db.Column(db.Integer, db.ForeignKey("apprentice.id"), nullable=False)
    status = db.Column(db.String(40), nullable=False, default=EVIDENCE_STATUS_NOT_SUBMITTED)
    observations = db.Column(db.Text, nullable=True)
    file_name = db.Column(db.String(255), nullable=True)
    file_path = db.Column(db.String(255), nullable=True)
    signature_file_name = db.Column(db.String(255), nullable=True)
    signature_file_path = db.Column(db.String(255), nullable=True)
    uploaded_at = db.Column(db.DateTime, nullable=True)
    approved_at = db.Column(db.DateTime, nullable=True)
    approved_by_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    activity = db.relationship("EvidenceActivity", back_populates="submissions")
    apprentice = db.relationship("Apprentice", back_populates="evidence_submissions")
    approved_by = db.relationship("User", foreign_keys=[approved_by_id], lazy="select")

    @property
    def status_label(self):
        return EVIDENCE_STATUS_LABELS.get(self.status, self.status)

    @property
    def status_color(self):
        return EVIDENCE_STATUS_COLORS.get(self.status, "#737373")


class InstructorSignature(db.Model):
    __tablename__ = "instructor_signature"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), unique=True, nullable=False)
    file_name = db.Column(db.String(255), nullable=False)
    file_path = db.Column(db.String(255), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    user = db.relationship("User", backref=db.backref("signature", uselist=False), lazy="select")
