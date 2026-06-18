from datetime import datetime
from extensions import db

class Bitacora(db.Model):
    __tablename__ = "bitacora"

    id = db.Column(db.Integer, primary_key=True)
    apprentice_id = db.Column(db.Integer, db.ForeignKey("apprentice.id"), nullable=False)
    uploaded_by_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    title = db.Column(db.String(120), nullable=False)
    notes = db.Column(db.Text, nullable=True)
    file_name = db.Column(db.String(255), nullable=True)
    file_path = db.Column(db.String(255), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    uploaded_by = db.relationship("User", lazy="select")
