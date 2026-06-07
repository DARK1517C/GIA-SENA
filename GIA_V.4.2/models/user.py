from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin
from extensions import db

class User(UserMixin, db.Model):
    __tablename__ = "user"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(30), nullable=False, default="docente")
    full_name = db.Column(db.String(150), nullable=False)
    email = db.Column(db.String(150), nullable=True)
    document_type = db.Column(db.String(20), nullable=True)
    document_number = db.Column(db.String(30), nullable=True)
    managed_group_numbers = db.Column(db.Text, nullable=True)
    active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    # Relaciones
    apprentices = db.relationship(
        "Apprentice",
        foreign_keys="Apprentice.created_by",
        backref="owner",
        lazy="select",
    )

    groups = db.relationship(
        "TrainingGroup",
        foreign_keys="TrainingGroup.created_by",
        backref="creator",
        lazy="select",
    )

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
