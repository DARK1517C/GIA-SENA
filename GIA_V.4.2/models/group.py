from extensions import db

class TrainingGroup(db.Model):
    __tablename__ = "training_group"

    id = db.Column(db.Integer, primary_key=True)
    created_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    group_number = db.Column(db.String(30), unique=True, nullable=False)
    program_name = db.Column(db.String(150), nullable=False)
    lead_instructor = db.Column(db.String(150), nullable=True)
    followup_instructor = db.Column(db.String(150), nullable=True)
    municipality = db.Column(db.String(120), nullable=True)
    program_level = db.Column(db.String(80), nullable=True)
    modality = db.Column(db.String(80), nullable=True)
    sofia_group_status = db.Column(db.String(80), nullable=True)
    group_validity = db.Column(db.String(80), nullable=True)
    group_start_date = db.Column(db.String(40), nullable=True)
    training_end_date = db.Column(db.String(40), nullable=True)
    ep_start_date = db.Column(db.String(40), nullable=True)
    apprentices_statistics = db.Column(db.String(120), nullable=True)
    apprentices_training = db.Column(db.String(30), nullable=True)
    apprentices_enabled = db.Column(db.String(30), nullable=True)
    apprentices_rap_pending = db.Column(db.String(30), nullable=True)
    apprentices_practice = db.Column(db.String(30), nullable=True)
    apprentices_without_alternative = db.Column(db.String(30), nullable=True)
    apprentices_certified = db.Column(db.String(30), nullable=True)
    productive_modalities = db.Column(db.String(120), nullable=True)
    learning_contract = db.Column(db.String(30), nullable=True)
    internship = db.Column(db.String(30), nullable=True)
    productive_project = db.Column(db.String(30), nullable=True)
    employment_link = db.Column(db.String(30), nullable=True)
    evidence_activities = db.relationship(
        "EvidenceActivity",
        back_populates="group",
        lazy="select",
        cascade="all, delete-orphan",
    )

    # Campos que se usan en el "Record de fichas" (clave, etiqueta)
    RECORD_FIELDS = [
        ('consecutive', 'CONSECUTIVO'),
        ('group_number', 'N° DE FICHA'),
        ('lead_instructor', 'NOMBRE DE INSTRUCTOR(A) LÍDER DE LA FICHA'),
        ('followup_instructor', 'NOMBRE DE INSTRUCTOR(A) DE SEGUIMIENTO ETAPA PRODUCTIVA (EP)'),
        ('program_name', 'NOMBRE DEL PROGRAMA DE FORMACIÓN'),
        ('municipality', 'MUNICIPIO'),
        ('program_level', 'NIVEL DE PROGRAMA'),
        ('modality', 'MODALIDAD'),
        ('sofia_group_status', 'ESTADO DE LA FICHA EN SOFÍAPLUS'),
        ('group_start_date', 'FECHA INICIO DE LA FICHA EN SOFIAPLUS'),
        ('training_end_date', 'FECHA FIN DE LA FORMACIÓN EN SOFIAPLUS'),
        ('ep_start_date', 'FECHA INICIO DE ETAPA PRODUCTIVA'),
        ('group_validity', 'VIGENCIA DE LA FICHA'),
        ('apprentices_training', 'APRENDICES EN FORMACIÓN'),
        ('apprentices_enabled', 'APRENDICES HABILITADOS PARA INICIAR ETAPA PRODUCTIVA'),
        ('apprentices_rap_pending', 'APRENDICES QUE DEBEN RAP'),
        ('apprentices_practice', 'APRENDICES EN PRÁCTICA'),
        ('learning_contract', 'CONTRATO DE APRENDIZAJE'),
        ('internship', 'PASANTIA'),
        ('productive_project', 'PROYECTO PRODUCTIVO'),
        ('apprentices_without_alternative', 'APRENDICES SIN ALTERNATIVA DE PRÁCTIVA'),
        ('apprentices_certified', 'APRENDICES CERTIFICADOS'),
    ]
