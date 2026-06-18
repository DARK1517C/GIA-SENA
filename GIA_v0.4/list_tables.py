# list_tables.py
import os
from app import create_app
from extensions import db

app = create_app()
with app.app_context():
    print("Usando DB URI:", app.config.get("SQLALCHEMY_DATABASE_URI"))
    try:
        # SQLAlchemy 1.x
        tables = db.engine.table_names()
    except Exception:
        # SQLAlchemy 1.4+/2.0 alternativa
        tables = [t.name for t in db.metadata.sorted_tables]
    print("Tablas en la DB:", tables)
