# seed_admin.py
from app import create_app
from extensions import db
from models import User
from werkzeug.security import generate_password_hash

app = create_app()
with app.app_context():
    # Asegurar que las tablas existen antes de consultar
    db.create_all()

    # Verificar de nuevo las tablas (opcional, para debug)
    try:
        tables = db.engine.table_names()
    except Exception:
        tables = [t.name for t in db.metadata.sorted_tables]
    print("Tablas en la DB (post create_all):", tables)

    # Crear admin si no existe
    if not User.query.filter_by(username="admin").first():
        admin = User(
            username="admin",
            full_name="Administrador",
            email="admin@example.com",
            role="super_admin",
            password_hash=generate_password_hash("admin123"),
            active=True
        )
        db.session.add(admin)
        db.session.commit()
        print("Admin creado: username=admin password=admin123")
    else:
        print("Admin ya existe")
