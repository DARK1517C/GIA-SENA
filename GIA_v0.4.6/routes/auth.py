from flask import Blueprint, render_template, redirect, url_for, request, flash
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash
from models import User, Apprentice
from extensions import db

auth_bp = Blueprint("auth", __name__, url_prefix="/auth")

@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for("dashboard.index"))
        flash("Usuario o contraseña incorrectos", "danger")
    return render_template("auth/login.html")

@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("auth.login"))

@auth_bp.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    if request.method == "POST":
        new_password = request.form.get("new_password", "")
        confirm_password = request.form.get("confirm_password", "")
        current_password = request.form.get("current_password", "")
        if new_password:
            if not current_user.check_password(current_password):
                flash("La contraseña actual no es correcta.", "warning")
                return redirect(url_for("auth.profile"))
            if new_password != confirm_password:
                flash("La confirmacion de contraseña no coincide.", "warning")
                return redirect(url_for("auth.profile"))
            current_user.password_hash = generate_password_hash(new_password)
            db.session.commit()
            flash("Contraseña actualizada.", "success")
            return redirect(url_for("auth.profile"))

    apprentice = None
    if current_user.role == "aprendiz":
        apprentice = Apprentice.query.filter_by(student_user_id=current_user.id).first()
    return render_template("auth/profile.html", apprentice=apprentice)
