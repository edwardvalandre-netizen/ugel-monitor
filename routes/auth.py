from flask import Blueprint, render_template, request, redirect, url_for, flash, session, make_response
from werkzeug.security import check_password_hash
from utils.db import get_db_connection

auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/')
def index():
    return redirect(url_for('auth.login'))

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form['usuario']
        contrasena = request.form['contrasena']
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT * FROM usuarios WHERE usuario = %s", (usuario,))
        user = cur.fetchone()
        conn.close()

        if user and user['activo'] and check_password_hash(user['contrasena'], contrasena):
            session['user_id'] = user['id']
            session['rol'] = user['rol']
            session['nombre'] = user['nombre_completo']
            return redirect(url_for('visitas.dashboard'))
        else:
            flash('Usuario inactivo o credenciales incorrectos')

    resp = make_response(render_template('login.html'))
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp

@auth_bp.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('auth.login'))
