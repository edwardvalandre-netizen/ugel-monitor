from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from werkzeug.security import generate_password_hash
from utils.db import get_db_connection
import psycopg2

usuarios_bp = Blueprint('usuarios', __name__)

@usuarios_bp.route('/gestion_usuarios')
def gestion_usuarios():
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('visitas.dashboard'))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT id, usuario, nombre_completo, rol, activo, creado_en FROM usuarios ORDER BY creado_en DESC")
    usuarios = cur.fetchall()
    conn.close()
    return render_template('gestion_usuarios.html', usuarios=usuarios)

@usuarios_bp.route('/crear_usuario', methods=['POST'])
def crear_usuario():
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado', 'danger')   # ⚠️ Aquí usamos "danger"
        return redirect(url_for('visitas.dashboard'))

    usuario = request.form['usuario']
    contrasena = request.form['contrasena']
    nombre_completo = request.form['nombre_completo']
    rol = request.form['rol']

    contrasena_hash = generate_password_hash(contrasena)

    conn = get_db_connection()
    cur = conn.cursor()
    try:
        cur.execute('''
            INSERT INTO usuarios (usuario, contrasena, nombre_completo, rol, activo)
            VALUES (%s, %s, %s, %s, %s)
        ''', (usuario, contrasena_hash, nombre_completo, rol, True))
        conn.commit()
        flash('Usuario creado exitosamente', 'success')   # ✅ Mensaje de éxito
    except psycopg2.IntegrityError:
        conn.rollback()
        flash('Error: El nombre de usuario ya existe', 'danger')   # ❌ Mensaje de error
    finally:
        conn.close()

    return redirect(url_for('usuarios.gestion_usuarios'))


@usuarios_bp.route('/editar_usuario/<int:usuario_id>')
def editar_usuario(usuario_id):
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('visitas.dashboard'))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM usuarios WHERE id = %s", (usuario_id,))
    usuario = cur.fetchone()
    conn.close()

    if not usuario:
        flash('Usuario no encontrado')
        return redirect(url_for('usuarios.gestion_usuarios'))

    return render_template('editar_usuario.html', usuario=usuario)

@usuarios_bp.route('/actualizar_usuario/<int:usuario_id>', methods=['POST'])
def actualizar_usuario(usuario_id):
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('visitas.dashboard'))

    nombre_completo = request.form['nombre_completo']
    rol = request.form['rol']
    activo = 'activo' in request.form
    contrasena = request.form['contrasena']

    conn = get_db_connection()
    cur = conn.cursor()

    if contrasena:
        contrasena_hash = generate_password_hash(contrasena)
        cur.execute('''
            UPDATE usuarios 
            SET nombre_completo = %s, rol = %s, activo = %s, contrasena = %s
            WHERE id = %s
        ''', (nombre_completo, rol, activo, contrasena_hash, usuario_id))
    else:
        cur.execute('''
            UPDATE usuarios 
            SET nombre_completo = %s, rol = %s, activo = %s
            WHERE id = %s
        ''', (nombre_completo, rol, activo, usuario_id))

    conn.commit()
    conn.close()
    flash('Usuario actualizado exitosamente', 'success')   # ✅ Mensaje de éxito
    return redirect(url_for('usuarios.gestion_usuarios'))

@usuarios_bp.route('/eliminar_usuario/<int:usuario_id>', methods=['POST'])
def eliminar_usuario(usuario_id):
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('visitas.dashboard'))

    if usuario_id == session['user_id']:
        flash('No puedes eliminarte a ti mismo')
        return redirect(url_for('usuarios.gestion_usuarios'))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("UPDATE usuarios SET activo = FALSE WHERE id = %s", (usuario_id,))
    conn.commit()
    conn.close()
    flash('Usuario desactivado exitosamente', 'success')   # ✅ Mensaje de éxito
    return redirect(url_for('usuarios.gestion_usuarios'))
