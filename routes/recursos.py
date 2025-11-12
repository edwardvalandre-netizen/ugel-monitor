from flask import Blueprint, render_template, request, redirect, url_for, flash, session, send_file
import os

recursos_bp = Blueprint('recursos', __name__)

@recursos_bp.route('/recursos')
def recursos():
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    # Carpeta base de recursos
    recursos_path = os.path.join(os.path.dirname(__file__), '..', 'recursos')

    # Categorías de recursos
    categorias = {
        "rubricas": "Rúbricas de evaluación",
        "guias": "Guías pedagógicas",
        "planes": "Planes y proyectos",
        "normas": "Normas y documentos oficiales"
    }

    recursos_disponibles = {}
    for carpeta, descripcion in categorias.items():
        carpeta_path = os.path.join(recursos_path, carpeta)
        if os.path.exists(carpeta_path):
            archivos = os.listdir(carpeta_path)
            recursos_disponibles[descripcion] = archivos

    return render_template('recursos.html', recursos=recursos_disponibles)

@recursos_bp.route('/descargar_recurso/<categoria>/<archivo>')
def descargar_recurso(categoria, archivo):
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    recursos_path = os.path.join(os.path.dirname(__file__), '..', 'recursos')
    archivo_path = os.path.join(recursos_path, categoria, archivo)

    if os.path.exists(archivo_path):
        return send_file(archivo_path, as_attachment=True)
    else:
        flash("Archivo no encontrado")
        return redirect(url_for('recursos.recursos'))
