from flask import Blueprint, render_template, request, redirect, url_for, flash, session, send_file
from datetime import datetime
import os
from utils.db import get_db_connection
from utils.informes_pdf import generar_pdf_visita, generar_pdf_mensual
from utils.informes_ppt import generar_ppt_visita
from utils.informes_excel import exportar_excel_visitas

visitas_bp = Blueprint('visitas', __name__)

@visitas_bp.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    user_id = session['user_id']
    rol = session['rol']
    mes_filtro = request.args.get('mes', '')

    conn = get_db_connection()
    cur = conn.cursor()

    where_conditions = []
    params = []

    if mes_filtro:
        where_conditions.append("fecha LIKE %s")
        params.append(f"{mes_filtro}%")

    if rol not in ['admin', 'jefe']:
        where_conditions.append("usuario_id = %s")
        params.append(user_id)

    where_clause = " WHERE " + " AND ".join(where_conditions) if where_conditions else ""
    cur.execute(f"SELECT * FROM visitas{where_clause} ORDER BY id DESC", params)
    visitas = cur.fetchall()

    total_visitas = len(visitas)
    visitas_mes = total_visitas if mes_filtro else 0

    if not mes_filtro:
        mes_actual = datetime.now().strftime("%Y-%m")
        count_params = [f"{mes_actual}%"]
        if rol not in ['admin', 'jefe']:
            cur.execute("SELECT COUNT(*) FROM visitas WHERE fecha LIKE %s AND usuario_id = %s", count_params + [user_id])
        else:
            cur.execute("SELECT COUNT(*) FROM visitas WHERE fecha LIKE %s", (f"{mes_actual}%",))
        visitas_mes = cur.fetchone()['count']

    meta_mensual = 30
    porcentaje_meta = (visitas_mes / meta_mensual * 100) if meta_mensual > 0 else 0

    cur.execute(f"SELECT nivel, COUNT(*) FROM visitas{where_clause} GROUP BY nivel", params)
    niveles = dict(cur.fetchall())

    cur.execute(f"SELECT tipo_visita, COUNT(*) FROM visitas{where_clause} GROUP BY tipo_visita", params)
    tipos = dict(cur.fetchall())

    conn.close()

    return render_template('dashboard.html',
                           visitas=visitas,
                           total_visitas=total_visitas,
                           visitas_mes=visitas_mes,
                           porcentaje_meta=porcentaje_meta,
                           niveles=niveles,
                           tipos=tipos,
                           mes_filtro=mes_filtro)

def generar_numero_informe():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT COALESCE(MAX(id), 0) FROM visitas")
    last_id = cur.fetchone()['coalesce'] + 1
    conn.close()
    año = datetime.now().year
    return f"INF-{año}-{last_id:03d}"

@visitas_bp.route('/nueva_visita', methods=['GET', 'POST'])
def nueva_visita():
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    if request.method == 'POST':
        user_id = session['user_id']
        datos = {k: request.form.get(k, '') for k in ['fecha', 'institucion', 'nivel', 'tipo', 'fortalezas', 'mejoras', 'recomendaciones', 'compromisos']}
        numero_informe = generar_numero_informe()

        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute('''
            INSERT INTO visitas (
                usuario_id, numero_informe, fecha, institucion, nivel, tipo_visita,
                fortalezas, mejoras, recomendaciones, compromisos
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        ''', (
            user_id, numero_informe, datos['fecha'], datos['institucion'], datos['nivel'], datos['tipo'],
            datos['fortalezas'], datos['mejoras'], datos['recomendaciones'], datos['compromisos']
        ))
        conn.commit()
        conn.close()
        flash('Visita registrada con éxito')
        return redirect(url_for('visitas.dashboard'))

    return render_template('nueva_visita.html')

@visitas_bp.route('/generar_pdf/<int:visita_id>')
def generar_pdf(visita_id):
    return generar_pdf_visita(visita_id)

@visitas_bp.route('/generar_ppt/<int:visita_id>')
def generar_ppt(visita_id):
    return generar_ppt_visita(visita_id)

@visitas_bp.route('/exportar_excel')
def exportar_excel():
    return exportar_excel_visitas()

@visitas_bp.route('/generar_informe_mensual/<mes>')
def generar_informe_mensual(mes):
    return generar_pdf_mensual(mes)
