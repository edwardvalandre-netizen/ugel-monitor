from flask import session, redirect, url_for, flash, send_file
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.colors import Color, black, gray, whitesmoke
from reportlab.lib.units import inch
from utils.db import get_db_connection
import os

def generar_pdf_visita(visita_id):
    if 'user_id' not in session:
        return redirect(url_for('auth.login'))

    user_id = session['user_id']
    rol = session['rol']

    conn = get_db_connection()
    cur = conn.cursor()

    if rol in ['admin', 'jefe']:
        cur.execute("""
            SELECT v.*, u.nombre_completo as especialista_nombre
            FROM visitas v JOIN usuarios u ON v.usuario_id = u.id
            WHERE v.id = %s
        """, (visita_id,))
    else:
        cur.execute("""
            SELECT v.*, u.nombre_completo as especialista_nombre
            FROM visitas v JOIN usuarios u ON v.usuario_id = u.id
            WHERE v.id = %s AND v.usuario_id = %s
        """, (visita_id, user_id))

    visita = cur.fetchone()
    conn.close()

    if not visita:
        flash("Visita no encontrada o no autorizada")
        return redirect(url_for('visitas.dashboard'))

    filename = f"informe_visita_{visita_id}.pdf"
    filepath = os.path.join(os.path.dirname(__file__), filename)

    doc = SimpleDocTemplate(filepath, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    logo_path = os.path.join(os.path.dirname(__file__), 'templates_ppt', 'logo.png')
    if os.path.exists(logo_path):
        im = Image(logo_path, 1.5*inch, 0.8*inch)
        im.hAlign = 'CENTER'
        story.append(im)
        story.append(Spacer(1, 24))

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], alignment=TA_CENTER, textColor=Color(0, 0.2, 0.4))
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=12, alignment=TA_LEFT)

    story.append(Paragraph("INFORME DE VISITA PEDAGÓGICA", title_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<b>Número de Informe:</b> {visita['numero_informe']}", normal_style))
    story.append(Spacer(1, 12))

    campos = [
        ("Institución Educativa", visita['institucion']),
        ("Nivel", visita['nivel']),
        ("Tipo de Visita", visita['tipo_visita']),
        ("Especialista", visita['especialista_nombre']),
        ("Fecha", visita['fecha']),
    ]
    for label, value in campos:
        story.append(Paragraph(f"<b>{label}:</b> {value}", normal_style))

    story.append(Spacer(1, 18))
    story.append(Paragraph("<b>OBSERVACIONES ESTRUCTURADAS</b>", styles['Heading3']))
    for campo, texto in [
        ("✅ Fortalezas", visita['fortalezas']),
        ("⚠️ Áreas de mejora", visita['mejoras']),
        ("💡 Recomendaciones", visita['recomendaciones']),
        ("📝 Compromisos del docente", visita['compromisos']),
    ]:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<b>{campo}:</b>", normal_style))
        story.append(Paragraph(texto or "No registradas.", normal_style))

    story.append(Spacer(1, 36))
    story.append(Paragraph("_____________________________________", normal_style))
    story.append(Paragraph("Firma del Especialista", normal_style))

    doc.build(story)
    return send_file(filepath, as_attachment=True)

def generar_pdf_mensual(mes):
    # Puedes mover aquí tu lógica de informe mensual si lo deseas
    pass
