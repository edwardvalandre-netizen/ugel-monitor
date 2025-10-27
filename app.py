from flask import Flask, render_template, request, redirect, url_for, flash
import sqlite3
from datetime import datetime

app = Flask(__name__)
app.secret_key = "ugel_lauricocha_2025"  # para mensajes flash

# Inicializar base de datos
def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS visitas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_informe TEXT UNIQUE,
            fecha TEXT,
            institucion TEXT,
            nivel TEXT,
            tipo_visita TEXT,
            especialista TEXT,
            observaciones TEXT,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE,
            contrasena TEXT  -- en producción usar hash, pero para prácticas OK
        )
    ''')
    # Usuario de prueba
    c.execute("INSERT OR IGNORE INTO usuarios (usuario, contrasena) VALUES ('admin', '123456')")
    conn.commit()
    conn.close()

@app.route('/')
def index():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form['usuario']
        contrasena = request.form['contrasena']
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT * FROM usuarios WHERE usuario = ? AND contrasena = ?", (usuario, contrasena))
        user = c.fetchone()
        conn.close()
        if user:
            return redirect(url_for('dashboard'))
        else:
            flash('Usuario o contraseña incorrectos')
    return render_template('login.html')

from datetime import datetime

@app.route('/dashboard')
def dashboard():
    conn = get_db_connection()
    c = conn.cursor()
    
    # Todas las visitas
    c.execute("SELECT * FROM visitas ORDER BY id DESC")
    visitas = c.fetchall()
    
    # Totales
    total_visitas = len(visitas)
    
    # Visitas este mes
    mes_actual = datetime.now().strftime("%Y-%m")
    c.execute("SELECT COUNT(*) FROM visitas WHERE fecha LIKE ?", (f"{mes_actual}%",))
    visitas_mes = c.fetchone()[0]
    
    # Conteo por nivel
    c.execute("SELECT nivel, COUNT(*) FROM visitas GROUP BY nivel")
    niveles = dict(c.fetchall())
    
    # Conteo por tipo
    c.execute("SELECT tipo_visita, COUNT(*) FROM visitas GROUP BY tipo_visita")
    tipos = dict(c.fetchall())
    
    conn.close()
    
    return render_template('dashboard.html', 
                         visitas=visitas,
                         total_visitas=total_visitas,
                         visitas_mes=visitas_mes,
                         niveles=niveles,
                         tipos=tipos)

def generar_numero_informe():
    """Genera un número de informe único: INF-2025-001, INF-2025-002, etc."""
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM visitas")
    count = c.fetchone()[0] + 1
    conn.close()
    año = datetime.now().year
    return f"INF-{año}-{count:03d}"
@app.route('/nueva_visita', methods=['GET', 'POST'])
def nueva_visita():
    if request.method == 'POST':
        fecha = request.form['fecha']
        institucion = request.form['institucion']
        nivel = request.form['nivel']
        tipo = request.form['tipo']
        especialista = request.form['especialista']
        observaciones = request.form['observaciones']

        conn = get_db_connection()
        c = conn.cursor()
        numero_informe = generar_numero_informe()
        c.execute('''
            INSERT INTO visitas (numero_informe, fecha, institucion, nivel, tipo_visita, especialista, observaciones)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (numero_informe, fecha, institucion, nivel, tipo, especialista, observaciones))
        conn.commit()
        conn.close()
        flash('Visita registrada con éxito')
        return redirect(url_for('dashboard'))
    return render_template('nueva_visita.html')

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os
from flask import send_file

@app.route('/generar_ppt/<int:visita_id>')
def generar_ppt(visita_id):
    # Obtener datos de la visita
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT * FROM visitas WHERE id = ?", (visita_id,))
    visita = c.fetchone()
    conn.close()

    if not visita:
        flash("Visita no encontrada")
        return redirect(url_for('dashboard'))

    # Crear presentación
    prs = Presentation()
    prs.slide_width = Inches(10)  # Tamaño estándar
    prs.slide_height = Inches(7.5)

    # --- Diapositiva 1: Portada ---
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Diseño en blanco

    # Logo (izquierda superior)
    logo_path = os.path.join(os.path.dirname(__file__), 'templates_ppt', 'logo.png')
    if os.path.exists(logo_path):
        slide.shapes.add_picture(logo_path, Inches(0.5), Inches(0.3), height=Inches(0.8))

    # Título central
    title_box = slide.shapes.add_textbox(Inches(0), Inches(2.5), Inches(10), Inches(2))
    tf = title_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"REPORTE DE VISITA PEDAGÓGICA\n{visita[1]}"  # ← Correcto
    p.font.bold = True
    p.font.size = Pt(28)
    p.font.color.rgb = RGBColor(0, 51, 102)  # Azul UGEL
    p.alignment = PP_ALIGN.CENTER

    # Subtítulo
    sub_box = slide.shapes.add_textbox(Inches(0), Inches(4.2), Inches(10), Inches(1))
    tf2 = sub_box.text_frame
    tf2.clear()
    p2 = tf2.paragraphs[0]
    p2.text = f"Institución: {visita[3]}\nEspecialista: {visita[6]}\nFecha: {visita[2]} | Tipo: {visita[5]}"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(0, 0, 0)
    p2.alignment = PP_ALIGN.CENTER

    # --- Diapositiva 2: Observaciones ---
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])

    # Título
    title2_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    tf3 = title2_box.text_frame
    tf3.clear()
    p3 = tf3.paragraphs[0]
    p3.text = "OBSERVACIONES Y RECOMENDACIONES"
    p3.font.bold = True
    p3.font.size = Pt(22)
    p3.font.color.rgb = RGBColor(0, 51, 102)

    # Contenido
    obs_box = slide2.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5))
    tf4 = obs_box.text_frame
    tf4.clear()
    p4 = tf4.paragraphs[0]
    p4.text = f"Responsable: {visita[6]}\n\n{visita[7] if visita[7] else 'Sin observaciones.'}"
    p4.font.size = Pt(16)
    p4.font.color.rgb = RGBColor(0, 0, 0)

    # Guardar en memoria
    filename = f"reporte_visita_{visita_id}.pptx"
    filepath = os.path.join(os.path.dirname(__file__), filename)
    prs.save(filepath)

    # Enviar archivo para descarga
    return send_file(filepath, as_attachment=True)
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import Color
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os
from flask import send_file

@app.route('/generar_pdf/<int:visita_id>')
def generar_pdf(visita_id):
    # Obtener datos de la visita
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT * FROM visitas WHERE id = ?", (visita_id,))
    visita = c.fetchone()
    conn.close()

    if not visita:
        flash("Visita no encontrada")
        return redirect(url_for('dashboard'))

    # Nombre del archivo
    filename = f"informe_visita_{visita_id}.pdf"
    filepath = os.path.join(os.path.dirname(__file__), filename)

    # Crear PDF
    doc = SimpleDocTemplate(filepath, pagesize=letter,
                            rightMargin=72, leftMargin=72,
                            topMargin=72, bottomMargin=72)
    styles = getSampleStyleSheet()

    # Estilo personalizado
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=14,
        textColor=Color(0, 51/255, 102/255),  # Azul UGEL
        alignment=TA_CENTER
    )

    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=12,
        spaceAfter=10,
        alignment=TA_LEFT
    )

    # Contenido
    story = []

    # Logo
    logo_path = os.path.join(os.path.dirname(__file__), 'templates_ppt', 'logo.png')
    if os.path.exists(logo_path):
        im = Image(logo_path, 1.5*inch, 0.8*inch)
        im.hAlign = 'CENTER'
        story.append(im)
        story.append(Spacer(1, 24))

    # Título
    story.append(Paragraph("INFORME DE VISITA PEDAGÓGICA", title_style))
    story.append(Spacer(1, 12))

    # Número de informe
    story.append(Paragraph(f"<b>Número de Informe:</b> {visita[1]}", normal_style))
    story.append(Spacer(1, 12))

    # Datos
    datos = [
        f"<b>Institución Educativa:</b> {visita[3]}",
        f"<b>Nivel:</b> {visita[4]}",
        f"<b>Tipo de Visita:</b> {visita[5]}",
        f"<b>Especialista:</b> {visita[6]}",
        f"<b>Fecha:</b> {visita[2]}",
    ]

    for dato in datos:
        story.append(Paragraph(dato, normal_style))

    story.append(Spacer(1, 18))

    # Observaciones
    story.append(Paragraph("<b>OBSERVACIONES Y RECOMENDACIONES</b>", normal_style))
    story.append(Spacer(1, 6))
    obs = visita[7] if visita[7] else "Sin observaciones registradas."
    story.append(Paragraph(obs, normal_style))

    story.append(Spacer(1, 36))

    # Espacio para firma
    story.append(Paragraph("_____________________________________", normal_style))
    story.append(Paragraph("Firma del Especialista", normal_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("_____________________________________", normal_style))
    story.append(Paragraph("Sello de la Institución Educativa", normal_style))

    # Generar PDF
    doc.build(story)

    return send_file(filepath, as_attachment=True)
from openpyxl import Workbook
from flask import send_file
import os

@app.route('/exportar_excel')
def exportar_excel():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT numero_informe, fecha, institucion, nivel, tipo_visita, especialista, observaciones FROM visitas ORDER BY id DESC")
    visitas = c.fetchall()
    conn.close()

    # Crear libro de Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Visitas Pedagógicas"

    # Encabezados
    headers = ["N° Informe", "Fecha", "Institución", "Nivel", "Tipo de Visita", "Especialista", "Observaciones"]
    ws.append(headers)

    # Datos
    for visita in visitas:
        ws.append(visita)

    # Ajustar ancho de columnas
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = min(adjusted_width, 50)

    # Guardar
    filename = "visitas_pedagogicas.xlsx"
    filepath = os.path.join(os.path.dirname(__file__), filename)
    wb.save(filepath)

    return send_file(filepath, as_attachment=True)
if __name__ == '__main__':
    init_db()
    app.run()