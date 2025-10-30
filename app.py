from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, render_template, request, redirect, url_for, flash, session
import os
import psycopg2
from psycopg2.extras import RealDictCursor
from datetime import datetime
db_initialized = False

app = Flask(__name__)
app.secret_key = "ugel_lauricocha_2025"

# Conexión a PostgreSQL
def get_db_connection():
    conn = psycopg2.connect(
        os.environ['DATABASE_URL'],
        cursor_factory=RealDictCursor
    )
    return conn

# Inicializar base de datos
def init_db():
    conn = get_db_connection()
    cur = conn.cursor()
    # Tabla de usuarios
    cur.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY,
            usuario TEXT UNIQUE NOT NULL,
            contrasena TEXT NOT NULL,
            nombre_completo TEXT,
            rol TEXT NOT NULL CHECK (rol IN ('especialista', 'jefe', 'admin')),
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # Tabla de visitas
    cur.execute('''
        CREATE TABLE IF NOT EXISTS visitas (
            id SERIAL PRIMARY KEY,
            usuario_id INTEGER REFERENCES usuarios(id),
            numero_informe TEXT UNIQUE,
            fecha TEXT,
            institucion TEXT,
            nivel TEXT,
            tipo_visita TEXT,
            observaciones TEXT,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # Usuario administrador
    cur.execute('''
        INSERT INTO usuarios (usuario, contrasena, nombre_completo, rol)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (usuario) DO NOTHING
    ''', ('admin', '123456', 'Administrador', 'admin'))
    
    conn.commit()
    cur.close()
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
        cur = conn.cursor()
        cur.execute("SELECT * FROM usuarios WHERE usuario = %s", (usuario,))
        user = cur.fetchone()
        conn.close()
        
        # Validar SOLO si user existe y está activo
        if user and user['activo'] and check_password_hash(user['contrasena'], contrasena):
            session['user_id'] = user['id']
            session['rol'] = user['rol']
            session['nombre'] = user['nombre_completo']
            return redirect(url_for('dashboard'))
        else:
            flash('Usuario inactivo o credenciales incorrectos')
    
    # Si es GET o el login falló, muestra el formulario
    return render_template('login.html')



@app.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    user_id = session['user_id']
    rol = session['rol']

    conn = get_db_connection()
    cur = conn.cursor()
    
    #Filtrar visitas por rol
    if rol in ['admin', 'jefe']:
        cur.execute("SELECT * FROM visitas ORDER BY id DESC")
    else:
        cur.execute("SELECT * FROM visitas WHERE usuario_id = %s ORDER BY id DESC", (user_id,))
    visitas = cur.fetchall()

    # Totales
    total_visitas = len(visitas)

    # Visitas este mes
    mes_actual = datetime.now().strftime("%Y-%m")
    if rol in ['admin', 'jefe']:
        cur.execute("SELECT COUNT(*) FROM visitas WHERE fecha LIKE %s", (f"{mes_actual}%",))
    else:
        cur.execute("SELECT COUNT(*) FROM visitas WHERE usuario_id = %s AND fecha LIKE %s", (user_id, f"{mes_actual}%",))
    visitas_mes = cur.fetchone()['count']

    # Conteo por nivel
    if rol in ['admin', 'jefe']:
        cur.execute("SELECT nivel, COUNT(*) FROM visitas GROUP BY nivel")
    else:
        cur.execute("SELECT nivel, COUNT(*) FROM visitas WHERE usuario_id = %s GROUP BY nivel", (user_id,))
    niveles = dict(cur.fetchall())

    # Conteo por tipo
    if rol in ['admin', 'jefe']:
        cur.execute("SELECT tipo_visita, COUNT(*) FROM visitas GROUP BY tipo_visita")
    else:
        cur.execute("SELECT tipo_visita, COUNT(*) FROM visitas WHERE usuario_id = %s GROUP BY tipo_visita", (user_id,))
    tipos = dict(cur.fetchall())
    
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
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM visitas")
    count = cur.fetchone()['count'] + 1
    conn.close()
    año = datetime.now().year
    return f"INF-{año}-{count:03d}"

@app.route('/nueva_visita', methods=['GET', 'POST'])
def nueva_visita():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    if request.method == 'POST':
        user_id = session['user_id']
        fecha = request.form['fecha']
        institucion = request.form['institucion']
        nivel = request.form['nivel']
        tipo = request.form['tipo']
        especialista = session['nombre'] # Usa el nombre del usuario logueado
        observaciones = request.form['observaciones']

        conn = get_db_connection()
        cur = conn.cursor()
        numero_informe = generar_numero_informe()
        cur.execute('''
            INSERT INTO visitas (usuario_id, numero_informe, fecha, institucion, nivel, tipo_visita, observaciones)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        ''', (user_id, numero_informe, fecha, institucion, nivel, tipo, observaciones))
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
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    user_id = session['user_id']
    rol = session['rol']

    conn = get_db_connection()
    cur = conn.cursor()

    if rol in ['admin', 'jefe']:
        cur.execute("SELECT * FROM visitas WHERE id = %s", (visita_id,))
    else:
        cur.execute("SELECT * FROM visitas WHERE id = %s AND usuario_id = %s", (visita_id, user_id))
    visita = cur.fetchone()
    conn.close()

    if not visita:
        flash("Visita no encontrada o no autorizada")
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
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    user_id = session['user_id']
    rol = session['rol']

    conn = get_db_connection()
    cur = conn.cursor()

    if rol in ['admin', 'jefe']:
        cur.execute("SELECT * FROM visitas WHERE id = %s", (visita_id,))
    else:
        cur.execute("SELECT * FROM visitas WHERE id = %s AND usuario_id = %s", (visita_id, user_id))
    visita = cur.fetchone()
    conn.close()
    
    if not visita:
        flash("Visita no encontrada o no autorizada")
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

@app.route('/gestion_usuarios')
def gestion_usuarios():
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('dashboard'))
    
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT id, usuario, nombre_completo, rol, activo, creado_en FROM usuarios ORDER BY creado_en DESC")
    usuarios = cur.fetchall()
    conn.close()
    return render_template('gestion_usuarios.html', usuarios=usuarios)



@app.route('/crear_usuario', methods=['POST'])
def crear_usuario():
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('dashboard'))
    
    usuario = request.form['usuario']
    contrasena = request.form['contrasena']
    nombre_completo = request.form['nombre_completo']
    rol = request.form['rol']
    
    # Hashear la contraseña
    contrasena_hash = generate_password_hash(contrasena)
    
    conn = get_db_connection()
    cur = conn.cursor()
    try:
        cur.execute('''
            INSERT INTO usuarios (usuario, contrasena, nombre_completo, rol)
            VALUES (%s, %s, %s, %s)
        ''', (usuario, contrasena_hash, nombre_completo, rol))
        conn.commit()
        flash('Usuario creado exitosamente')
    except psycopg2.IntegrityError:
        conn.rollback()
        flash('Error: El nombre de usuario ya existe')
    finally:
        conn.close()
    
    return redirect(url_for('gestion_usuarios'))

@app.route('/exportar_excel')
def exportar_excel():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    user_id = session['user_id']
    rol = session['rol']
    conn = get_db_connection()
    cur = conn.cursor()

    if rol in ['admin', 'jefe']:
        cur.execute("SELECT numero_informe, fecha, institucion, nivel, tipo_visita, observaciones FROM visitas ORDER BY id DESC")
    else:
        cur.execute("SELECT numero_informe, fecha, institucion, nivel, tipo_visita, observaciones FROM visitas WHERE usuario_id = %s ORDER BY id DESC", (user_id,))
    visitas = cur.fetchall()
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

@app.before_request
def initialize():
    global db_initialized
    if not db_initialized:
        init_db()
        db_initialized = True

@app.route('/editar_usuario/<int:usuario_id>')
def editar_usuario(usuario_id):
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('dashboard'))
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM usuarios WHERE id = %s", (usuario_id,))
    usuario = cur.fetchone()
    conn.close()
    if not usuario:
        flash('Usuario no encontrado')
        return redirect(url_for('gestion_usuarios'))
    return render_template('editar_usuario.html', usuario=usuario)

@app.route('/actualizar_usuario/<int:usuario_id>', methods=['POST'])
def actualizar_usuario(usuario_id):
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('dashboard'))
    
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
    flash('Usuario actualizado exitosamente')
    return redirect(url_for('gestion_usuarios'))   

@app.route('/eliminar_usuario/<int:usuario_id>', methods=['POST'])
def eliminar_usuario(usuario_id):
    if 'rol' not in session or session['rol'] != 'admin':
        flash('Acceso no autorizado')
        return redirect(url_for('dashboard'))
    
    # Evita que el admin se elimine a sí mismo
    if usuario_id == session['user_id']:
        flash('No puedes eliminarte a ti mismo')
        return redirect(url_for('gestion_usuarios'))
    
    conn = get_db_connection()
    cur = conn.cursor()
    # En lugar de DELETE, desactivamos
    cur.execute("UPDATE usuarios SET activo = FALSE WHERE id = %s", (usuario_id,))
    conn.commit()
    conn.close()
    flash('Usuario desactivado exitosamente')
    return redirect(url_for('gestion_usuarios'))
@app.route('/logout')
def logout():
    session.clear()  # Elimina toda la sesión
    return redirect(url_for('login'))
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
