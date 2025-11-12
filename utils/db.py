import os
import psycopg2
from psycopg2.extras import RealDictCursor
from werkzeug.security import generate_password_hash

def get_db_connection():
    conn = psycopg2.connect(
        os.environ['DATABASE_URL'],
        cursor_factory=RealDictCursor
    )
    return conn

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
            activo BOOLEAN DEFAULT TRUE,
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
            fortalezas TEXT,
            mejoras TEXT,
            recomendaciones TEXT,
            compromisos TEXT,
            creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Usuario admin por defecto
    admin_pass = generate_password_hash('123456')
    cur.execute('''
        INSERT INTO usuarios (usuario, contrasena, nombre_completo, rol, activo)
        VALUES (%s, %s, %s, %s, %s)
        ON CONFLICT (usuario) DO NOTHING
    ''', ('admin', admin_pass, 'Administrador', 'admin', True))

    conn.commit()
    cur.close()
    conn.close()
