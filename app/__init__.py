from flask import Flask
from app.utils.db import init_db  # ✅ utils está dentro de app


def create_app():
    app = Flask(__name__)
    app.secret_key = "ugel_lauricocha_2025"

    # Inicializar base de datos
    with app.app_context():
        init_db()

    # Importar y registrar Blueprints
    from routes.auth import auth_bp
    from routes.visitas import visitas_bp
    from routes.usuarios import usuarios_bp
    from routes.recursos import recursos_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(visitas_bp)
    app.register_blueprint(usuarios_bp)
    app.register_blueprint(recursos_bp)

    return app

