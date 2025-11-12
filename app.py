from flask import Flask
from routes.usuarios import usuarios_bp
from routes.visitas import visitas_bp
from routes.auth import auth_bp
from recursos.recursos import recursos_bp
from utils.db import init_db

def create_app():
    app = Flask(__name__)
    app.secret_key = "ugel_lauricocha_2025"

    with app.app_context():
        init_db()

    # Registrar Blueprints
    app.register_blueprint(auth_bp)
    app.register_blueprint(visitas_bp)
    app.register_blueprint(usuarios_bp)
    app.register_blueprint(recursos_bp)

    return app

app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
