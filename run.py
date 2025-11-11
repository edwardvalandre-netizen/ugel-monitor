from app import create_app

app = create_app()

if __name__ == "__main__":
    # debug=True solo en desarrollo, en producción quítalo
    app.run(debug=True)
