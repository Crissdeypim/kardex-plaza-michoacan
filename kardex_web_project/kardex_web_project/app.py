
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    kardex = request.files.get("kardex")
    reporte = request.files.get("reporte")

    if not kardex or not reporte:
        return "Debes subir ambos archivos."

    return "Archivos recibidos correctamente."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
