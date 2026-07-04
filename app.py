#!/usr/bin/env python3
"""
app.py

Interfaz web (Flask) para generar_informes.py.

Permite subir un Excel, generar el informe Word y descargarlo.

Uso:
    python app.py
    # y abrir http://127.0.0.1:5000 en el navegador

Dependencias:
    pip install flask pandas openpyxl matplotlib python-docx
"""

import os
import tempfile
import uuid
from pathlib import Path

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.utils import secure_filename

# Reutilizamos la lógica ya existente del script.
from generar_informes import cargar_datos, generar_grafica, generar_word

app = Flask(__name__)
app.secret_key = "gestoria-martinez-dev"  # solo para los mensajes flash de sesión.

# Carpeta temporal donde guardamos las subidas y los informes generados.
CARPETA_TRABAJO = Path(tempfile.gettempdir()) / "gestoria_informes"
CARPETA_TRABAJO.mkdir(exist_ok=True)

EXTENSIONES_OK = {".xlsx", ".xls"}


@app.route("/", methods=["GET"])
def index():
    # Si ya hay un informe generado en esta sesión, mostramos el botón de descarga.
    informe_listo = bool(session.get("informe"))
    return render_template("index.html", informe_listo=informe_listo)


@app.route("/generar", methods=["POST"])
def generar():
    archivo = request.files.get("excel")
    if archivo is None or archivo.filename == "":
        flash("Selecciona primero un archivo Excel.", "error")
        return redirect(url_for("index"))

    nombre = secure_filename(archivo.filename)
    if Path(nombre).suffix.lower() not in EXTENSIONES_OK:
        flash("El archivo debe ser un Excel (.xlsx o .xls).", "error")
        return redirect(url_for("index"))

    # Identificador único para no pisar archivos entre peticiones.
    ident = uuid.uuid4().hex
    ruta_excel = CARPETA_TRABAJO / f"{ident}_{nombre}"
    ruta_img = CARPETA_TRABAJO / f"{ident}_grafica.png"
    ruta_docx = CARPETA_TRABAJO / f"{ident}_informe.docx"
    archivo.save(ruta_excel)

    try:
        df = cargar_datos(ruta_excel)
        total_ingresos = df.loc[df["categoria"] == "ingreso", "importe"].sum()
        total_gastos = df.loc[df["categoria"] == "gasto", "importe"].sum()
        generar_grafica(total_ingresos, total_gastos, ruta_img)
        generar_word(df, ruta_img, ruta_docx)
    except (ValueError, FileNotFoundError) as e:
        flash(f"No se pudo generar el informe: {e}", "error")
        return redirect(url_for("index"))
    except Exception as e:  # noqa: BLE001 - mostramos cualquier fallo inesperado.
        flash(f"Error inesperado al procesar el archivo: {e}", "error")
        return redirect(url_for("index"))

    # Guardamos en sesión la ruta y un nombre amigable para la descarga.
    session["informe"] = str(ruta_docx)
    session["nombre_descarga"] = f"informe_{Path(nombre).stem}.docx"
    flash("Informe generado correctamente.", "ok")
    return redirect(url_for("index"))


@app.route("/descargar", methods=["GET"])
def descargar():
    ruta = session.get("informe")
    if not ruta or not Path(ruta).exists():
        flash("No hay ningún informe generado todavía.", "error")
        return redirect(url_for("index"))
    return send_file(
        ruta,
        as_attachment=True,
        download_name=session.get("nombre_descarga", "informe.docx"),
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
