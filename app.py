
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from pathlib import Path
from datetime import datetime, date
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
import unicodedata
import re
import uuid
import os

app = Flask(__name__)
app.secret_key = "kardex-plaza-michoacan-web"

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

PURPLE_FILL = PatternFill(start_color="D9B3FF", end_color="D9B3FF", fill_type="solid")
HEADER_FILL = PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid")

def norm(v):
    if v is None:
        return ""
    txt = str(v).strip().upper()
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = re.sub(r"[^A-Z0-9 ]+", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()

def norm_name(v):
    return " ".join(sorted(norm(v).split()))

def clean_uni(v):
    if v is None:
        return ""
    txt = str(v).replace(".0", "").strip()
    txt = re.sub(r"\D", "", txt)
    return txt.lstrip("0")

def is_apto(v):
    s = norm(v)
    return s == "APTO"

def as_dt(v):
    if v in [None, ""]:
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and 30000 <= float(v) <= 60000:
        try:
            return from_excel(v)
        except Exception:
            pass
    txt = str(v).strip()
    for fmt in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(txt.split(".")[0], fmt)
        except Exception:
            pass
    return None

def canonical_course(v):
    c = norm(v)

    # Equivalencias especiales
    if c == "PLD" or "PREVENCION DE LAVADO DE DINERO" in c:
        return "PLD"

    if "PROTECCION DE DATOS PERSONALES" in c:
        return "PROTECCION DE DATOS PERSONALES"

    if "CONOCE EL MODULO DE APRENDIZAJE" in c and "USUARIO FINAL" in c:
        return "WORKDAY MOD APRENDIZAJE USUARIO FINAL"

    if "CONOCE EL MODULO DE APRENDIZAJE" in c and "COORDINADOR DE FORMACION" in c:
        return "WORKDAY MOD APRENDIZAJE COORD FORMACION"

    if "WORKDAY" in c and "APRENDIZAJE" in c and "USUARIO FINAL" in c:
        return "WORKDAY MOD APRENDIZAJE USUARIO FINAL"

    if "WORKDAY" in c and "APRENDIZAJE" in c and ("COORD FORMACION" in c or "COORDINADOR DE FORMACION" in c):
        return "WORKDAY MOD APRENDIZAJE COORD FORMACION"

    # ORACLE SIM y ORACLE SIM ROC son distintos
    if "ORACLE SIM ROC" in c or ("ORACLE SIM" in c and "ROC" in c):
        return "ORACLE SIM ROC"

    if c in ["ORACLE SIM", "SISTEMA ORACLE SIM"]:
        return "ORACLE SIM"

    if "SICAD" in c:
        return "SICAD"

    if "FARMAX SAD" in c:
        return "FARMAX SAD"

    # IMPORTANTE: NO mezclar estos cursos
    if c in ["TECNICAS DE VENTA", "TECNICAS DE VENTAS"]:
        return "TECNICAS DE VENTAS"

    if c in ["TECNICAS DE VENTA BASICO", "TECNICAS DE VENTAS BASICO"]:
        return "TECNICAS DE VENTAS BASICO"

    aliases = {
        "NET PAY": "NETPAY",
        "NETPAY": "NETPAY",
        "AHORROPAGOS": "AHORRO PAGOS",
        "AHORROCEL": "AHORRO PAGOS",
        "AHORRO PAGOS": "AHORRO PAGOS",
        "BUEN USO DEL MONEDERO DEL AHORRO": "MONEDERO DEL AHORRO",
        "EL PODER DE LA MARCA PROPIA 2026": "MARCA PROPIA",
        "EL PODER DE LA MARCA PROPIA A": "MARCA PROPIA",
        "CUENTA MONEDERO SINTETIZADA A": "CUENTA MONEDERO",
        "CUENTA MONEDERO SINTETIZADA STAFF": "CUENTA MONEDERO",
        "CONECTA INVENTARIOS": "CONECTA",
        "SERVICIO Y VENTA DE 2 ESFUERZO": "SERVICIO Y VENTAS DE 2 ESFUERZO",
        "DISPENSACION A": "SICAD",
    }
    return aliases.get(c, c)

def find_report_headers(ws):
    ws.reset_dimensions()
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=50, values_only=True), start=1):
        values = [norm(v) for v in row]
        if "UNI" in values and "CURSO" in values:
            return row_idx, {norm(v): i for i, v in enumerate(row) if v is not None}
    raise ValueError("No encontré la fila de encabezados en el reporte. Debe contener UNI y Curso.")

def read_report(report_path):
    wb = load_workbook(report_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row, headers = find_report_headers(ws)

    required = ["UNI", "NOMBRE", "CURSO", "FECHA DE FINALIZACION", "PUNTUACION DE CURSO", "EVALUACION DE CURSO"]
    missing = [x for x in required if x not in headers]
    if missing:
        raise ValueError("Faltan columnas en el reporte: " + ", ".join(missing))

    records = {}
    valid_read = 0
    ignored = 0

    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        try:
            u = clean_uni(row[headers["UNI"]])
            name_raw = row[headers["NOMBRE"]]
            course_raw = row[headers["CURSO"]]
            fecha_raw = row[headers["FECHA DE FINALIZACION"]]
            score_raw = row[headers["PUNTUACION DE CURSO"]]
            eval_raw = row[headers["EVALUACION DE CURSO"]]
        except Exception:
            continue

        if not u or not name_raw or not course_raw:
            continue

        try:
            score = float(score_raw)
        except Exception:
            ignored += 1
            continue

        if score < 80:
            ignored += 1
            continue

        if not is_apto(eval_raw):
            ignored += 1
            continue

        dt = as_dt(fecha_raw)
        if dt is None or dt.year == 1900:
            ignored += 1
            continue

        key = (u, norm_name(name_raw), canonical_course(course_raw))
        valid_read += 1

        if key not in records or dt > records[key]["dt"]:
            records[key] = {
                "dt": datetime(dt.year, dt.month, dt.day),
                "fecha": datetime(dt.year, dt.month, dt.day),
                "score": int(score) if float(score).is_integer() else score,
                "course_raw": str(course_raw),
            }

    return records, valid_read, ignored

def detect_blocks(ws):
    blocks = []
    for col in range(1, ws.max_column + 1):
        if norm(ws.cell(5, col).value) == "FECHA DE EJECUCION" and norm(ws.cell(5, col + 1).value) == "CALIFICACION":
            course_name = ""
            for k in range(col, max(0, col - 10), -1):
                val = ws.cell(4, k).value
                if val and norm(val) not in ["FECHA DE EJECUCION", "CALIFICACION", "FECHA DE PROGRAMACION"]:
                    course_name = str(val).strip()
                    break
            blocks.append({
                "course_original": course_name,
                "course_key": canonical_course(course_name),
                "fecha_col": col,
                "cal_col": col + 1,
                "prog_col": col - 1,
            })
    return blocks

def correct_excel_dates(ws):
    converted = 0
    formatted = 0
    fecha_cols = []
    for col in range(1, ws.max_column + 1):
        header = norm(ws.cell(5, col).value)
        if header in ["FECHA DE EJECUCION", "FECHA DE PROGRAMACION", "FECHA DE FINALIZACION"]:
            fecha_cols.append(col)

    for col in fecha_cols:
        for row in range(6, ws.max_row + 1):
            cell = ws.cell(row, col)
            val = cell.value
            if isinstance(val, (int, float)) and 30000 <= float(val) <= 60000:
                dt = from_excel(val)
                cell.value = datetime(dt.year, dt.month, dt.day)
                cell.number_format = "DD/MM/YYYY"
                converted += 1
            elif isinstance(val, datetime):
                cell.value = datetime(val.year, val.month, val.day)
                cell.number_format = "DD/MM/YYYY"
                formatted += 1
            elif isinstance(val, date):
                cell.number_format = "DD/MM/YYYY"
                formatted += 1
    return converted, formatted

def clean_tecnicas_ventas_na(ws, blocks):
    """
    Limpia SOLO Técnicas de Ventas cuando Fecha de Programación no aplica.
    No toca Técnicas de Ventas Básico.
    """
    cleaned_rows = 0
    for block in blocks:
        if block["course_key"] != "TECNICAS DE VENTAS":
            continue

        prog_col = block["prog_col"]
        fecha_col = block["fecha_col"]
        cal_col = block["cal_col"]

        # Puestos permitidos, detectados desde fórmula típica.
        allowed_positions = set()
        for row in range(6, ws.max_row + 1):
            prog_cell = ws.cell(row, prog_col)
            if isinstance(prog_cell.value, str) and prog_cell.value.startswith("="):
                refs = re.findall(r"\$?[A-Z]{1,3}\$?\d+", prog_cell.value)
                for ref in refs:
                    clean = ref.replace("$", "")
                    if not clean.startswith("H") and not clean.startswith("E"):
                        try:
                            allowed_positions.add(norm(ws[clean].value))
                        except Exception:
                            pass
                break

        for row in range(6, ws.max_row + 1):
            prog_cell = ws.cell(row, prog_col)
            fecha_cell = ws.cell(row, fecha_col)
            cal_cell = ws.cell(row, cal_col)

            should_clear = False
            if norm(prog_cell.value) == "NA":
                should_clear = True
            elif allowed_positions:
                current_position = norm(ws.cell(row, 8).value)
                if current_position not in allowed_positions:
                    should_clear = True

            if should_clear:
                had_data = fecha_cell.value not in [None, ""] or cal_cell.value not in [None, ""]
                if had_data:
                    if not (isinstance(fecha_cell.value, str) and fecha_cell.value.startswith("=")):
                        fecha_cell.value = None
                    if not (isinstance(cal_cell.value, str) and cal_cell.value.startswith("=")):
                        cal_cell.value = None
                    cleaned_rows += 1
    return cleaned_rows

def update_kardex(kardex_path, report_path, output_path):
    records, valid_read, ignored = read_report(report_path)

    wb = load_workbook(kardex_path, keep_links=False)
    ws = wb["Incorporacion"] if "Incorporacion" in wb.sheetnames else wb[wb.sheetnames[0]]

    blocks = detect_blocks(ws)

    people_rows = {}
    for row in range(6, ws.max_row + 1):
        u = clean_uni(ws.cell(row, 6).value)
        name = norm_name(ws.cell(row, 7).value)
        if u and name:
            people_rows[(u, name)] = row

    updates = []
    fechas = 0
    califs = 0
    skipped_formula = 0

    for block in blocks:
        for (u, name, course_key), rec in records.items():
            if course_key != block["course_key"]:
                continue

            row = people_rows.get((u, name))
            if not row:
                continue

            fcell = ws.cell(row, block["fecha_col"])
            ccell = ws.cell(row, block["cal_col"])
            changed = False

            if isinstance(fcell.value, str) and fcell.value.startswith("="):
                skipped_formula += 1
            else:
                current_dt = as_dt(fcell.value)
                if fcell.value in [None, ""] or current_dt is None or rec["dt"] > current_dt:
                    fcell.value = rec["fecha"]
                    fcell.number_format = "DD/MM/YYYY"
                    fcell.fill = PURPLE_FILL
                    fechas += 1
                    changed = True

            if isinstance(ccell.value, str) and ccell.value.startswith("="):
                skipped_formula += 1
            else:
                if ccell.value != rec["score"]:
                    ccell.value = rec["score"]
                    ccell.number_format = "0"
                    ccell.fill = PURPLE_FILL
                    califs += 1
                    changed = True

            if changed:
                updates.append([
                    u,
                    ws.cell(row, 7).value,
                    rec["course_raw"],
                    block["course_original"],
                    rec["fecha"],
                    rec["score"],
                ])

    tecnicas_cleaned = clean_tecnicas_ventas_na(ws, blocks)
    seriales, fechas_formateadas = correct_excel_dates(ws)

    # Hojas de auditoría
    for sh in ["Resumen_Actualizacion", "Actualizaciones", "Reglas_Cursos"]:
        if sh in wb.sheetnames:
            del wb[sh]

    summary = wb.create_sheet("Resumen_Actualizacion")
    summary_rows = [
        ["Concepto", "Cantidad"],
        ["Registros válidos únicos", len(records)],
        ["Registros aptos >=80 leídos", valid_read],
        ["Actualizaciones realizadas", len(updates)],
        ["Fechas cargadas/actualizadas", fechas],
        ["Calificaciones cargadas/actualizadas", califs],
        ["Ignorados por regla", ignored],
        ["Cursos detectados en Kardex", len(blocks)],
        ["Celdas con fórmula respetadas", skipped_formula],
        ["Técnicas de Ventas limpiadas por NA", tecnicas_cleaned],
        ["Seriales de Excel convertidos a fecha", seriales],
        ["Fechas formateadas sin hora", fechas_formateadas],
    ]

    for r, row in enumerate(summary_rows, 1):
        for c, value in enumerate(row, 1):
            cell = summary.cell(r, c, value)
            if r == 1:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = HEADER_FILL
    summary.column_dimensions["A"].width = 48
    summary.column_dimensions["B"].width = 20

    act = wb.create_sheet("Actualizaciones")
    act.append(["UNI", "Nombre", "Curso Reporte", "Curso Kardex", "Fecha", "Calificación"])
    for cell in act[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
    for row in updates:
        act.append(row)
    for col, width in enumerate([14, 40, 65, 55, 22, 18], 1):
        act.column_dimensions[get_column_letter(col)].width = width

    rules = wb.create_sheet("Reglas_Cursos")
    rules.append(["Regla", "Detalle"])
    rules.append(["Técnicas de Ventas", "No se mezcla con Técnicas de Ventas Básico. Son cursos distintos."])
    rules.append(["Workday Usuario Final", "Conoce el módulo de aprendizaje | Usuario final = Workday Mód. Aprendizaje Usuario Final"])
    rules.append(["Workday Coord Formación", "Conoce el módulo de aprendizaje | Coordinador de Formación = Workday Mód. Aprendizaje Coord Formación"])
    rules.append(["PLD", "Prevención de lavado de dinero (PLD) = PLD"])
    rules.append(["Protección Datos", "Protección de Datos Personales período (2025-2026) = Protección de datos personales"])
    for cell in rules[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
    rules.column_dimensions["A"].width = 35
    rules.column_dimensions["B"].width = 100

    for sheet in [summary, act, rules]:
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(output_path)

    return {
        "filename": output_path.name,
        "validos": len(records),
        "leidos": valid_read,
        "actualizaciones": len(updates),
        "fechas": fechas,
        "calificaciones": califs,
        "ignorados": ignored,
        "cursos": len(blocks),
        "formulas": skipped_formula,
        "tecnicas": tecnicas_cleaned,
        "seriales": seriales,
        "fechas_formateadas": fechas_formateadas,
    }

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/actualizar", methods=["POST"])
def actualizar():
    try:
        kardex = request.files.get("kardex")
        reporte = request.files.get("reporte")

        if not kardex or not reporte:
            flash("Debes subir el Kardex y el reporte.")
            return redirect(url_for("index"))

        job_id = uuid.uuid4().hex[:8]
        kardex_path = UPLOAD_DIR / f"kardex_{job_id}.xlsx"
        report_path = UPLOAD_DIR / f"reporte_{job_id}.xlsx"
        output_path = OUTPUT_DIR / f"KARDEX_ACTUALIZADO_{job_id}.xlsx"

        kardex.save(kardex_path)
        reporte.save(report_path)

        result = update_kardex(kardex_path, report_path, output_path)
        return render_template("resultado.html", result=result)

    except Exception as e:
        flash(f"Error al procesar: {str(e)}")
        return redirect(url_for("index"))

@app.route("/descargar/<filename>")
def descargar(filename):
    path = OUTPUT_DIR / filename
    return send_file(path, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
