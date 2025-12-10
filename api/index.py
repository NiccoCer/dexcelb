# ===== COPIA TUTTO QUESTO FILE =====

import os
import io

from flask import (
    Flask, render_template, request, jsonify, send_file, session
)
from openpyxl import load_workbook
from werkzeug.utils import secure_filename

from api.utils import (
    serialize_workbook, deserialize_workbook, get_headers,
    mark_row_status, generate_copy_text, add_row_to_workbook,
    delete_row_from_workbook, merge_workbooks,
    DEFAULT_TEMPLATE_CONVERTITA, DEFAULT_TEMPLATE_NON_CONV
)

# ===== CONFIG BASE =====
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
)

app.secret_key = "dexcelb-super-secret-2025-v4.8"

TEMPLATES = {
    "convertita": DEFAULT_TEMPLATE_CONVERTITA,
    "non_conv": DEFAULT_TEMPLATE_NON_CONV,
}


# ===== ROUTE PRINCIPALE =====
@app.route("/", methods=["GET", "POST"])
def index():
    """Upload del DB principale"""
    if request.method == "POST" and "excel_file" in request.files:
        file = request.files["excel_file"]
        if file and file.filename:
            filename = secure_filename(file.filename)
            stream = io.BytesIO(file.read())
            wb = load_workbook(stream, data_only=True)
            session["workbook"] = serialize_workbook(wb)
            session["headers"] = get_headers(wb)
            session.modified = True
            return jsonify({
                "status": "DB caricato",
                "filename": filename,
                "headers": session["headers"],
            })
        return jsonify({"error": "Nessun file"}), 400

    return render_template(
        "index.html",
        db_loaded=bool(session.get("workbook")),
        headers=session.get("headers", []),
    )


# ===== API: TABELLA =====
@app.route("/api/table", methods=["GET"])
def get_table():
    """Ritorna dati tabella"""
    if "workbook" not in session:
        return jsonify({"error": "No DB"}), 400

    wb = deserialize_workbook(session["workbook"])
    ws = wb.active
    headers = session.get("headers") or get_headers(wb)
    data = []

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx == 1:
            continue
        data.append([str(c) if c else "" for c in row])

    return jsonify({"headers": headers, "data": data})


# ===== API: AGGIORNA STATO RIGA =====
@app.route("/api/mark_row", methods=["POST"])
def mark_row():
    """Segna riga come convertita/non_conv/clear"""
    if "workbook" not in session:
        return jsonify({"error": "No DB"}), 400

    payload = request.get_json(force=True)
    row_idx = int(payload.get("row_idx", 0))
    status = payload.get("status")

    wb = deserialize_workbook(session["workbook"])
    excel_row_idx = row_idx + 2  # +1 header, +1 offset
    mark_row_status(wb, excel_row_idx, status)
    session["workbook"] = serialize_workbook(wb)
    session.modified = True

    return jsonify({"success": True})


# ===== API: COPIA TESTO =====
@app.route("/api/copy_text/<int:row_idx>", methods=["GET"])
def copy_text(row_idx):
    """Genera testo formattato per riga"""
    if "workbook" not in session:
        return jsonify({"error": "No DB"}), 400

    wb = deserialize_workbook(session["workbook"])
    excel_row_idx = row_idx + 2
    text = generate_copy_text(wb, excel_row_idx, TEMPLATES)

    return jsonify({"text": text})


# ===== API: INSERIMENTO MANUALE =====
@app.route("/api/add_row", methods=["POST"])
def add_row():
    """Aggiunge riga manualmente compilando form"""
    if "workbook" not in session:
        return jsonify({"error": "No DB"}), 400

    payload = request.get_json(force=True)
    row_values = payload.get("values", [])

    wb = deserialize_workbook(session["workbook"])
    add_row_to_workbook(wb, row_values)
    session["workbook"] = serialize_workbook(wb)
    session.modified = True

    return jsonify({"success": True})


# ===== API: ELIMINA RIGA =====
@app.route("/api/delete_row/<int:row_idx>", methods=
