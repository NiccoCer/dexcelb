# ===== api/index.py - COMPLETO E FUNZIONANTE =====

import os
import io
import uuid
from datetime import datetime, timedelta

from flask import Flask, render_template, request, jsonify, send_file
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

# ===== STORAGE IN MEMORIA CON COOKIE =====
SESSIONS = {}
SESSION_TIMEOUT = timedelta(hours=1)


def cleanup_expired_sessions():
    """Rimuove sessioni scadute"""
    now = datetime.now()
    expired = [sid for sid, data in SESSIONS.items() 
               if now - data['created'] > SESSION_TIMEOUT]
    for sid in expired:
        del SESSIONS[sid]


def get_or_create_session_id(request):
    """Ottiene o crea un ID sessione dal browser"""
    session_id = request.cookies.get("dexcelb_session_id")
    if not session_id or session_id not in SESSIONS:
        session_id = str(uuid.uuid4())
        SESSIONS[session_id] = {
            'workbook': None,
            'headers': [],
            'created': datetime.now()
        }
    return session_id


# ===== ROUTE PRINCIPALE =====
@app.route("/", methods=["GET", "POST"])
def index():
    """
    GET  -> Restituisce la pagina HTML
    POST -> Upload del file Excel DB
    """
    cleanup_expired_sessions()
    session_id = get_or_create_session_id(request)
    
    if request.method == "POST" and "excel_file" in request.files:
        file = request.files["excel_file"]
        if file and file.filename:
            filename = secure_filename(file.filename)
            stream = io.BytesIO(file.read())
            
            try:
                wb = load_workbook(stream, data_only=True)
                SESSIONS[session_id]["workbook"] = serialize_workbook(wb)
                SESSIONS[session_id]["headers"] = get_headers(wb)
                
                response = jsonify({
                    "status": "DB caricato",
                    "filename": filename,
                    "headers": SESSIONS[session_id]["headers"],
                })
                response.set_cookie("dexcelb_session_id", session_id, max_age=3600, samesite="Lax")
                return response
            except Exception as e:
                return jsonify({"error": f"Errore caricamento: {str(e)}"}), 400
        
        return jsonify({"error": "Nessun file selezionato"}), 400

    response = jsonify({"ok": True})
    response.set_cookie("dexcelb_session_id", session_id, max_age=3600, samesite="Lax")
    return render_template("index.html")


# ===== API: TABELLA =====
@app.route("/api/table", methods=["GET"])
def get_table():
    """Restituisce dati della tabella in JSON"""
    cleanup_expired_sessions()
    session_id = get_or_create_session_id(request)
    
    if not SESSIONS[session_id]["workbook"]:
        return jsonify({"error": "Nessun DB caricato"}), 400

    try:
        wb = deserialize_workbook(SESSIONS[session_id]["workbook"])
        ws = wb.active
        headers = SESSIONS[session_id]["headers"] or get_headers(wb)
        data = []

        for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row_idx == 1:
                continue
            data.append([str(c) if c else "" for c in row])

        return jsonify({"headers": headers, "data": data})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===== API: AGGIORNA STATO RIGA =====
@app.route("/api/mark_row", methods=["POST"])
def mark_row():
    """Segna riga come convertita/non_conv/clear"""
    cleanup_expired_sessions()
    session_id = get_or_create_session_id(request)
    
    if not SESSIONS[session_id]["workbook"]:
        return jsonify({"error": "Nessun DB caricato"}), 400

    try:
        payload = request.get_json(force=True)
        row_idx = int(payload.get("row_idx", 0))
        status = payload.get("status")

        wb = deserialize_workbook(SESSIONS[session_id]["workbook"])
        excel_row_idx = row_idx + 2
        mark_row_status(wb, excel_row_idx, status)
        SESSIONS[session_id]["workbook"] = serialize_workbook(wb)

        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===== API: COPIA TESTO =====
@app.route("/api/copy_text/<int:row_idx>", methods=["GET"])
def copy_text(row_idx):
    """Genera testo formattato per riga"""
    cleanup_expired_sessions()
    session_id = get_or_create_session_id(request)
    
    if not SESSIONS[session_id]["workbook"]:
        return jsonify({"error": "Nessun DB caricato"}), 400

    try:
        wb = deserialize_workbook(SESSIONS[session_id]["workbook"])
        excel_row_idx = row_idx + 2
        text = generate_copy_text(wb, excel_row_idx, TEMPLATES)

        return jsonify({"text": text})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===== API: INSERIMENTO MANUALE =====
@app.route("/api/add_row", methods=["POST"])
def add_row():
    """Aggiunge riga manualmente"""
    cleanup_expired_sessions()
    session_id = get_or_create_session_id(request)
    
    if not SESSIONS[session_id]["workbook"]:
        return jsonify({"error": "Nessun DB caricato"}), 400

    try:
        payload = request.get_json(force=True)
        row_values = payload.get("values", [])

        wb = deserialize_workbook(SESSIONS[session_id]["workbook"])
        add_row_to_workbook(wb, row_values)
        SESSIONS[session_id]["workbook"] = serialize_workbook(wb)

        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===== API: ELIMINA RIGA =====
@app.route("/api/delete_row/<int:row_idx>", methods=["POST"])
def delete_row(row_idx):
    """Elimina una riga"""
    cleanup_expired_sessions()
    session_id = get_or_create_session_id(request)
    
    if not SESSIONS[session_id]["workbook"]:
        return jsonify({"error": "Nessun DB caricato"}), 400

    try:
        wb = deserialize_workbook(SESSIONS[session_id]["workbook"])
        excel_row_idx = row_idx + 2
        delete_row_from_workbook(wb, excel_row_idx)
        SESSIONS[session_id]["workbook"] = serialize_workbook(wb)

        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===== API: UNISCI FILE =====
@app.route("/api/merge", methods=["POST"])
def merge_files():
    """Unisce file sorgente nel DB master"""
    cleanup_expired_sessions()
    session_id = get_or_create_session_id(request)
    
    if not SESSIONS[session_id]["workbook"]:
        return jsonify({"error": "Nessun DB caricato"}), 400

    if "merge_file" not in request.files:
        return jsonify({"error": "Nessun file da unire"}), 400

    try:
        file = request.files["merge_file"]
        stream = io.BytesIO(file.read())

        wb_master = deserialize_workbook(SESSIONS[session_id]["workbook"])
        wb_source = load_workbook(stream, data_only=True)

        wb_merged, count = merge_workbooks(wb_master, wb_source, ["NOME", "COGNOME", "ZONA"])
        SESSIONS[session_id]["workbook"] = serialize_workbook(wb_merged)

        return jsonify({"success": True, "imported": count})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ===== API: SCARICA DB =====
@app.route("/download", methods=["GET"])
def download_db():
    """Scarica il DB Excel modificato"""
    cleanup_expired_sessions()
    session_id = get_or_create_session_id(request)
    
    if not SESSIONS[session_id]["workbook"]:
        return jsonify({"error": "Nessun DB caricato"}), 400

    try:
        wb = deserialize_workbook(SESSIONS[session_id]["workbook"])
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="dexcelb.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)

# ===== FINE api/index.py =====
