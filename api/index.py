from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_file,
    session,
)
import io
from openpyxl import load_workbook  # ← questa mancava
from werkzeug.utils import secure_filename
from api.utils import (
    serialize_workbook,
    deserialize_workbook,
    get_headers,
    mark_row_status,
    generate_copy_text,
)

app = Flask(__name__)
app.secret_key = "dexcelb-super-secret-2025-v4.8"

TEMPLATES = {
    "convertita": (
        "-{NOME} {COGNOME};\n"
        "- {TELEFONO};\n"
        "- {MQ};\n"
        "- {INDIRIZZO};\n"
        "(Già chiamato, si aspetta una chiamata in giornata)"
    ),
    "non_conv": (
        "-{NOME} {COGNOME};\n"
        "- {TELEFONO};\n"
        "- {MQ};\n"
        "- {INDIRIZZO};\n"
        "Passata non convertita, continuiamo a provare a contattarla."
    ),
}


@app.route("/", methods=["GET", "POST"])
def index():
    """
    GET  -> restituisce la pagina HTML
    POST -> upload del DB Excel
    """
    if request.method == "POST" and "excel_file" in request.files:
        file = request.files["excel_file"]
        if file and file.filename:
            filename = secure_filename(file.filename)
            stream = io.BytesIO(file.read())

            # Carica workbook da memoria
            wb = load_workbook(stream, data_only=True)

            # Salva nella sessione come JSON
            session["workbook"] = serialize_workbook(wb)
            session["headers"] = get_headers(wb)

            return jsonify(
                {
                    "status": "DB caricato",
                    "filename": filename,
                    "headers": session["headers"],
                }
            )

        return jsonify({"error": "Nessun file selezionato"}), 400

    return render_template(
        "index.html",
        db_loaded=bool(session.get("workbook")),
        headers=session.get("headers", []),
    )


@app.route("/api/table", methods=["GET"])
def get_table():
    """
    Restituisce i dati della tabella in JSON.
    """
    if "workbook" not in session:
        return jsonify({"error": "Nessun DB caricato"}), 400

    wb = deserialize_workbook(session["workbook"])
    ws = wb.active

    headers = session.get("headers") or [cell.value for cell in ws[1]]
    data = []

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx == 1:
            continue  # salta header
        data.append([str(c) if c is not None else "" for c in row])

    return jsonify({"headers": headers, "data": data})


@app.route("/api/mark_row", methods=["POST"])
def mark_row():
    """
    Aggiorna lo stato di una riga: convertita / non_conv / clear
    row_idx: indice riga 0-based lato tabella (senza header).
    """
    if "workbook" not in session:
        return jsonify({"error": "Nessun DB caricato"}), 400

    payload = request.get_json(force=True)
    row_idx = int(payload.get("row_idx", 0))  # indice 0-based della tabella
    status = payload.get("status")

    wb = deserialize_workbook(session["workbook"])

    # +2: perché row 1 = header, row 2 = prima riga dati
    excel_row_idx = row_idx + 2

    wb = mark_row_status(wb, excel_row_idx, status)
    session["workbook"] = serialize_workbook(wb)

    return jsonify({"success": True})


@app.route("/api/copy_text/<int:row_idx>", methods=["GET"])
def copy_text(row_idx):
    """
    Restituisce il testo formattato per la riga selezionata.
    row_idx: indice 0-based lato tabella.
    """
    if "workbook" not in session:
        return jsonify({"error": "Nessun DB caricato"}), 400

    wb = deserialize_workbook(session["workbook"])

    # +2 come sopra (header + offset)
    excel_row_idx = row_idx + 2
    text = generate_copy_text(wb, excel_row_idx, TEMPLATES)

    return jsonify({"text": text})


@app.route("/download", methods=["GET"])
def download_db():
    """
    Scarica il DB Excel modificato.
    """
    if "workbook" not in session:
        return jsonify({"error": "Nessun DB caricato"}), 400

    wb = deserialize_workbook(session["workbook"])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="dexcelb.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    # Per debug locale, su Vercel viene ignorato
    app.run(debug=True)
