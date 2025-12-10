from flask import Flask, render_template, request, jsonify, send_file, session
import io
import json
from werkzeug.utils import secure_filename
from api.utils import *

app = Flask(__name__)
app.secret_key = 'dexcelb-super-secret-2025-v4.8'

TEMPLATES = {
    'convertita': '-{NOME} {COGNOME};\n- {TELEFONO};\n- {MQ};\n- {INDIRIZZO};\n(Già chiamato, si aspetta una chiamata in giornata)',
    'non_conv': '-{NOME} {COGNOME};\n- {TELEFONO};\n- {MQ};\n- {INDIRIZZO};\nPassata non convertita, continuiamo a provare a contattarla.'
}

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST' and 'excel_file' in request.files:
        file = request.files['excel_file']
        if file.filename:
            stream = io.BytesIO(file.read())
            wb = load_workbook(stream, data_only=True)
            session['workbook'] = serialize_workbook(wb)
            session['headers'] = get_headers(wb)
            return jsonify({'status': 'DB caricato', 'headers': session['headers']})
    
    return render_template('index.html', 
                         db_loaded=bool(session.get('workbook')),
                         headers=session.get('headers', []))

@app.route('/api/table')
def get_table():
    if 'workbook' not in session:
        return jsonify({'error': 'Nessun DB caricato'}), 400
    wb = deserialize_workbook(session['workbook'])
    ws = wb.active
    data = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        if row_idx == 1: continue
        data.append([str(c) if c else '' for c in row])
    return jsonify({'data': data, 'headers': session['headers']})

@app.route('/api/mark_row', methods=['POST'])
def mark_row():
    if 'workbook' not in session:
        return jsonify({'error': 'No DB'}), 400
    
    row_idx = int(request.json['row_idx']) + 1  # +1 per header
    status = request.json['status']
    wb = deserialize_workbook(session['workbook'])
    wb = mark_row_status(wb, row_idx, status)
    session['workbook'] = serialize_workbook(wb)
    return jsonify({'success': True})

@app.route('/api/copy_text/<int:row_idx>')
def copy_text(row_idx):
    wb = deserialize_workbook(session['workbook'])
    text = generate_copy_text(wb, row_idx + 2, TEMPLATES)  # +2 per header
    return jsonify({'text': text})

@app.route('/download')
def download_db():
    if 'workbook' not in session:
        return jsonify({'error': 'No DB'}), 400
    wb = deserialize_workbook(session['workbook'])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='dexcelb.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
