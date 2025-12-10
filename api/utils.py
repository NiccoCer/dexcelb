import json
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

COL_CONVERTITA = "CONVERTITA"
COL_NON_CONV = "PASSATA NON CONVERTITA"

def serialize_workbook(wb):
    """Serializza workbook per sessione"""
    data = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        data[sheet_name] = {}
        data[sheet_name]['headers'] = [cell.value for cell in ws[1]]
        data[sheet_name]['rows'] = []
        for row in ws.iter_rows(values_only=True)[1:]:
            data[sheet_name]['rows'].append([str(c) if c is not None else '' for c in row])
    return json.dumps(data)

def deserialize_workbook(serialized_data):
    """Ricarica workbook da serializzato"""
    wb = Workbook()
    ws = wb.active
    data = json.loads(serialized_data)
    
    # Headers
    headers = data['Sheet']['headers']
    for col, header in enumerate(headers, 1):
        ws[f'{get_column_letter(col)}1'] = header
    
    # Rows
    for row_idx, row_data in enumerate(data['Sheet']['rows'], 2):
        for col_idx, cell_data in enumerate(row_data, 1):
            ws[f'{get_column_letter(col_idx)}{row_idx}'] = cell_data
    
    return wb

def get_headers(wb):
    return [cell.value for cell in wb.active[1]]

def find_column_index(ws, header_name):
    headers = [cell.value for cell in ws[1]]
    for i, h in enumerate(headers):
        if str(h).strip().upper() == str(header_name).strip().upper():
            return i + 1
    return None

def mark_row_status(wb, row_idx, status):
    """'convertita', 'non_conv', 'clear'"""
    ws = wb.active
    col_conv = find_column_index(ws, COL_CONVERTITA)
    col_non_conv = find_column_index(ws, COL_NON_CONV)
    
    if col_conv and col_non_conv:
        if status == 'clear':
            ws.cell(row=row_idx, column=col_conv).value = ""
            ws.cell(row=row_idx, column=col_non_conv).value = ""
        elif status == 'convertita':
            ws.cell(row=row_idx, column=col_conv).value = "X"
            ws.cell(row=row_idx, column=col_non_conv).value = ""
        elif status == 'non_conv':
            ws.cell(row=row_idx, column=col_conv).value = ""
            ws.cell(row=row_idx, column=col_non_conv).value = "X"
    return wb

def generate_copy_text(wb, row_idx, templates):
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    
    row_data = {}
    for i, header in enumerate(headers):
        row_data[str(header).strip().upper()] = ws.cell(row=row_idx, column=i+1).value or ""
    
    conv_val = str(row_data.get(COL_CONVERTITA, "")).strip()
    non_conv_val = str(row_data.get(COL_NON_CONV, "")).strip()
    
    if conv_val:
        template = templates['convertita']
    elif non_conv_val:
        template = templates['non_conv']
    else:
        template = templates['non_conv']
    
    return template.format(
        NOME=row_data.get('NOME', ''),
        COGNOME=row_data.get('COGNOME', ''),
        TELEFONO=row_data.get('TELEFONO', ''),
        MQ=row_data.get('MQ', ''),
        INDIRIZZO=row_data.get('INDIRIZZO', '')
    )
