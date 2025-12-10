import json
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

COL_CONVERTITA = "CONVERTITA"
COL_NON_CONV = "PASSATA NON CONVERTITA"


def serialize_workbook(wb):
    """
    Serializza il workbook in JSON per salvarlo nella sessione.
    Salva tutti i fogli con headers + righe.
    """
    data = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_data = {
            "headers": [cell.value for cell in ws[1]],
            "rows": []
        }
        for row in ws.iter_rows(values_only=True):
            # Salta l'header
            if row == tuple(ws[1][i].value for i in range(len(ws[1]))):
                continue
            sheet_data["rows"].append(
                [str(c) if c is not None else "" for c in row]
            )
        data[sheet_name] = sheet_data

    return json.dumps(data)


def deserialize_workbook(serialized_data):
    """
    Ricostruisce un workbook da JSON serializzato.
    Usa il primo foglio trovato.
    """
    data = json.loads(serialized_data)
    wb = Workbook()
    ws = wb.active

    # Prendi il primo sheet
    sheet_name = next(iter(data.keys()))
    sheet = data[sheet_name]

    headers = sheet["headers"] or []
    rows = sheet["rows"] or []

    # Scrivi header
    for col, header in enumerate(headers, start=1):
        ws[f"{get_column_letter(col)}1"] = header

    # Scrivi righe
    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, cell_data in enumerate(row_data, start=1):
            ws[f"{get_column_letter(col_idx)}{row_idx}"] = cell_data

    return wb


def get_headers(wb):
    ws = wb.active
    return [cell.value for cell in ws[1]]


def find_column_index(ws, header_name):
    """
    Ritorna indice 1-based della colonna con nome header_name.
    """
    if ws.max_column is None:
        return None
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if str(val).strip().upper() == str(header_name).strip().upper():
            return col
    return None


def mark_row_status(wb, row_idx, status):
    """
    status: 'convertita', 'non_conv', 'clear'
    row_idx è l'indice riga Excel (1-based).
    """
    ws = wb.active
    col_conv = find_column_index(ws, COL_CONVERTITA)
    col_non_conv = find_column_index(ws, COL_NON_CONV)

    if not col_conv or not col_non_conv:
        # Se le colonne non ci sono, non fare nulla ma non esplodere
        return wb

    if status == "clear":
        ws.cell(row=row_idx, column=col_conv).value = ""
        ws.cell(row=row_idx, column=col_non_conv).value = ""
    elif status == "convertita":
        ws.cell(row=row_idx, column=col_conv).value = "X"
        ws.cell(row=row_idx, column=col_non_conv).value = ""
    elif status == "non_conv":
        ws.cell(row=row_idx, column=col_conv).value = ""
        ws.cell(row=row_idx, column=col_non_conv).value = "X"

    return wb


def generate_copy_text(wb, row_idx, templates):
    """
    row_idx: indice riga Excel (1-based).
    templates: {'convertita': '...', 'non_conv': '...'}
    """
    ws = wb.active
    headers = [cell.value for cell in ws[1]]

    mappa = {}
    for i, h in enumerate(headers):
        key = (str(h) if h is not None else "").strip().upper()
        if not key:
            continue
        mappa[key] = ws.cell(row=row_idx, column=i + 1).value or ""

    conv_val = str(mappa.get(COL_CONVERTITA, "")).strip()
    non_conv_val = str(mappa.get(COL_NON_CONV, "")).strip()

    if conv_val:
        template = templates["convertita"]
    elif non_conv_val:
        template = templates["non_conv"]
    else:
        template = templates["non_conv"]

    return template.format(
        NOME=mappa.get("NOME", ""),
        COGNOME=mappa.get("COGNOME", ""),
        TELEFONO=mappa.get("TELEFONO", ""),
        MQ=mappa.get("MQ", ""),
        INDIRIZZO=mappa.get("INDIRIZZO", ""),
    )
