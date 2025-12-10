const API_URL = '/api'; // Vercel gestirà il routing
const COL_CONVERTITA = "CONVERTITA";
const COL_NON_CONV = "PASSATA NON CONVERTITA";

let currentData = { colonne: [], righe: [], db_name: '(nessuno)' };
let currentTemplates = {};

document.addEventListener('DOMContentLoaded', () => {
    refreshData();
    document.getElementById('merge-files').addEventListener('change', updateMergeFileList);
    loadTemplates();
});

// --- FUNZIONI UTILITY ---

function showSection(sectionId) {
    document.querySelectorAll('#main-stack > section').forEach(section => {
        section.style.display = 'none';
    });
    document.getElementById(sectionId).style.display = 'block';

    if (sectionId === 'dashboard') {
        document.getElementById('btn-back-menu').style.display = 'none';
    } else {
        document.getElementById('btn-back-menu').style.display = 'inline-block';
    }
    
    // Logica specifica per sezione
    if (sectionId === 'table') {
        renderTable(currentData.righe);
    }
    if (sectionId === 'manual') {
        renderManualForm();
    }
}

async function apiFetch(endpoint, options = {}) {
    try {
        const response = await fetch(`${API_URL}${endpoint}`, options);
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail || `Errore HTTP: ${response.status}`);
        }
        return await response.json();
    } catch (error) {
        console.error("API Error:", error);
        alert(`Errore: ${error.message}`);
        return null;
    }
}

function updateStatus(id, message, isError = false) {
    const el = document.getElementById(id);
    el.textContent = message;
    el.style.color = isError ? '#dc2626' : '#4ade80';
}

// --- GESTIONE DATI E UI ---

async function refreshData() {
    const data = await apiFetch('/data');
    if (data) {
        currentData = data;
        document.getElementById('db-status').textContent = `DB in uso: ${data.db_name}`;
        
        const clientiCount = data.righe.length > 0 ? data.righe.length - 1 : 0;
        document.getElementById('table-status').textContent = `Righe caricate: ${data.righe.length} | Clienti: ${clientiCount}`;
        
        // Aggiorna la combo di filtro
        const filterCol = document.getElementById('filter-col');
        filterCol.innerHTML = '';
        data.colonne.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            filterCol.appendChild(option);
        });

        // Se siamo nella sezione tabella, ricarica
        if (document.getElementById('table').style.display !== 'none') {
             renderTable(currentData.righe);
        }
        
        // Se siamo nella sezione manuale, ricarica il form
        if (document.getElementById('manual').style.display !== 'none') {
             renderManualForm();
        }
    }
}

function renderTable(righe) {
    const tableHeadRow = document.querySelector('#data-table thead tr');
    const tableBody = document.querySelector('#data-table tbody');
    tableHeadRow.innerHTML = '<th>#</th>';
    tableBody.innerHTML = '';

    if (!currentData.colonne || currentData.colonne.length === 0) {
        document.getElementById('table-status').textContent = "Nessuna colonna nel DB.";
        return;
    }

    currentData.colonne.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        tableHeadRow.appendChild(th);
    });

    // Trova gli indici delle colonne di stato
    const convIdx = currentData.colonne.indexOf(COL_CONVERTITA);
    const nonConvIdx = currentData.colonne.indexOf(COL_NON_CONV);

    righe.forEach(riga => {
        const tr = document.createElement('tr');
        
        // Tag di classe per il colore condizionale
        let rowClass = '';
        if (convIdx >= 0 && (riga.valori[convIdx] || '').toUpperCase() === 'X') {
            rowClass = 'row-convertita';
        } else if (nonConvIdx >= 0 && (riga.valori[nonConvIdx] || '').toUpperCase() === 'X') {
            rowClass = 'row-non-convertita';
        }
        tr.className = rowClass;
        
        // Data attribute per identificare la riga Excel e i valori
        tr.dataset.rigaExcel = riga.riga_excel;
        tr.dataset.valori = JSON.stringify(riga.valori);

        // Colonna #
        const tdIdx = document.createElement('td');
        tdIdx.textContent = riga.riga_excel;
        tr.appendChild(tdIdx);

        // Altre colonne
        riga.valori.forEach(val => {
            const td = document.createElement('td');
            td.textContent = val;
            tr.appendChild(td);
        });

        // Aggiungi listener per selezionare la riga
        tr.addEventListener('click', () => {
            document.querySelectorAll('#data-table tbody tr').forEach(row => {
                row.classList.remove('selected');
            });
            tr.classList.add('selected');
        });
        
        tableBody.appendChild(tr);
    });

    const righeCount = righe.length;
    const clientiCount = righe.length > 0 ? righe.length - 1 : 0;
    document.getElementById('table-status').textContent = `Righe visualizzate: ${righeCount} | Clienti: ${clientiCount}`;
}

// --- GESTIONE TABELLA (FILTRI / STATO) ---

function applyFilter() {
    const colName = document.getElementById('filter-col').value;
    const searchText = document.getElementById('filter-text').value.toLowerCase().trim();

    if (!colName || searchText === '') {
        renderTable(currentData.righe);
        return;
    }

    const colIndex = currentData.colonne.indexOf(colName);
    if (colIndex === -1) {
        renderTable(currentData.righe);
        return;
    }

    const filteredRows = currentData.righe.filter(riga => {
        // Mantiene l'intestazione
        if (riga.riga_excel === 1) return true; 

        const cellValue = riga.valori[colIndex];
        return (String(cellValue || '').toLowerCase()).includes(searchText);
    });

    renderTable(filteredRows);
}

function resetFilter() {
    document.getElementById('filter-text').value = '';
    renderTable(currentData.righe);
}


async function setRowStatus(convertita, non_convertita, pulisci = false) {
    const selectedRow = document.querySelector('#data-table tbody tr.selected');
    if (!selectedRow) {
        alert("Seleziona una riga prima di modificare lo stato.");
        return;
    }
    const rigaExcel = parseInt(selectedRow.dataset.rigaExcel);

    if (rigaExcel === 1) {
        alert("Non puoi modificare la riga di intestazione.");
        return;
    }
    
    document.querySelector('.action-bar').style.opacity = '0.5';

    const result = await apiFetch('/row/status', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ riga_excel: rigaExcel, convertita, non_convertita, pulisci })
    });

    document.querySelector('.action-bar').style.opacity = '1';
    if (result) {
        alert(result.messaggio);
        await refreshData();
    }
}

function copyRowText() {
    const selectedRow = document.querySelector('#data-table tbody tr.selected');
    if (!selectedRow) {
        alert("Seleziona una riga prima di copiare il testo.");
        return;
    }
    if (parseInt(selectedRow.dataset.rigaExcel) === 1) {
        alert("Non puoi copiare la riga di intestazione.");
        return;
    }

    const valori = JSON.parse(selectedRow.dataset.valori);
    const headers = currentData.colonne;
    
    // Mappa i valori per nome colonna
    const mappa = {};
    headers.forEach((h, i) => {
        mappa[h.toUpperCase()] = valori[i] || '';
    });
    
    // Determina il template da usare
    const convVal = mappa[COL_CONVERTITA.toUpperCase()] || '';
    const nonConvVal = mappa[COL_NON_CONV.toUpperCase()] || '';
    
    let template;
    if (convVal === 'X') {
        template = currentTemplates.convertita;
    } else if (nonConvVal === 'X') {
        template = currentTemplates.non_convertita;
    } else {
        template = currentTemplates.non_convertita; // Default
    }

    // Sostituisci i segnaposto
    let testoCopiato = template;
    testoCopiato = testoCopiato.replace(/{NOME}/g, mappa['NOME'] || '');
    testoCopiato = testoCopiato.replace(/{COGNOME}/g, mappa['COGNOME'] || '');
    testoCopiato = testoCopiato.replace(/{TELEFONO}/g, mappa['TELEFONO'] || '');
    testoCopiato = testoCopiato.replace(/{MQ}/g, mappa['MQ'] || '');
    testoCopiato = testoCopiato.replace(/{INDIRIZZO}/g, mappa['INDIRIZZO'] || '');

    // Copia negli appunti
    navigator.clipboard.writeText(testoCopiato.trim())
        .then(() => alert("Testo copiato negli appunti!"))
        .catch(err => alert("Errore nella copia: " + err));
}

// --- GESTIONE IMPORTAZIONE ---

async function handleImport() {
    const fileInput = document.getElementById('import-file');
    const file = fileInput.files[0];
    if (!file) {
        alert("Seleziona un file da importare.");
        return;
    }
    
    updateStatus('import-status', 'Importazione in corso...', false);
    
    const formData = new FormData();
    formData.append('file', file);

    const result = await apiFetch('/import', {
        method: 'POST',
        body: formData,
    });
    
    if (result) {
        updateStatus('import-status', result.messaggio);
        fileInput.value = ''; // Resetta il campo file
        await refreshData();
    } else {
        updateStatus('import-status', 'Importazione fallita.', true);
    }
}

// --- GESTIONE AGGIUNTA MANUALE ---

function renderManualForm() {
    const container = document.getElementById('manual-form-container');
    container.innerHTML = '';
    
    if (!currentData.colonne || currentData.colonne.length === 0) {
        container.innerHTML = `<p class="card-subtitle">Seleziona prima un DB valido.</p>`;
        return;
    }
    
    // Crea un div che sarà lo scrollable container
    const formGrid = document.createElement('div');
    formGrid.className = 'manual-form-grid';
    
    currentData.colonne.forEach((col, index) => {
        const label = document.createElement('label');
        label.textContent = `${col}:`;
        label.setAttribute('for', `manual-entry-${index}`);

        const input = document.createElement('input');
        input.type = 'text';
        input.id = `manual-entry-${index}`;
        input.name = col;

        container.appendChild(label);
        container.appendChild(input);
    });
}

async function handleAddRow() {
    const inputs = document.querySelectorAll('#manual-form-container input');
    if (inputs.length === 0) {
        alert("Carica prima il DB per generare i campi.");
        return;
    }
    
    const rowValues = Array.from(inputs).map(input => input.value);
    
    if (rowValues.every(val => val.trim() === '')) {
        alert("Compila almeno un campo.");
        return;
    }
    
    updateStatus('manual-status', 'Aggiunta riga in corso...', false);

    const result = await apiFetch('/row/add', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: rowValues })
    });
    
    if (result) {
        updateStatus('manual-status', result.messaggio);
        inputs.forEach(input => input.value = ''); // Pulisci i campi
        await refreshData();
    } else {
        updateStatus('manual-status', 'Aggiunta riga fallita.', true);
    }
}

// --- GESTIONE UNIONE (MERGE) ---

function updateMergeFileList() {
    const fileInput = document.getElementById('merge-files');
    const fileListDiv = document.getElementById('merge-file-list');
    
    if (fileInput.files.length === 0) {
        fileListDiv.textContent = "Nessun file selezionato.";
        return;
    }

    const fileNames = Array.from(fileInput.files).map(f => f.name).join('\n');
    fileListDiv.textContent = fileNames;
}

async function handleMerge() {
    const fileInput = document.getElementById('merge-files');
    const files = fileInput.files;
    if (files.length === 0) {
        alert("Seleziona almeno un file da unire.");
        return;
    }
    
    updateStatus('merge-status', 'Unione file in corso...', false);
    document.getElementById('merge').style.opacity = '0.5';

    const formData = new FormData();
    Array.from(files).forEach(file => {
        formData.append('files', file);
    });

    const result = await apiFetch('/merge', {
        method: 'POST',
        body: formData,
    });
    
    document.getElementById('merge').style.opacity = '1';
    if (result) {
        updateStatus('merge-status', result.messaggio);
        fileInput.value = ''; // Resetta il campo file
        document.getElementById('merge-file-list').textContent = "Nessun file selezionato.";
        await refreshData();
    } else {
        updateStatus('merge-status', 'Unione file fallita.', true);
    }
}

// --- GESTIONE TEMPLATE ---

async function loadTemplates() {
    const data = await apiFetch('/templates');
    if (data) {
        currentTemplates = data;
        document.getElementById('tpl-convertita').value = data.convertita;
        document.getElementById('tpl-non-convertita').value = data.non_convertita;
    }
}

async function saveTemplates() {
    const convText = document.getElementById('tpl-convertita').value;
    const nonConvText = document.getElementById('tpl-non-convertita').value;

    if (!convText.trim() || !nonConvText.trim()) {
        updateStatus('templates-status', 'I template non possono essere vuoti.', true);
        return;
    }
    
    updateStatus('templates-status', 'Salvataggio in corso...', false);

    const result = await apiFetch('/templates/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ convertita: convText, non_convertita: nonConvText })
    });

    if (result) {
        updateStatus('templates-status', result.messaggio);
        // Aggiorna i template in memoria per la copia
        currentTemplates.convertita = convText;
        currentTemplates.non_convertita = nonConvText;
    } else {
        updateStatus('templates-status', 'Salvataggio fallito.', true);
    }
}