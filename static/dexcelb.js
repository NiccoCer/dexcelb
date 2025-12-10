let table, dbLoaded = false, headers = [];
const toast = new bootstrap.Toast(document.getElementById('liveToast'));

$(document).ready(function() {
    setupFileDragDrop();
    $('#dbFile').on('change', checkFile);
});

function checkFile() {
    const file = document.getElementById('dbFile').files[0];
    document.getElementById('loadBtn').disabled = !file;
}

function loadDB() {
    const file = document.getElementById('dbFile').files[0];
    if (!file) return showToast('Seleziona un file Excel', 'warning');
    
    const formData = new FormData();
    formData.append('excel_file', file);
    
    fetch('/', {
        method: 'POST',
        body: formData
    })
    .then(r => r.json())
    .then(data => {
        if (data.status === 'DB caricato') {
            dbLoaded = true;
            headers = data.headers;
            document.getElementById('dbStatus').innerHTML = 
                `<span class="text-success"><i class="fas fa-check"></i> DB caricato: ${file.name}</span>`;
            document.getElementById('loadBtn').innerText = 'Ricarica';
            showTable();
            showToast('DB caricato con successo!', 'success');
        }
    }).catch(err => showToast('Errore caricamento: ' + err, 'danger'));
}

function showTable() {
    if (!dbLoaded) return showToast('Carica prima un DB', 'warning');
    
    document.getElementById('dashboardCards').style.display = 'none';
    document.getElementById('tableContainer').style.display = 'block';
    
    fetch('/api/table')
    .then(r => r.json())
    .then(data => {
        if (data.error) return showToast(data.error, 'warning');
        
        $('#dataTable thead').html(`
            ${data.headers.map(h => `<th>${h || ''}</th>`).join('')}
            <th style="width:80px">Azioni</th>
        `);
        
        $('#dataTable tbody').html(
            data.data.map((row, idx) => `
                <tr data-row="${idx}">
                    ${row.map(cell => `<td>${cell}</td>`).join('')}
                    <td>
                        <button class="btn btn-sm btn-success" onclick="markRow(${idx}, 'convertita')"><i class="fas fa-check"></i></button>
                        <button class="btn btn-sm btn-warning" onclick="markRow(${idx}, 'non_conv')"><i class="fas fa-clock"></i></button>
                        <button class="btn btn-sm btn-info" onclick="copyRow(${idx})"><i class="fas fa-copy"></i></button>
                    </td>
                </tr>
            `).join('')
        );
        
        initDataTable();
    });
}

function initDataTable() {
    if ($.fn.DataTable.isDataTable('#dataTable')) {
        table.destroy();
    }
    table = $('#dataTable').DataTable({
        pageLength: 50,
        order: [[0, 'asc']],
        language: { search: 'Filtra:', paginate: { first: '<<', last: '>>', next: '>', previous: '<' } }
    });
}

function markRow(rowIdx, status) {
    fetch('/api/mark_row', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({row_idx: rowIdx, status: status})
    }).then(() => {
        showToast('Riga aggiornata!', 'success');
        showTable(); // Refresh
    });
}

function markSelected(status) {
    const selected = table.rows({selected: true}).data().toArray();
    if (selected.length === 0) return showToast('Seleziona righe', 'warning');
    // Batch update logic
}

function copyRow(rowIdx) {
    fetch(`/api/copy_text/${rowIdx}`)
    .then(r => r.json())
    .then(data => {
        navigator.clipboard.writeText(data.text);
        showToast('Testo copiato negli appunti!', 'success');
    });
}

function copySelected() {
    const selectedRows = table.rows({selected: true}).data().toArray();
    if (selectedRows.length === 0) return showToast('Seleziona righe', 'warning');
    // Copy first selected
    copyRow(selectedRows[0][0]);
}

function downloadDB() {
    if (!dbLoaded) return showToast('Carica un DB prima', 'warning');
    window.location.href = '/download';
}

function showToast(msg, type = 'success') {
    const toastEl = document.getElementById('liveToast');
    toastEl.querySelector('.toast-header').className = `toast-header bg-${type === 'success' ? 'success' : type === 'warning' ? 'warning' : 'danger'} text-white`;
    toastEl.querySelector('.toast-body').textContent = msg;
    toast.show();
}

function clearFilters() {
    table.search('').draw();
}

function setupFileDragDrop() {
    const dropZone = document.getElementById('dbFile').closest('.input-group');
    dropZone.addEventListener('dragover', e => e.preventDefault());
    dropZone.addEventListener('drop', e => {
        e.preventDefault();
        document.getElementById('dbFile').files = e.dataTransfer.files;
        checkFile();
    });
}

// Select rows on click
$(document).on('click', '#dataTable tbody tr', function() {
    $(this).toggleClass('selected');
});
