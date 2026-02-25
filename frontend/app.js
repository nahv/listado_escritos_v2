let pieChart, barChart;

function formatDate(dateString) {
    const date = new Date(dateString);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function parseSummary(info) {
    // If backend returns structured object
    if (typeof info === 'object') {
        return {
            total_records: info.total_records || 0,
            unique_exptes_count: info.unique_exptes_count || 0,
            escritos_count: info.presentaciones_count || 0,
            proyectos_count: info.proyectos_count || 0,
            oldest_escrito: info.oldest_record || '',
            days_difference: info.days_difference || 0,
            today_date: info.today_date || '',
            transferencias_count: info.transferencias_count || 0,
            most_titles: info.most_titles || [],
            period: info.period || '',
            presentaciones_by_date: info.presentaciones_by_date || []
        };
    }

    // Fallback to previous string parsing (unchanged)
    const lines = info.split('\n').map(l => l.trim()).filter(Boolean);
    let total_records = 0, unique_exptes_count = 0, escritos_count = 0, proyectos_count = 0;
    let oldest_escrito = '', days_difference = 0, today_date = '', transferencias_count = 0;
    let most_titles = [];

    lines.forEach(line => {
        if (line.match(/presentaciones en un total de/)) {
            const m = line.match(/(\d+)\s+presentaciones en un total de (\d+) causas/);
            if (m) {
                total_records = parseInt(m[1]);
                unique_exptes_count = parseInt(m[2]);
            }
        }
        if (line.match(/Escritos:/)) {
            const m = line.match(/Escritos:\s*(\d+)\s*\/\s*Proyectos:\s*(\d+)/);
            if (m) {
                escritos_count = parseInt(m[1]);
                proyectos_count = parseInt(m[2]);
            }
        }
        if (line.match(/Escrito más antiguo/)) {
            const m = line.match(/Escrito más antiguo (\d{2}\/\d{2}\/\d{4}) - (\d+) días corridos al (\d{2}\/\d{2}\/\d{4})/);
            if (m) {
                oldest_escrito = m[1];
                days_difference = parseInt(m[2]);
                today_date = m[3];
            }
        }
        if (line.match(/escritos incluyen la palabra "transferencia"/)) {
            const m = line.match(/(\d+) escritos incluyen la palabra "transferencia"/);
            if (m) {
                transferencias_count = parseInt(m[1]);
            }
        }
    });

    // Parse most_titles table
    const tableStart = info.indexOf('Escritos más repetidos:');
    if (tableStart !== -1) {
        const tableText = info.substring(tableStart).split('\n').slice(1);
        tableText.forEach((row, idx) => {
            // Skip header row (first row) and any row with NaN
            if (idx === 0) return;
            const parts = row.trim().split(/\s{2,}/);
            if (
                parts.length === 2 &&
                parts[0] &&
                parts[1] &&
                parts[0] !== 'Suma' &&
                !isNaN(parseInt(parts[1]))
            ) {
                most_titles.push({ Título: parts[0], Cantidad: parseInt(parts[1]) });
            }
        });
    }

    // Compose period string
    const period = oldest_escrito && today_date ? `${oldest_escrito} a ${today_date}` : '';

    return {
        total_records,
        unique_exptes_count,
        escritos_count,
        proyectos_count,
        oldest_escrito,
        days_difference,
        today_date,
        transferencias_count,
        most_titles,
        period
    };
}

function renderDashboard(summary) {
    // Update period header with icons
    document.getElementById('period-header').innerHTML = `
<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" fill="currentColor" class="bi bi-calendar3 mb-2" viewBox="0 0 16 16">
  <path d="M14 0H2a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V2a2 2 0 0 0-2-2M1 3.857C1 3.384 1.448 3 2 3h12c.552 0 1 .384 1 .857v10.286c0 .473-.448.857-1 .857H2c-.552 0-1-.384-1-.857z"/>
  <path d="M6.5 7a1 1 0 1 0 0-2 1 1 0 0 0 0 2m3 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2m3 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2m-9 3a1 1 0 1 0 0-2 1 1 0 0 0 0 2m3 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2m3 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2m3 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2m-9 3a1 1 0 1 0 0-2 1 1 0 0 0 0 2m3 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2m3 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2"/>
</svg> ${summary.period} <br><br>
<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" fill="currentColor" class="bi bi-calendar2-event mb-2" viewBox="0 0 16 16">
  <path d="M11 7.5a.5.5 0 0 1 .5-.5h1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-1a.5.5 0 0 1-.5-.5z"/>
  <path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5M2 2a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V3a1 1 0 0 0-1-1z"/>
  <path d="M2.5 4a.5.5 0 0 1 .5-.5h10a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5H3a.5.5 0 0 1-.5-.5z"/>
</svg>
<strong>${summary.days_difference} días corridos</strong>
<br> del primer escrito (${summary.oldest_escrito})
<br> a hoy (${summary.today_date})
`;

    // Update summary text with icons and order
    document.getElementById('summary-text').innerHTML = `
        <br>
        <strong>
        ${summary.total_records} presentaciones 
        <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" fill="currentColor" class="bi bi-journal-text mb-2" viewBox="0 0 16 16">
        <path d="M5 10.5a.5.5 0 0 1 .5-.5h2a.5.5 0 0 1 0 1h-2a.5.5 0 0 1-.5-.5m0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5m0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5m0-2a.5.5 0 0 1 .5-.5h5a.5.5 0 0 1 0 1h-5a.5.5 0 0 1-.5-.5"/>
        <path d="M3 0h10a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2v-1h1v1a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1H3a1 1 0 0 0-1 1v1H1V2a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v9a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1H3a2 2 0 0 1 2-2"/>
        <path d="M1 5v-.5a.5.5 0 0 1 1 0V5h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1zm0 3v-.5a.5.5 0 0 1 1 0V8h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1zm0 3v-.5a.5.5 0 0 1 1 0v.5h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1H2v-.5a.5.5 0 0 0-1 0"/>
        </svg>
        <br> en 
        ${summary.unique_exptes_count} expedientes 
        <svg xmlns="http://www.w3.org/2000/svg" width="34" height="34" fill="currentColor" class="bi bi-journals" viewBox="0 0 16 16">
        <path d="M5 0h8a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2 2 2 0 0 1-2 2H3a2 2 0 0 1-2-2h1a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V4a1 1 0 0 0-1-1H3a1 1 0 0 0-1 1H1a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v9a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1H3a2 2 0 0 1 2-2"/>
        <path d="M1 6v-.5a.5.5 0 0 1 1 0V6h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1zm0 3v-.5a.5.5 0 0 1 1 0V9h.5a.5.5 0 0 1 0 1h-2a.5.5 0 0 1 0-1zm0 2.5v.5H.5a.5.5 0 0 0 0 1h2a.5.5 0 0 0 0-1H2v-.5a.5.5 0 0 0-1 0"/>
        </svg>
        </strong>
        <br>
        <br>
        ${summary.escritos_count} escritos <br>
        ${summary.proyectos_count} proyectos
    `;

    // Render pie chart
    renderPieChart(summary);

    // Populate presentaciones by date table & chart
    if (summary.presentaciones_by_date) {
        populatePresentacionesByDateTable(summary.presentaciones_by_date);
    }

    // Populate top titles table and render bar chart
    populateTopTitlesTable(summary.most_titles);
}

function renderPieChart(summary) {
    const ctx = document.getElementById('pieChart').getContext('2d');
    if (pieChart) pieChart.destroy();
    pieChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['Proyectos', 'Escritos'],
            datasets: [{
                data: [summary.proyectos_count, summary.escritos_count],
                backgroundColor: ['#90CAF9', '#1360a4']
            }]
        },
        options: { maintainAspectRatio: false }
    });
}

function populateTopTitlesTable(mostTitles) {
    const tableBody = document.getElementById('top-titles-table');
    tableBody.innerHTML = '';
    const labels = [];
    const data = [];

    mostTitles.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `<td>${item.Título}</td><td>${item.Cantidad}</td>`;
        tableBody.appendChild(row);

        labels.push(item.Título);
        data.push(item.Cantidad);
    });

    renderBarChart(labels, data);
}

function renderBarChart(labels, data) {
    const ctx = document.getElementById('topTitlesBarChart').getContext('2d');
    if (barChart) barChart.destroy();
    barChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: ' ',
                data: data,
                backgroundColor: labels.map(() => `#${Math.floor(Math.random() * 16777215).toString(16)}`)
            }]
        },
        options: {
            maintainAspectRatio: false,
            // Remove indexAxis for vertical bars (default)
            scales: {
                x: { title: { display: true, text: 'Suma' } },
                y: { title: { display: true, text: 'Cantidad' }, beginAtZero: true }
            }
        }
    });
}

// New: populate presentaciones-by-date table and bar chart
function populatePresentacionesByDateTable(list) {
    const tableBody = document.getElementById('presentaciones-by-date-table');
    tableBody.innerHTML = '';
    const labels = [];
    const data = [];
    
    const dayNames = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    
    list.forEach(item => {
        // Parse date (dd/mm/yyyy) and get day name
        const parts = item.Fecha.split('/');
        const dateObj = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
        const dayName = dayNames[dateObj.getDay()];
        const displayDate = `${item.Fecha}<br><small style="font-size:0.75rem;">${dayName}</small>`;
        
        const row = document.createElement('tr');
        row.innerHTML = `<td style="font-size:1.1rem;">${displayDate}</td><td>${item.Escritos}</td><td>${item.Proyectos}</td><td>${item.Total}</td>`;
        tableBody.appendChild(row);
        
        labels.push(`${item.Fecha}\n${dayName}`);
        data.push(item.Total);
    });
    renderDatesBarChart(labels, data);
}

let datesBarChart;
function renderDatesBarChart(labels, data) {
    const ctx = document.getElementById('datesBarChart').getContext('2d');
    if (datesBarChart) datesBarChart.destroy();
    datesBarChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Total Presentaciones',
                data: data,
                backgroundColor: labels.map(() => `#${Math.floor(Math.random() * 16777215).toString(16)}`)
            }]
        },
        options: {
            maintainAspectRatio: false,
            scales: {
                x: { title: { display: true, text: 'Fecha' } },
                y: { title: { display: true, text: 'Total Presentaciones' }, beginAtZero: true }
            }
        }
    });
}

function animateWindowResize(targetWidth, targetHeight, duration = 400) {
    if (!window.pywebview) return;
    window.pywebview.api.get_window_size().then(size => {
        let startWidth = size.width;
        let startHeight = size.height;
        let startTime = null;

        function step(timestamp) {
            if (!startTime) startTime = timestamp;
            let progress = Math.min((timestamp - startTime) / duration, 1);
            let newWidth = Math.round(startWidth + (targetWidth - startWidth) * progress);
            let newHeight = Math.round(startHeight + (targetHeight - startHeight) * progress);
            window.pywebview.api.set_window_size(newWidth, newHeight);
            // Center window after resizing
            if (progress === 1) {
                // Get screen size and center
                window.pywebview.api.center_window(newWidth, newHeight);
            }
            if (progress < 1) {
                window.requestAnimationFrame(step);
            }
        }
        window.requestAnimationFrame(step);
    });
}

function showActions() {
    document.getElementById('exportBtn').classList.remove('d-none');
    document.getElementById('clearButton').classList.remove('d-none');
    document.getElementById('proveyentesSection').classList.remove('d-none');
    // Optionally animate
    document.getElementById('exportBtn').style.opacity = 0;
    document.getElementById('clearButton').style.opacity = 0;
    document.getElementById('proveyentesSection').style.opacity = 0;
    setTimeout(() => {
        document.getElementById('exportBtn').style.transition = "opacity 0.5s";
        document.getElementById('clearButton').style.transition = "opacity 0.5s";
        document.getElementById('proveyentesSection').style.transition = "opacity 0.5s";
        document.getElementById('exportBtn').style.opacity = 1;
        document.getElementById('clearButton').style.opacity = 1;
        document.getElementById('proveyentesSection').style.opacity = 1;
    }, 50);
}

function hideActions() {
    document.getElementById('exportBtn').classList.add('d-none');
    document.getElementById('clearButton').classList.add('d-none');
    document.getElementById('proveyentesSection').classList.add('d-none');
    document.getElementById('exportBtn').style.opacity = '';
    document.getElementById('clearButton').style.opacity = '';
    document.getElementById('proveyentesSection').style.opacity = '';
}

let popupCount = 0;

// Notification popout for produced files - stackable and with filename
function showDownloadPopup(filePath) {
    const filename = filePath.split(/[\\\/]/).pop(); // Extract filename from path
    const offsetY = popupCount * 100; // Stack vertically
    popupCount++;
    
    const container = document.createElement('div');
    container.style = `position:fixed;right:20px;bottom:${20 + offsetY}px;z-index:${20000 + popupCount};background:#0d6efd;color:white;padding:14px;border-radius:8px;box-shadow:0 6px 20px rgba(0,0,0,0.2);min-width:280px;`;
    container.innerHTML = `
        <div style="display:flex;flex-direction:column;gap:8px;">
            <div style="font-weight:600;">Archivo descargado</div>
            <div style="font-size:0.85rem;color:#e7f1ff;word-break:break-word;">${filename}</div>
            <div style="display:flex;gap:6px;">
                <button id="openProdBtn_${popupCount}" class="btn btn-sm btn-light">Abrir</button>
                <button id="closeProdBtn_${popupCount}" class="btn btn-sm btn-outline-light">Cerrar</button>
            </div>
        </div>
    `;
    document.body.appendChild(container);
    
    document.getElementById(`openProdBtn_${popupCount}`).onclick = function() {
        if (window.pywebview) {
            window.pywebview.api.open_file(filePath);
        }
    };
    
    document.getElementById(`closeProdBtn_${popupCount}`).onclick = function() {
        container.remove();
        popupCount--;
    };
    
    // Auto remove after 12s
    setTimeout(() => {
        if (document.body.contains(container)) {
            container.remove();
            popupCount--;
        }
    }, 12000);
}

// Handle select file
document.getElementById('selectFileBtn').onclick = async function() {
    if (window.pywebview) {
        const result = await window.pywebview.api.read_data('');
        if (result && result.status === 'ok') {
            document.getElementById('summary').style.display = 'block';
            const summary = parseSummary(result);
            renderDashboard(summary);
            showActions();
            animateWindowResize(1400, 1000, 400);
        } else {
            alert(result.message || 'No se pudo leer el archivo.');
        }
    } else {
        alert('pywebview API no disponible');
    }
};

// Excel export: show popup to open file
document.getElementById('exportBtn').onclick = async function() {
    if (window.pywebview) {
        const result = await window.pywebview.api.export_excel();
        if (result && result.status === 'ok') {
            showDownloadPopup(result.path);
        } else if (result && result.status === 'error') {
            alert(result.message);
        }
    } else {
        alert('pywebview API no disponible');
    }
};

// Export PDF repartidas and then show modal for continuous option
document.getElementById('exportPdfBtn').onclick = async function() {
    const nProveyentes = parseInt(document.getElementById('proveyentesInput').value);
    if (!nProveyentes || nProveyentes < 1) {
        alert('Ingrese un número válido de Proveyentes.');
        return;
    }
    if (window.pywebview) {
        const result = await window.pywebview.api.export_pdf(nProveyentes);
        if (result && result.status === 'ok') {
            showDownloadPopup(result.path);
            // Pre-fill modal date with last_assigned_date + 1 day if provided
            if (result.last_assigned_date) {
                // parse dd/mm/YYYY
                const parts = result.last_assigned_date.split('/');
                if (parts.length === 3) {
                    const d = parts[0].padStart(2,'0'), m = parts[1].padStart(2,'0'), y = parts[2];
                    const dt = new Date(`${y}-${m}-${d}`);
                    const next = new Date(dt.getTime() + 24*60*60*1000);
                    const iso = next.toISOString().slice(0,10);
                    document.getElementById('continuousStartDate').value = iso;
                }
            } else {
                // default to tomorrow
                const tomorrow = new Date(Date.now() + 24*60*60*1000).toISOString().slice(0,10);
                document.getElementById('continuousStartDate').value = tomorrow;
            }
            // Show modal
            const modalEl = document.getElementById('continuousExportModal');
            const modal = new mdb.Modal(modalEl);
            modal.show();
        } else {
            alert(result.message || 'Error al exportar PDF.');
        }
    } else {
        alert('pywebview API no disponible');
    }
};

// Hook modal export button
document.getElementById('doContinuousExportBtn').onclick = async function() {
    const startDate = document.getElementById('continuousStartDate').value;
    const nListados = parseInt(document.getElementById('continuousNumListados').value) || 1;
    if (!startDate) {
        alert('Ingrese una fecha de inicio válida.');
        return;
    }
    if (window.pywebview) {
        const result = await window.pywebview.api.export_pdf_continuous(startDate, nListados);
        if (result && result.status === 'ok') {
            showDownloadPopup(result.path);
            // hide modal
            const modalEl = document.getElementById('continuousExportModal');
            const modal = mdb.Modal.getInstance(modalEl);
            if (modal) modal.hide();
        } else {
            alert(result.message || 'Error al exportar PDF continuo.');
        }
    } else {
        alert('pywebview API no disponible');
    }
};

// Clear button
document.getElementById('clearButton').onclick = function() {
    document.getElementById('summary').style.display = 'none';
    document.getElementById('summary-text').textContent = '';
    document.getElementById('period-header').textContent = '';
    document.getElementById('top-titles-table').innerHTML = '';
    document.getElementById('presentaciones-by-date-table').innerHTML = '';
    hideActions();
    if (pieChart) pieChart.destroy();
    if (barChart) barChart.destroy();
    if (datesBarChart) datesBarChart.destroy();
    animateWindowResize(1200, 933, 400);
};

// Add pywebview API methods for window size and centering
if (window.pywebview) {
    window.pywebview.api.get_window_size = function() {
        return new Promise(resolve => {
            // Dummy fallback if not implemented in backend
            resolve({ width: window.innerWidth, height: window.innerHeight });
        });
    };
    window.pywebview.api.set_window_size = function(width, height) {
        // Dummy fallback if not implemented in backend
        window.resizeTo(width, height);
    };
    window.pywebview.api.center_window = function(width, height) {
        // Dummy fallback if not implemented in backend
        const screenW = window.screen.width;
        const screenH = window.screen.height;
        const x = Math.max(0, Math.round((screenW - width) / 2));
        const y = Math.max(0, Math.round((screenH - height) / 2));
        window.moveTo(x, y);
    };
    window.pywebview.api.open_file = window.pywebview.api.open_file || function(path) {
        return new Promise(resolve => {
            // Fallback: try to open in new tab if path is a file:// URL (best-effort)
            try {
                window.open("file://" + path);
            } catch (e) {}
            resolve({ status: 'ok' });
        });
    };
}
