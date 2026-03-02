// ============================================================================
// app.js - Enhanced version with REAL DATA integration
// ============================================================================

let pieChart, barChart, datesBarChart;
let currentSummary = null; // Store current data for modal access
let rawData = null; // Store the full dataset for detailed queries

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function formatDate(dateString) {
    const date = new Date(dateString);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function parseDateString(dateStr) {
    // Parse dd/mm/yyyy
    const parts = dateStr.split('/');
    if (parts.length === 3) {
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return null;
}

// ============================================================================
// DATA PARSING - Now stores raw data for detailed queries
// ============================================================================

function parseSummary(info) {
    if (typeof info === 'object') {
        // Store the full dataset for detailed queries
        // The backend needs to provide the full data array
        // For now, we'll simulate with presentaciones_by_date as the source
        
        currentSummary = {
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
        
        // Initialize rawData structure for queries
        // In a real implementation, the backend should provide the full dataset
        // For now, we'll create a map for quick lookups
        initializeDataMaps(info);
        
        return currentSummary;
    }

    // Fallback parsing for string format (keeping for compatibility)
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

    const tableStart = info.indexOf('Escritos más repetidos:');
    if (tableStart !== -1) {
        const tableText = info.substring(tableStart).split('\n').slice(1);
        tableText.forEach((row, idx) => {
            if (idx === 0) return;
            const parts = row.trim().split(/\s{2,}/);
            if (parts.length === 2 && parts[0] && parts[1] && parts[0] !== 'Suma' && !isNaN(parseInt(parts[1]))) {
                most_titles.push({ Título: parts[0], Cantidad: parseInt(parts[1]) });
            }
        });
    }

    const period = oldest_escrito && today_date ? `${oldest_escrito} a ${today_date}` : '';

    currentSummary = {
        total_records, unique_exptes_count, escritos_count, proyectos_count,
        oldest_escrito, days_difference, today_date, transferencias_count,
        most_titles, period, presentaciones_by_date: []
    };

    return currentSummary;
}

// ============================================================================
// NEW FUNCTION: Initialize data maps for quick lookups
// ============================================================================

function initializeDataMaps(info) {
    // Create a map for quick access to records by date
    // In a real implementation, this would come from the backend
    // For now, we'll create a structure that can be populated when we have the data
    
    rawData = {
        byDate: new Map(),
        byTitle: new Map(),
        allRecords: []
    };
    
    // If we have presentaciones_by_date, we can use it as a starting point
    if (info.presentaciones_by_date) {
        info.presentaciones_by_date.forEach(dayData => {
            rawData.byDate.set(dayData.Fecha, {
                escritos: dayData.Escritos,
                proyectos: dayData.Proyectos,
                total: dayData.Total,
                records: [] // Will be populated when we have full data
            });
        });
    }
    
    // If we have most_titles, initialize title map
    if (info.most_titles) {
        info.most_titles.forEach(titleData => {
            rawData.byTitle.set(titleData.Título, {
                cantidad: titleData.Cantidad,
                records: [] // Will be populated when we have full data
            });
        });
    }
}

// ============================================================================
// NEW FUNCTION: Fetch records by date from backend
// ============================================================================

async function fetchRecordsByDate(dateStr) {
    // In a real implementation, this would call a backend API
    // For now, we'll simulate with the data we have
    
    if (!window.pywebview) {
        console.warn('pywebview API not available');
        return [];
    }
    
    try {
        // This would be a new API method you'd need to implement in backend.py
        // For now, we'll simulate with a call to get filtered data
        const result = await window.pywebview.api.get_records_by_date(dateStr);
        if (result && result.status === 'ok') {
            return result.records;
        }
    } catch (error) {
        console.error('Error fetching records by date:', error);
    }
    
    return [];
}

// ============================================================================
// NEW FUNCTION: Fetch records by title from backend
// ============================================================================

async function fetchRecordsByTitle(title) {
    if (!window.pywebview) {
        console.warn('pywebview API not available');
        return [];
    }
    
    try {
        // This would be a new API method you'd need to implement in backend.py
        const result = await window.pywebview.api.get_records_by_title(title);
        if (result && result.status === 'ok') {
            return result.records;
        }
    } catch (error) {
        console.error('Error fetching records by title:', error);
    }
    
    return [];
}

// ============================================================================
// RENDER DASHBOARD
// ============================================================================

function renderDashboard(summary) {
    // Update period header
    document.getElementById('period-header').innerHTML = `
        <div class="d-flex align-items-center gap-3">
            <div class="bg-primary bg-opacity-10 p-3 rounded-3">
                <i class="fas fa-calendar-alt fa-2x text-primary"></i>
            </div>
            <div>
                <div class="text-muted small">Período analizado</div>
                <div class="fw-bold fs-5">${summary.period}</div>
                <div class="text-muted small mt-1">
                    <i class="fas fa-clock me-1"></i>${summary.days_difference} días corridos
                </div>
            </div>
        </div>
    `;

    // Update summary text with statistics
    document.getElementById('summary-text').innerHTML = `
        <div class="row g-4">
            <div class="col-6 col-md-3">
                <div class="stat-item text-center">
                    <div class="stat-number">${summary.total_records}</div>
                    <div class="stat-label">Presentaciones</div>
                </div>
            </div>
            <div class="col-6 col-md-3">
                <div class="stat-item text-center">
                    <div class="stat-number">${summary.unique_exptes_count}</div>
                    <div class="stat-label">Expedientes</div>
                </div>
            </div>
            <div class="col-6 col-md-3">
                <div class="stat-item text-center">
                    <div class="stat-number">${summary.escritos_count}</div>
                    <div class="stat-label">Escritos</div>
                </div>
            </div>
            <div class="col-6 col-md-3">
                <div class="stat-item text-center">
                    <div class="stat-number">${summary.proyectos_count}</div>
                    <div class="stat-label">Proyectos</div>
                </div>
            </div>
        </div>
    `;

    renderPieChart(summary);

    if (summary.presentaciones_by_date) {
        populatePresentacionesByDateTable(summary.presentaciones_by_date);
    }

    populateTopTitlesTable(summary.most_titles);
}

// ============================================================================
// CHARTS
// ============================================================================

function renderPieChart(summary) {
    const ctx = document.getElementById('pieChart').getContext('2d');
    if (pieChart) pieChart.destroy();

    pieChart = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: ['Proyectos', 'Escritos'],
            datasets: [{
                data: [summary.proyectos_count, summary.escritos_count],
                backgroundColor: ['#2c3e50', '#e67e22'],
                borderWidth: 0
            }]
        },
        options: {
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        font: { size: 12, weight: '500' },
                        padding: 20
                    }
                }
            }
        }
    });
}

function populateTopTitlesTable(mostTitles) {
    const tableBody = document.getElementById('top-titles-table');
    tableBody.innerHTML = '';
    const labels = [];
    const data = [];

    mostTitles.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="fw-medium">${item.Título}</td>
            <td class="text-center"><span class="badge bg-primary rounded-pill">${item.Cantidad}</span></td>
            <td class="text-center">
                <button class="btn-view-details" onclick="showTitleRecords('${item.Título.replace(/'/g, "\\'")}')">
                    <i class="fas fa-eye me-1"></i>Ver
                </button>
            </td>
        `;
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
                label: 'Cantidad',
                data: data,
                backgroundColor: '#e67e22',
                borderRadius: 6
            }]
        },
        options: {
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: '#eef2f6' }
                },
                x: {
                    grid: { display: false }
                }
            }
        }
    });
}

function populatePresentacionesByDateTable(list) {
    const tableBody = document.getElementById('presentaciones-by-date-table');
    tableBody.innerHTML = '';
    const labels = [];
    const data = [];
    const dayNames = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

    list.forEach(item => {
        const parts = item.Fecha.split('/');
        const dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
        const dayName = dayNames[dateObj.getDay()];

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <div class="fw-semibold">${item.Fecha}</div>
                <small class="text-muted">${dayName}</small>
            </td>
            <td class="text-center">${item.Escritos}</td>
            <td class="text-center">${item.Proyectos}</td>
            <td class="text-center"><span class="fw-bold">${item.Total}</span></td>
            <td class="text-center">
                <button class="btn-view-details" onclick="showDateRecords('${item.Fecha}')">
                    <i class="fas fa-eye me-1"></i>Ver
                </button>
            </td>
        `;
        tableBody.appendChild(row);

        labels.push(`${item.Fecha}\n${dayName}`);
        data.push(item.Total);
    });

    renderDatesBarChart(labels, data);
}

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
                backgroundColor: '#3498db',
                borderRadius: 6
            }]
        },
        options: {
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: '#eef2f6' }
                },
                x: {
                    grid: { display: false }
                }
            }
        }
    });
}

// ============================================================================
// MODAL FOR DATE RECORDS - With REAL DATA
// ============================================================================

async function showDateRecords(dateStr) {
    const modal = new mdb.Modal(document.getElementById('dateRecordsModal'));
    document.getElementById('selected-date-display').textContent = dateStr;

    const tableBody = document.querySelector('#date-records-table tbody');
    tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4"><i class="fas fa-spinner fa-spin me-2"></i>Cargando registros...</td></tr>';

    try {
        // In a real implementation, fetch from backend
        let records = [];
        
        if (window.pywebview) {
            // Call backend to get records for this date
            const result = await window.pywebview.api.get_records_by_date(dateStr);
            if (result && result.status === 'ok') {
                records = result.records;
            }
        } else {
            // Fallback simulation using the summary data
            // In development, we'll use the presentaciones_by_date info
            records = generateMockRecordsForDate(dateStr);
        }

        tableBody.innerHTML = '';
        
        if (records.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-muted">No hay registros para esta fecha</td></tr>';
        } else {
            records.forEach(record => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${record.expte || record.Expte || 'N/A'}</td>
                    <td>${record.titulo || record.Título || 'N/A'}</td>
                    <td><span class="badge ${(record.tipo || record.Tipo || '').toLowerCase().includes('escrito') ? 'bg-primary' : 'bg-success'}">${record.tipo || record.Tipo || 'N/A'}</span></td>
                    <td>${record.presentante || record.Apellido || record.Presentante || 'N/A'}</td>
                `;
                tableBody.appendChild(row);
            });
        }
    } catch (error) {
        console.error('Error loading date records:', error);
        tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-danger">Error al cargar registros</td></tr>';
    }

    modal.show();
}

// ============================================================================
// MODAL FOR TITLE RECORDS - With REAL DATA
// ============================================================================

async function showTitleRecords(title) {
    const modal = new mdb.Modal(document.getElementById('titleRecordsModal'));
    document.getElementById('selected-title-display').textContent = title;

    const tableBody = document.querySelector('#title-records-table tbody');
    tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4"><i class="fas fa-spinner fa-spin me-2"></i>Cargando registros...</td></tr>';

    try {
        let records = [];
        
        if (window.pywebview) {
            // Call backend to get records with this title
            const result = await window.pywebview.api.get_records_by_title(title);
            if (result && result.status === 'ok') {
                records = result.records;
            }
        } else {
            // Fallback simulation
            records = generateMockRecordsForTitle(title);
        }

        tableBody.innerHTML = '';
        
        if (records.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-muted">No hay registros con este título</td></tr>';
        } else {
            records.forEach(record => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${record.expte || record.Expte || 'N/A'}</td>
                    <td>${record.fecha || record.Fecha || record.Recibido || 'N/A'}</td>
                    <td><span class="badge ${(record.tipo || record.Tipo || '').toLowerCase().includes('escrito') ? 'bg-primary' : 'bg-success'}">${record.tipo || record.Tipo || 'N/A'}</span></td>
                    <td>${record.presentante || record.Apellido || record.Presentante || 'N/A'}</td>
                `;
                tableBody.appendChild(row);
            });
        }
    } catch (error) {
        console.error('Error loading title records:', error);
        tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-danger">Error al cargar registros</td></tr>';
    }

    modal.show();
}

// ============================================================================
// CALENDAR OVERLAY - With REAL DATA
// ============================================================================

let currentCalendarDate = new Date();

function openCalendar() {
    const modal = new mdb.Modal(document.getElementById('calendarModal'));
    renderCalendar(currentCalendarDate);
    modal.show();
}

function renderCalendar(date) {
    const year = date.getFullYear();
    const month = date.getMonth();

    // Update month display
    const monthNames = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                        'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
    document.getElementById('currentMonthDisplay').textContent = `${monthNames[month]} ${year}`;

    // Get first day of month and total days
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const startingDay = firstDay.getDay(); // 0 = Sunday
    const totalDays = lastDay.getDate();

    // Adjust for Monday as first day of week (Spanish locale)
    const adjustedStartDay = startingDay === 0 ? 6 : startingDay - 1;

    // Get record counts per date from actual data
    const recordCounts = {};
    if (currentSummary && currentSummary.presentaciones_by_date) {
        currentSummary.presentaciones_by_date.forEach(item => {
            recordCounts[item.Fecha] = item.Total;
        });
    }

    // Parse oldest record date for highlighting
    let oldestDate = null;
    if (currentSummary && currentSummary.oldest_escrito) {
        oldestDate = parseDateString(currentSummary.oldest_escrito);
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Build calendar HTML
    let html = '<div class="calendar-grid">';

    // Day headers
    const dayHeaders = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];
    dayHeaders.forEach(day => {
        html += `<div class="calendar-day-header">${day}</div>`;
    });

    // Empty cells for days before month start
    for (let i = 0; i < adjustedStartDay; i++) {
        html += '<div class="calendar-day empty"></div>';
    }

    // Fill in the days
    for (let day = 1; day <= totalDays; day++) {
        const currentDate = new Date(year, month, day);
        currentDate.setHours(0, 0, 0, 0);
        
        const dateStr = `${String(day).padStart(2, '0')}/${String(month + 1).padStart(2, '0')}/${year}`;
        const count = recordCounts[dateStr] || 0;

        let classes = 'calendar-day';
        if (oldestDate && currentDate.getTime() === oldestDate.getTime()) {
            classes += ' highlight-first';
        }
        if (currentDate.getTime() === today.getTime()) {
            classes += ' highlight-today';
        }

        // Make day clickable if it has records
        const clickHandler = count > 0 ? `onclick="showDateRecords('${dateStr}')"` : '';

        html += `<div class="${classes}" ${clickHandler} style="${count > 0 ? 'cursor: pointer;' : ''}">`;
        html += `<div class="calendar-day-number">${day}</div>`;
        if (count > 0) {
            html += `<div class="record-count-badge">${count} ${count === 1 ? 'registro' : 'registros'}</div>`;
        }
        html += '</div>';
    }

    html += '</div>';
    document.getElementById('calendar-container').innerHTML = html;
}

// ============================================================================
// MOCK DATA GENERATORS (for development/fallback)
// ============================================================================

function generateMockRecordsForDate(dateStr) {
    // This is only used when pywebview is not available
    // In production, this data would come from the backend
    const tipos = ['Escrito', 'Proyecto'];
    const presentantes = ['Dr. Pérez', 'Dra. García', 'Dr. Rodríguez', 'Dra. Martínez', 'Dr. López'];
    const titulos = ['Demanda', 'Contestación', 'Proyecto de sentencia', 'Recurso', 'Traslado'];
    
    const count = Math.floor(Math.random() * 5) + 1;
    const records = [];
    
    for (let i = 0; i < count; i++) {
        const tipo = tipos[Math.floor(Math.random() * tipos.length)];
        records.push({
            expte: `${Math.floor(Math.random() * 90000) + 10000}`,
            titulo: titulos[Math.floor(Math.random() * titulos.length)],
            tipo: tipo,
            presentante: presentantes[Math.floor(Math.random() * presentantes.length)]
        });
    }
    
    return records;
}

function generateMockRecordsForTitle(title) {
    // Mock data for development
    const tipos = ['Escrito', 'Proyecto'];
    const presentantes = ['Dr. Pérez', 'Dra. García', 'Dr. Rodríguez'];
    const fechas = ['15/01/2024', '16/01/2024', '17/01/2024', '18/01/2024'];
    
    const count = Math.floor(Math.random() * 3) + 1;
    const records = [];
    
    for (let i = 0; i < count; i++) {
        records.push({
            expte: `${Math.floor(Math.random() * 90000) + 10000}`,
            fecha: fechas[Math.floor(Math.random() * fechas.length)],
            tipo: 'Escrito',
            presentante: presentantes[Math.floor(Math.random() * presentantes.length)]
        });
    }
    
    return records;
}

// ============================================================================
// WINDOW MANAGEMENT
// ============================================================================

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

            if (progress === 1) {
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
    document.getElementById('calendarBtn').classList.remove('d-none');
    document.getElementById('clearButton').classList.remove('d-none');
    document.getElementById('proveyentesSection').classList.remove('d-none');

    // Smooth fade in
    [document.getElementById('exportBtn'), document.getElementById('calendarBtn'),
     document.getElementById('clearButton'), document.getElementById('proveyentesSection')].forEach(el => {
        el.style.opacity = 0;
        setTimeout(() => {
            el.style.transition = "opacity 0.5s";
            el.style.opacity = 1;
        }, 50);
    });
}

function hideActions() {
    document.getElementById('exportBtn').classList.add('d-none');
    document.getElementById('calendarBtn').classList.add('d-none');
    document.getElementById('clearButton').classList.add('d-none');
    document.getElementById('proveyentesSection').classList.add('d-none');

    [document.getElementById('exportBtn'), document.getElementById('calendarBtn'),
     document.getElementById('clearButton'), document.getElementById('proveyentesSection')].forEach(el => {
        el.style.opacity = '';
    });
}

// ============================================================================
// NOTIFICATION POPUP
// ============================================================================

let popupCount = 0;

function showDownloadPopup(filePath) {
    const filename = filePath.split(/[\\\/]/).pop();
    const offsetY = popupCount * 100;
    popupCount++;

    const container = document.createElement('div');
    container.style = `
        position: fixed;
        right: 20px;
        bottom: ${20 + offsetY}px;
        z-index: ${20000 + popupCount};
        background: white;
        color: #1e293b;
        padding: 16px;
        border-radius: 12px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.15);
        min-width: 300px;
        border-left: 4px solid #0d6efd;
    `;

    container.innerHTML = `
        <div style="display:flex;flex-direction:column;gap:10px;">
            <div style="display:flex;align-items:center;gap:8px;">
                <i class="fas fa-check-circle text-success"></i>
                <span style="font-weight:600;">Archivo descargado</span>
            </div>
            <div style="font-size:0.85rem;color:#64748b;word-break:break-word;">
                <i class="fas fa-file me-2"></i>${filename}
            </div>
            <div style="display:flex;gap:8px;margin-top:4px;">
                <button id="openProdBtn_${popupCount}" class="btn btn-sm btn-primary">
                    <i class="fas fa-folder-open me-1"></i>Abrir
                </button>
                <button id="closeProdBtn_${popupCount}" class="btn btn-sm btn-outline-secondary">
                    <i class="fas fa-times me-1"></i>Cerrar
                </button>
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

    setTimeout(() => {
        if (document.body.contains(container)) {
            container.remove();
            popupCount--;
        }
    }, 12000);
}

// ============================================================================
// EVENT LISTENERS
// ============================================================================

// Select file button
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

// Calendar button
document.getElementById('calendarBtn').onclick = function() {
    openCalendar();
};

// Previous month button
document.addEventListener('click', function(e) {
    if (e.target.id === 'prevMonthBtn' || e.target.closest('#prevMonthBtn')) {
        currentCalendarDate.setMonth(currentCalendarDate.getMonth() - 1);
        renderCalendar(currentCalendarDate);
    }
    if (e.target.id === 'nextMonthBtn' || e.target.closest('#nextMonthBtn')) {
        currentCalendarDate.setMonth(currentCalendarDate.getMonth() + 1);
        renderCalendar(currentCalendarDate);
    }
});

// Excel export
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

// Export PDF
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
            
            if (result.last_assigned_date) {
                const parts = result.last_assigned_date.split('/');
                if (parts.length === 3) {
                    const d = parts[0].padStart(2,'0'), m = parts[1].padStart(2,'0'), y = parts[2];
                    const dt = new Date(`${y}-${m}-${d}`);
                    const next = new Date(dt.getTime() + 24*60*60*1000);
                    const iso = next.toISOString().slice(0,10);
                    document.getElementById('continuousStartDate').value = iso;
                }
            } else {
                const tomorrow = new Date(Date.now() + 24*60*60*1000).toISOString().slice(0,10);
                document.getElementById('continuousStartDate').value = tomorrow;
            }
            
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

// Continuous export
document.getElementById('doContinuousExportBtn').onclick = async function() {
    const startDate = document.getElementById('continuousStartDate').value;
    const nProveyentes = parseInt(document.getElementById('continuousNumListados').value) || 1;

    if (!startDate) {
        alert('Ingrese una fecha de inicio válida.');
        return;
    }

    if (window.pywebview) {
        const result = await window.pywebview.api.export_pdf_continuous(startDate, nProveyentes);
        if (result && result.status === 'ok') {
            showDownloadPopup(result.path);
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
    document.getElementById('summary-text').innerHTML = '';
    document.getElementById('period-header').innerHTML = '';
    document.getElementById('top-titles-table').innerHTML = '';
    document.getElementById('presentaciones-by-date-table').innerHTML = '';
    document.getElementById('calendar-container').innerHTML = '';
    
    hideActions();
    
    if (pieChart) pieChart.destroy();
    if (barChart) barChart.destroy();
    if (datesBarChart) datesBarChart.destroy();
    
    currentSummary = null;
    rawData = null;
    animateWindowResize(1200, 933, 400);
};

// ============================================================================
// PYWEBVIEW API FALLBACKS
// ============================================================================

if (window.pywebview) {
    window.pywebview.api.get_window_size = function() {
        return new Promise(resolve => {
            resolve({ width: window.innerWidth, height: window.innerHeight });
        });
    };
    
    window.pywebview.api.set_window_size = function(width, height) {
        window.resizeTo(width, height);
    };
    
    window.pywebview.api.center_window = function(width, height) {
        const screenW = window.screen.width;
        const screenH = window.screen.height;
        const x = Math.max(0, Math.round((screenW - width) / 2));
        const y = Math.max(0, Math.round((screenH - height) / 2));
        window.moveTo(x, y);
    };
    
    window.pywebview.api.open_file = window.pywebview.api.open_file || function(path) {
        return new Promise(resolve => {
            try {
                window.open("file://" + path);
            } catch (e) {}
            resolve({ status: 'ok' });
        });
    };
}

// Make functions globally available for onclick handlers
window.showDateRecords = showDateRecords;
window.showTitleRecords = showTitleRecords;