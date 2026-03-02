// ============================================================================
// app.js - Enhanced version with Static HTML Export
// ============================================================================

let pieChart, barChart, datesBarChart;
let currentSummary = null; // Store current data for modal access
let rawData = null; // Store the full dataset for detailed queries
let currentCalendarDate = new Date();
let calendarDataCache = null; // Cache calendar data for better performance

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
        
        initializeDataMaps(info);
        buildCalendarCache(); // Build calendar cache when data loads
        
        return currentSummary;
    }

    // Fallback parsing (keeping for compatibility)
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
// Initialize data maps for quick lookups - Updated to store all records
// ============================================================================

function initializeDataMaps(info) {
    rawData = {
        byDate: new Map(),
        byTitle: new Map(),
        allRecords: []
    };
    
    // If we have the full dataset from the backend, use it
    if (info.raw_records) {
        rawData.allRecords = info.raw_records;
        
        // Build byDate and byTitle maps from raw records
        info.raw_records.forEach(record => {
            // Group by date
            if (!rawData.byDate.has(record.fecha)) {
                rawData.byDate.set(record.fecha, []);
            }
            rawData.byDate.get(record.fecha).push(record);
            
            // Group by title
            if (!rawData.byTitle.has(record.titulo)) {
                rawData.byTitle.set(record.titulo, []);
            }
            rawData.byTitle.get(record.titulo).push(record);
        });
    }
    
    // If we don't have raw records, use presentaciones_by_date as fallback
    if (info.presentaciones_by_date && rawData.allRecords.length === 0) {
        info.presentaciones_by_date.forEach(dayData => {
            rawData.byDate.set(dayData.Fecha, {
                escritos: dayData.Escritos,
                proyectos: dayData.Proyectos,
                total: dayData.Total,
                records: [] // Will be empty without full data
            });
        });
    }
    
    // Initialize title map from most_titles
    if (info.most_titles) {
        info.most_titles.forEach(titleData => {
            if (!rawData.byTitle.has(titleData.Título)) {
                rawData.byTitle.set(titleData.Título, {
                    cantidad: titleData.Cantidad,
                    records: []
                });
            }
        });
    }
}

// ============================================================================
// Build calendar cache for faster rendering
// ============================================================================

function buildCalendarCache() {
    if (!currentSummary || !currentSummary.presentaciones_by_date) return;
    
    calendarDataCache = new Map();
    currentSummary.presentaciones_by_date.forEach(item => {
        calendarDataCache.set(item.Fecha, {
            total: item.Total,
            escritos: item.Escritos,
            proyectos: item.Proyectos
        });
    });
}

// ============================================================================
// Fetch records by date from backend
// ============================================================================

async function fetchRecordsByDate(dateStr) {
    if (!window.pywebview) {
        console.warn('pywebview API not available');
        return [];
    }
    
    try {
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
// Fetch records by title from backend
// ============================================================================

async function fetchRecordsByTitle(title) {
    if (!window.pywebview) {
        console.warn('pywebview API not available');
        return [];
    }
    
    try {
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
    document.getElementById('period-header').innerHTML = `
        <div class="d-flex align-items-center gap-3">
            <div class="bg-primary bg-opacity-10 p-3 rounded-3">
                <i class="fas fa-calendar-alt fa-2x text-primary"></i>
            </div>
            <div>
                <div class="text-muted small">Período</div>
                <div class="fw-bold fs-5">${summary.period}</div>
                <div class="text-muted small mt-1">
                    <i class="fas fa-clock me-1"></i>${summary.days_difference} días corridos
                </div>
            </div>
        </div>
    `;

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

// ============================================================================
// Updated table population functions with chronological descending order
// ============================================================================

function populatePresentacionesByDateTable(list) {
    const tableBody = document.getElementById('presentaciones-by-date-table');
    tableBody.innerHTML = '';
    const labels = [];
    const data = [];
    const dayNames = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

    // Sort the list chronologically descending (most recent first)
    const sortedList = [...list].sort((a, b) => {
        // Parse dates in dd/mm/yyyy format
        const [aDay, aMonth, aYear] = a.Fecha.split('/').map(Number);
        const [bDay, bMonth, bYear] = b.Fecha.split('/').map(Number);
        
        // Create Date objects (Year, Month-1, Day) for comparison
        const dateA = new Date(aYear, aMonth - 1, aDay);
        const dateB = new Date(bYear, bMonth - 1, bDay);
        
        // Descending order (most recent first)
        return dateB - dateA;
    });

    sortedList.forEach(item => {
        const parts = item.Fecha.split('/');
        const dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
        const dayName = dayNames[dateObj.getDay()];

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>
                <div class="fw-semibold">${item.Fecha}</div>
                <small class="text-muted">${dayName}</small>
            </td>
            <td class="text-center align-middle">
                <span class="badge bg-primary-subtle text-primary">${item.Escritos}</span>
            </td>
            <td class="text-center align-middle">
                <span class="badge bg-success-subtle text-success">${item.Proyectos}</span>
            </td>
            <td class="text-center align-middle">
                <span class="fw-bold fs-6" style="color: #e67e22;">${item.Total}</span>
            </td>
            <td class="text-center align-middle">
                <button class="btn-view-details" onclick="showDateRecords('${item.Fecha}')">
                    <i class="fas fa-eye me-1"></i>Ver
                </button>
            </td>
        `;
        tableBody.appendChild(row);

        // Use the sorted order for chart labels as well
        labels.push(`${item.Fecha}\n${dayName}`);
        data.push(item.Total);
    });

    renderDatesBarChart(labels, data);
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
            <td class="text-center align-middle">
                <span class="badge bg-primary rounded-pill px-3 py-2">${item.Cantidad}</span>
            </td>
            <td class="text-center align-middle">
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
// ENHANCED CALENDAR OVERLAY - Beautiful and Detailed
// ============================================================================

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
    const escritosCounts = {};
    const proyectosCounts = {};
    
    if (currentSummary && currentSummary.presentaciones_by_date) {
        currentSummary.presentaciones_by_date.forEach(item => {
            recordCounts[item.Fecha] = item.Total;
            escritosCounts[item.Fecha] = item.Escritos;
            proyectosCounts[item.Fecha] = item.Proyectos;
        });
    }

    // Parse oldest record date for highlighting
    let oldestDate = null;
    if (currentSummary && currentSummary.oldest_escrito) {
        oldestDate = parseDateString(currentSummary.oldest_escrito);
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Build calendar HTML with enhanced styling
    let html = '<div class="calendar-grid">';

    // Day headers with improved styling
    const dayHeaders = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
    dayHeaders.forEach(day => {
        html += `<div class="calendar-day-header">${day.substring(0, 3)}</div>`;
    });

    // Empty cells for days before month start
    for (let i = 0; i < adjustedStartDay; i++) {
        html += '<div class="calendar-day empty"></div>';
    }

    // Fill in the days with enhanced information
    for (let day = 1; day <= totalDays; day++) {
        const currentDate = new Date(year, month, day);
        currentDate.setHours(0, 0, 0, 0);
        
        const dateStr = `${String(day).padStart(2, '0')}/${String(month + 1).padStart(2, '0')}/${year}`;
        const total = recordCounts[dateStr] || 0;
        const escritos = escritosCounts[dateStr] || 0;
        const proyectos = proyectosCounts[dateStr] || 0;

        let classes = 'calendar-day';
        if (oldestDate && currentDate.getTime() === oldestDate.getTime()) {
            classes += ' highlight-first';
        }
        if (currentDate.getTime() === today.getTime()) {
            classes += ' highlight-today';
        }
        if (total > 0) {
            classes += ' has-records';
        }

        // Make day clickable if it has records
        const clickHandler = total > 0 ? `onclick="showDateRecords('${dateStr}')"` : '';

        html += `<div class="${classes}" ${clickHandler} style="${total > 0 ? 'cursor: pointer;' : ''}">`;
        html += `<div class="calendar-day-number">${day}</div>`;
        
        if (total > 0) {
            html += `
                <div class="calendar-day-stats">
                    <div class="d-flex flex-column gap-1 mt-1">
                        <div class="d-flex align-items-center justify-content-between small">
                            <span><i class="fas fa-file-alt text-primary" style="font-size: 0.7rem;"></i> Escritos:</span>
                            <span class="fw-semibold text-primary">${escritos}</span>
                        </div>
                        <div class="d-flex align-items-center justify-content-between small">
                            <span><i class="fas fa-file text-success" style="font-size: 0.7rem;"></i> Proyectos:</span>
                            <span class="fw-semibold text-success">${proyectos}</span>
                        </div>
                        <div class="d-flex align-items-center justify-content-between small border-top mt-1 pt-1">
                            <span><i class="fas fa-calculator text-muted"></i> Total:</span>
                            <span class="fw-bold" style="color: #e67e22;">${total}</span>
                        </div>
                    </div>
                </div>
            `;
        } else {
            html += `<div class="calendar-day-empty text-muted small mt-2">Sin escritos</div>`;
        }
        
        html += '</div>';
    }

    html += '</div>';
    document.getElementById('calendar-container').innerHTML = html;
}

// ============================================================================
// MODAL FOR DATE RECORDS - Fixed z-index issue
// ============================================================================

async function showDateRecords(dateStr) {
    // First, hide the calendar modal
    const calendarModal = mdb.Modal.getInstance(document.getElementById('calendarModal'));
    if (calendarModal) {
        calendarModal.hide();
    }
    
    // Small delay to ensure calendar modal is hidden
    setTimeout(async () => {
        const modal = new mdb.Modal(document.getElementById('dateRecordsModal'));
        document.getElementById('selected-date-display').textContent = dateStr;

        const tableBody = document.querySelector('#date-records-table tbody');
        tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4"><i class="fas fa-spinner fa-spin me-2"></i>Cargando registros...</td></tr>';

        try {
            let records = [];
            
            if (window.pywebview) {
                const result = await window.pywebview.api.get_records_by_date(dateStr);
                if (result && result.status === 'ok') {
                    records = result.records;
                }
            } else {
                records = generateMockRecordsForDate(dateStr);
            }

            tableBody.innerHTML = '';
            
            if (records.length === 0) {
                tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-muted">No hay registros para esta fecha</td></tr>';
            } else {
                records.forEach(record => {
                    const row = document.createElement('tr');
                    const tipo = record.tipo || record.Tipo || 'N/A';
                    const isEscrito = tipo.toLowerCase().includes('escrito');
                    
                    row.innerHTML = `
                        <td>${record.expte || record.Expte || 'N/A'}</td>
                        <td>${record.titulo || record.Título || 'N/A'}</td>
                        <td><span class="badge ${isEscrito ? 'bg-primary' : 'bg-success'}">${tipo}</span></td>
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
    }, 300);
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
            const result = await window.pywebview.api.get_records_by_title(title);
            if (result && result.status === 'ok') {
                records = result.records;
            }
        } else {
            records = generateMockRecordsForTitle(title);
        }

        tableBody.innerHTML = '';
        
        if (records.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-muted">No hay registros con este título</td></tr>';
        } else {
            records.forEach(record => {
                const row = document.createElement('tr');
                const tipo = record.tipo || record.Tipo || 'N/A';
                const isEscrito = tipo.toLowerCase().includes('escrito');
                
                row.innerHTML = `
                    <td>${record.expte || record.Expte || 'N/A'}</td>
                    <td>${record.fecha || record.Fecha || record.Recibido || 'N/A'}</td>
                    <td><span class="badge ${isEscrito ? 'bg-primary' : 'bg-success'}">${tipo}</span></td>
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
// EXPORT STATIC HTML DASHBOARD
// ============================================================================

// ============================================================================
// EXPORT STATIC HTML DASHBOARD - Updated with full records
// ============================================================================

async function exportStaticDashboard() {
    if (!currentSummary) {
        alert('No hay datos cargados para exportar.');
        return;
    }

    try {
        // Show loading state
        const exportBtn = document.getElementById('exportStaticBtn');
        const originalText = exportBtn ? exportBtn.innerHTML : '';
        if (exportBtn) {
            exportBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>Exportando...';
            exportBtn.disabled = true;
        }

        // Fetch all records from backend if available
        let allRecords = [];
        if (window.pywebview && window.pywebview.api && typeof window.pywebview.api.get_all_records === 'function') {
            try {
                const result = await window.pywebview.api.get_all_records();
                if (result && result.status === 'ok') {
                    allRecords = result.records;
                }
            } catch (error) {
                console.warn('Could not fetch all records, using summary data only', error);
            }
        }

        // If we couldn't get all records, try to build them from available data
        if (allRecords.length === 0 && rawData && rawData.allRecords) {
            allRecords = rawData.allRecords;
        }

        // Prepare the data to be embedded with complete records
        const exportData = {
            summary: currentSummary,
            records: allRecords, // Include all records for modal functionality
            charts: {
                pieChart: pieChart ? pieChart.toBase64Image() : null,
                datesBarChart: datesBarChart ? datesBarChart.toBase64Image() : null,
                topTitlesBarChart: barChart ? barChart.toBase64Image() : null
            },
            timestamp: new Date().toLocaleString('es-ES', {
                year: 'numeric',
                month: 'long',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit'
            }),
            dataSource: 'Exportado desde Listado de Escritos'
        };

        // Generate the HTML content
        const htmlContent = generateStaticHTML(exportData);

        // Call backend to save the file
        if (window.pywebview && window.pywebview.api) {
            if (typeof window.pywebview.api.export_static_html === 'function') {
                const result = await window.pywebview.api.export_static_html({
                    ...exportData,
                    html_content: htmlContent
                });
                
                if (result && result.status === 'ok') {
                    showDownloadPopup(result.path, 'HTML');
                } else {
                    alert(result?.message || 'Error al exportar el dashboard.');
                }
            } else {
                // Fallback if method doesn't exist
                console.warn('export_static_html method not found, using fallback download');
                downloadStaticDashboard(htmlContent);
            }
        } else {
            // Fallback for web environment - download directly
            downloadStaticDashboard(htmlContent);
        }
    } catch (error) {
        console.error('Error exporting static dashboard:', error);
        alert('Error al exportar el dashboard: ' + error.message);
    } finally {
        const exportBtn = document.getElementById('exportStaticBtn');
        if (exportBtn) {
            exportBtn.innerHTML = '<i class="fas fa-file-code me-2"></i>Exportar Dashboard HTML';
            exportBtn.disabled = false;
        }
    }
}

function downloadStaticDashboard(htmlContent) {
    // Create blob and download
    const blob = new Blob([htmlContent], { type: 'text/html' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `dashboard_${formatDateForFilename(new Date())}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    // Show success message
    alert('Dashboard exportado correctamente como archivo HTML.');
}

function formatDateForFilename(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${year}${month}${day}_${hours}${minutes}`;
}

function downloadStaticDashboard(htmlContent) {
    // Create blob and download
    const blob = new Blob([htmlContent], { type: 'text/html' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `dashboard_${formatDateForFilename(new Date())}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    // Show success message
    alert('Dashboard exportado correctamente como archivo HTML.');
}

function formatDateForFilename(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${year}${month}${day}_${hours}${minutes}`;
}

function generateStaticHTML(data) {
    const dayNames = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    const summary = data.summary;
    const monthNames = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                        'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
    
    // Sort presentaciones by date descending for display
    const sortedPresentaciones = [...summary.presentaciones_by_date].sort((a, b) => {
        const [aDay, aMonth, aYear] = a.Fecha.split('/').map(Number);
        const [bDay, bMonth, bYear] = b.Fecha.split('/').map(Number);
        return new Date(bYear, bMonth - 1, bDay) - new Date(aYear, aMonth - 1, aDay);
    });

    // Get the current date for calendar
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth();

    return `<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard - Listado de Escritos (Exportado)</title>
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css">
    <!-- MDB UI Kit CSS -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/mdb-ui-kit/6.3.0/mdb.min.css">
    <!-- Font Awesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        body {
            background-color: #f8fafc;
            color: #1e293b;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            padding: 2rem;
        }
        .container {
            max-width: 1600px;
            margin: 0 auto;
        }
        h2 {
            font-weight: 600;
            font-size: 2rem;
            color: #0f172a;
            margin-bottom: 1.5rem;
        }
        h4 {
            font-weight: 600;
            font-size: 1.25rem;
            color: #334155;
            margin-bottom: 1.25rem;
            border-left: 4px solid #0d6efd;
            padding-left: 1rem;
        }
        .header-info {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 2rem;
            padding: 1rem;
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
            border: 1px solid #eef2f6;
        }
        .timestamp {
            color: #64748b;
            font-size: 0.95rem;
        }
        .timestamp i {
            margin-right: 0.5rem;
            color: #0d6efd;
        }
        .summary-stats {
            background: white;
            border-radius: 16px;
            padding: 1.5rem;
            margin-bottom: 2rem;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
            border: 1px solid #eef2f6;
        }
        .stat-item {
            padding: 1rem;
            border-radius: 12px;
            background: #f8fafc;
            text-align: center;
        }
        .stat-number {
            font-size: 2rem;
            font-weight: 700;
            color: #0d6efd;
            line-height: 1.2;
        }
        .stat-label {
            color: #64748b;
            font-size: 0.9rem;
            font-weight: 500;
        }
        .period-info {
            display: flex;
            align-items: center;
            gap: 1rem;
        }
        .period-icon {
            background: rgba(13, 110, 253, 0.1);
            padding: 1rem;
            border-radius: 12px;
        }
        .period-icon i {
            font-size: 2rem;
            color: #0d6efd;
        }
        .period-text {
            font-size: 1.1rem;
        }
        .period-days {
            color: #64748b;
            font-size: 0.9rem;
            margin-top: 0.25rem;
        }
        .table-responsive {
            border-radius: 12px;
            background: white;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
            border: 1px solid #eef2f6;
            margin-bottom: 1rem;
            overflow: visible;
        }
        .table {
            margin-bottom: 0;
            width: 100%;
        }
        .table thead th {
            background: #f8fafc;
            color: #475569;
            font-weight: 600;
            font-size: 0.9rem;
            text-transform: uppercase;
            padding: 1rem 0.75rem;
            border-bottom: 2px solid #e2e8f0;
        }
        .table tbody td {
            padding: 1rem 0.75rem;
            vertical-align: middle;
            border-bottom: 1px solid #eef2f6;
        }
        .table tbody tr {
            height: 70px;
        }
        .badge-count {
            display: inline-block;
            padding: 0.5rem 0.8rem;
            border-radius: 20px;
            font-weight: 500;
            font-size: 0.85rem;
        }
        .badge-escritos {
            background-color: #e7f1ff;
            color: #0d6efd;
        }
        .badge-proyectos {
            background-color: #e8f5e9;
            color: #198754;
        }
        .badge-total {
            background-color: #e67e22;
            color: white;
            padding: 0.5rem 1rem;
            border-radius: 30px;
            font-weight: 600;
        }
        .chart-container {
            background: white;
            border-radius: 16px;
            padding: 1.5rem;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
            border: 1px solid #eef2f6;
            margin-bottom: 1rem;
            min-height: 350px;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .chart-image {
            max-width: 100%;
            max-height: 300px;
            object-fit: contain;
        }
        .chart-placeholder {
            color: #94a3b8;
            font-style: italic;
            text-align: center;
            padding: 2rem;
        }
        .footer-note {
            margin-top: 3rem;
            padding: 1rem;
            text-align: center;
            color: #64748b;
            font-size: 0.9rem;
            border-top: 1px solid #eef2f6;
        }
        
        /* ===== MODAL STYLES ===== */
        .modal-content {
            border: none;
            border-radius: 16px;
            box-shadow: 0 30px 60px rgba(0,0,0,0.2);
        }
        .modal-header {
            background: linear-gradient(135deg, #f8fafc, #ffffff);
            border-bottom: 2px solid #eef2f6;
            padding: 1.5rem 1.5rem 1rem;
            border-radius: 16px 16px 0 0;
        }
        .modal-header .modal-title {
            font-weight: 600;
            color: #0f172a;
            font-size: 1.25rem;
        }
        .modal-body {
            padding: 1.5rem;
            max-height: 70vh;
            overflow-y: auto;
        }
        .modal-body::-webkit-scrollbar {
            width: 8px;
        }
        .modal-body::-webkit-scrollbar-track {
            background: #f1f1f1;
            border-radius: 8px;
        }
        .modal-body::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 8px;
        }
        .modal-body::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
        .modal-footer {
            border-top: 2px solid #eef2f6;
            padding: 1rem 1.5rem 1.5rem;
        }
        
        /* ===== CALENDAR STYLES ===== */
        .calendar-grid {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 8px;
            margin-top: 20px;
        }
        .calendar-day-header {
            text-align: center;
            font-weight: 600;
            color: #64748b;
            font-size: 0.85rem;
            padding: 10px;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }
        .calendar-day {
            background: white;
            border: 1px solid #eef2f6;
            border-radius: 12px;
            padding: 12px 8px;
            min-height: 120px;
            transition: all 0.2s ease;
            position: relative;
            cursor: pointer;
        }
        .calendar-day:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(0,0,0,0.05);
            border-color: #0d6efd;
        }
        .calendar-day-number {
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 8px;
            font-size: 1rem;
        }
        .calendar-day.empty {
            background: #f8fafc;
            border-color: #eef2f6;
            cursor: default;
        }
        .calendar-day.empty:hover {
            transform: none;
            box-shadow: none;
            border-color: #eef2f6;
        }
        .calendar-day.has-records {
            background: linear-gradient(135deg, #ffffff, #f8fafc);
        }
        .calendar-day.highlight-first {
            background: linear-gradient(135deg, #fff3e0, #ffe4bc);
            border-color: #ffb74d;
            border-width: 2px;
        }
        .calendar-day.highlight-first .calendar-day-number {
            color: #ed6c02;
            font-weight: 700;
        }
        .calendar-day.highlight-today {
            background: linear-gradient(135deg, #e3f2fd, #bbdefb);
            border-color: #1976d2;
            border-width: 2px;
        }
        .calendar-day.highlight-today .calendar-day-number {
            color: #0d47a1;
            font-weight: 700;
        }
        .calendar-day-stats {
            font-size: 0.75rem;
            margin-top: 4px;
            padding: 6px 4px;
            background: rgba(0,0,0,0.02);
            border-radius: 8px;
        }
        .calendar-day-empty {
            font-size: 0.7rem;
            margin-top: 8px;
            padding: 4px;
            text-align: center;
            color: #94a3b8;
            font-style: italic;
        }
        .calendar-month-title {
            font-size: 1.5rem;
            font-weight: 600;
            color: #0f172a;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
        }
        .calendar-nav-btn {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 30px;
            padding: 8px 20px;
            color: #475569;
            transition: all 0.2s ease;
            font-size: 0.9rem;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .calendar-nav-btn:hover {
            background: #0d6efd;
            color: white;
            border-color: #0d6efd;
        }
        
        /* ===== VIEW DETAILS BUTTON ===== */
        .btn-view-details {
            padding: 0.5rem 1rem;
            font-size: 0.85rem;
            border-radius: 8px;
            background: #e9ecef;
            color: #495057;
            border: 1px solid #dee2e6;
            transition: all 0.2s ease;
            white-space: nowrap;
        }
        .btn-view-details:hover {
            background: #0d6efd;
            color: white;
            border-color: #0d6efd;
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(13,110,253,0.2);
        }
        .btn-view-details i {
            font-size: 0.8rem;
            margin-right: 0.25rem;
        }
        
        /* Stats badges */
        .badge.bg-primary-subtle {
            background-color: #e7f1ff;
            color: #0d6efd;
            font-weight: 500;
            padding: 0.5rem 0.8rem;
            border-radius: 20px;
            font-size: 0.85rem;
        }
        
        .badge.bg-success-subtle {
            background-color: #e8f5e9;
            color: #198754;
            font-weight: 500;
            padding: 0.5rem 0.8rem;
            border-radius: 20px;
            font-size: 0.85rem;
        }
        
        @media print {
            body { padding: 0; }
            .header-info { break-inside: avoid; }
            .summary-stats { break-inside: avoid; }
            .table-responsive { break-inside: auto; }
            .btn-view-details, .calendar-nav-btn { display: none; }
            .calendar-day { break-inside: avoid; }
        }
        
        /* Animation for calendar days */
        @keyframes fadeInDay {
            from {
                opacity: 0;
                transform: translateY(10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        .calendar-day {
            animation: fadeInDay 0.3s ease-out;
            animation-fill-mode: both;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Header with timestamp -->
        <div class="header-info">
            <div class="period-info">
                <div class="period-icon">
                    <i class="fas fa-file-signature"></i>
                </div>
                <div>
                    <h2 style="margin-bottom: 0;">Listado de Escritos</h2>
                    <div class="period-days">Dashboard exportado</div>
                </div>
            </div>
            <div class="timestamp">
                <i class="fas fa-clock"></i>
                ${data.timestamp}
            </div>
        </div>

        <!-- Period and Stats -->
        <div class="summary-stats">
            <div class="row align-items-center">
                <div class="col-md-6">
                    <div class="period-info">
                        <div class="period-icon">
                            <i class="fas fa-calendar-alt"></i>
                        </div>
                        <div>
                            <div class="text-muted small">Período analizado</div>
                            <div class="fw-bold fs-5">${summary.period || 'No disponible'}</div>
                            <div class="period-days">
                                <i class="fas fa-clock me-1"></i>${summary.days_difference || 0} días corridos
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="row g-4">
                        <div class="col-6">
                            <div class="stat-item">
                                <div class="stat-number">${summary.total_records}</div>
                                <div class="stat-label">Presentaciones</div>
                            </div>
                        </div>
                        <div class="col-6">
                            <div class="stat-item">
                                <div class="stat-number">${summary.unique_exptes_count}</div>
                                <div class="stat-label">Expedientes</div>
                            </div>
                        </div>
                        <div class="col-6">
                            <div class="stat-item">
                                <div class="stat-number">${summary.escritos_count}</div>
                                <div class="stat-label">Escritos</div>
                            </div>
                        </div>
                        <div class="col-6">
                            <div class="stat-item">
                                <div class="stat-number">${summary.proyectos_count}</div>
                                <div class="stat-label">Proyectos</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Calendar Section with Navigation -->
        <div class="row mt-4">
            <div class="col-12">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h4><i class="fas fa-calendar-alt me-2"></i>Calendario de presentaciones</h4>
                    <div class="d-flex gap-3">
                        <span class="badge" style="background: #fff3e0; color: #ed6c02; border: 1px solid #ffb74d; padding: 0.5rem 1rem;">
                            <i class="fas fa-star me-1"></i>Primer presentación
                        </span>
                        <span class="badge" style="background: #e3f2fd; color: #0d47a1; border: 1px solid #1976d2; padding: 0.5rem 1rem;">
                            <i class="fas fa-calendar-check me-1"></i>Hoy
                        </span>
                    </div>
                </div>
                <div class="chart-container flex-column" style="min-height: auto;">
                    <div class="calendar-month-title">
                        <button class="calendar-nav-btn" id="prevMonthBtn">
                            <i class="fas fa-chevron-left"></i> Mes anterior
                        </button>
                        <span id="currentMonthDisplay">${monthNames[currentMonth]} ${currentYear}</span>
                        <button class="calendar-nav-btn" id="nextMonthBtn">
                            Mes siguiente <i class="fas fa-chevron-right"></i>
                        </button>
                    </div>
                    <div id="static-calendar"></div>
                </div>
            </div>
        </div>

        <!-- Tables Side by Side -->
        <div class="row mt-5">
            <!-- Presentaciones por fecha Table -->
            <div class="col-lg-6">
                <h4><i class="fas fa-calendar-check me-2"></i>Presentaciones por fecha</h4>
                <div class="table-responsive">
                    <table class="table table-hover">
                        <thead>
                            <tr>
                                <th>Fecha</th>
                                <th class="text-center">Escritos</th>
                                <th class="text-center">Proyectos</th>
                                <th class="text-center">Total</th>
                                <th class="text-center"> </th>
                            </tr>
                        </thead>
                        <tbody>
                            ${sortedPresentaciones.map(item => {
                                const parts = item.Fecha.split('/');
                                const dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
                                const dayName = dayNames[dateObj.getDay()];
                                return `
                                <tr>
                                    <td>
                                        <div class="fw-semibold">${item.Fecha}</div>
                                        <small class="text-muted">${dayName}</small>
                                    </td>
                                    <td class="text-center align-middle">
                                        <span class="badge-count badge-escritos">${item.Escritos}</span>
                                    </td>
                                    <td class="text-center align-middle">
                                        <span class="badge-count badge-proyectos">${item.Proyectos}</span>
                                    </td>
                                    <td class="text-center align-middle">
                                        <span class="badge-total">${item.Total}</span>
                                    </td>
                                    <td class="text-center align-middle">
                                        <button class="btn-view-details" onclick="showDateRecords('${item.Fecha}')">
                                            <i class="fas fa-eye me-1"></i>Ver
                                        </button>
                                    </td>
                                </tr>
                                `;
                            }).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
            
            <!-- Escritos más frecuentes Table -->
            <div class="col-lg-6">
                <h4><i class="fas fa-chart-bar me-2"></i>Escritos más frecuentes</h4>
                <div class="table-responsive">
                    <table class="table table-hover">
                        <thead>
                            <tr>
                                <th>Título</th>
                                <th class="text-center">Cantidad</th>
                                <th class="text-center"> </th>
                            </tr>
                        </thead>
                        <tbody>
                            ${summary.most_titles.map(item => `
                                <tr>
                                    <td class="fw-medium">${item.Título}</td>
                                    <td class="text-center align-middle">
                                        <span class="badge" style="background: #0d6efd; color: white; padding: 0.5rem 1rem; border-radius: 30px;">${item.Cantidad}</span>
                                    </td>
                                    <td class="text-center align-middle">
                                        <button class="btn-view-details" onclick="showTitleRecords('${item.Título.replace(/'/g, "\\'")}')">
                                            <i class="fas fa-eye me-1"></i>Ver
                                        </button>
                                    </td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Charts Below Tables -->
        <div class="row mt-4">
            <!-- Presentaciones por fecha Chart -->
            <div class="col-lg-6">
                <div class="chart-container">
                    ${data.charts.datesBarChart ? 
                        `<img src="${data.charts.datesBarChart}" alt="Gráfico de presentaciones por fecha" class="chart-image">` : 
                        '<div class="chart-placeholder"><i class="fas fa-chart-bar fa-3x mb-3" style="color: #cbd5e1;"></i><br>Gráfico no disponible</div>'
                    }
                </div>
            </div>
            
            <!-- Escritos más frecuentes Chart -->
            <div class="col-lg-6">
                <div class="chart-container">
                    ${data.charts.topTitlesBarChart ? 
                        `<img src="${data.charts.topTitlesBarChart}" alt="Gráfico de escritos más frecuentes" class="chart-image">` : 
                        '<div class="chart-placeholder"><i class="fas fa-chart-pie fa-3x mb-3" style="color: #cbd5e1;"></i><br>Gráfico no disponible</div>'
                    }
                </div>
            </div>
        </div>

        <!-- Pie Chart -->
        <div class="row mt-4">
            <div class="col-12">
                <div class="chart-container" style="min-height: 350px;">
                    ${data.charts.pieChart ? 
                        `<img src="${data.charts.pieChart}" alt="Gráfico de distribución" style="max-height: 300px; object-fit: contain;">` : 
                        '<div class="chart-placeholder"><i class="fas fa-chart-pie fa-3x mb-3" style="color: #cbd5e1;"></i><br>Gráfico no disponible</div>'
                    }
                </div>
            </div>
        </div>

        <!-- Footer with data source -->
        <div class="footer-note">
            <i class="fas fa-database me-2"></i>
            ${data.dataSource} · ${data.timestamp}
        </div>
    </div>

    <!-- Modal for Date Records -->
    <div class="modal fade" id="dateRecordsModal" tabindex="-1" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">
                        <i class="fas fa-calendar-day me-2"></i>Presentaciones del <span id="selected-date-display"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="table-responsive">
                        <table class="table table-sm table-hover" id="date-records-table">
                            <thead>
                                <tr>
                                    <th>Expediente</th>
                                    <th>Título</th>
                                    <th>Tipo</th>
                                    <th>Presentante</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">
                        <i class="fas fa-times me-2"></i>Cerrar
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Modal for Title Records -->
    <div class="modal fade" id="titleRecordsModal" tabindex="-1" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">
                        <i class="fas fa-file-alt me-2"></i>Presentaciones con título: <span id="selected-title-display"></span>
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="table-responsive">
                        <table class="table table-sm table-hover" id="title-records-table">
                            <thead>
                                <tr>
                                    <th>Expediente</th>
                                    <th>Fecha</th>
                                    <th>Tipo</th>
                                    <th>Presentante</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">
                        <i class="fas fa-times me-2"></i>Cerrar
                    </button>
                </div>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS and MDB -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mdb-ui-kit/6.3.0/mdb.min.js"></script>
    
    <script>
        // ============================================================================
        // EMBEDDED DATA - Complete records for full functionality
        // ============================================================================
        
        const embeddedData = ${JSON.stringify({
            summary: data.summary,
            records: data.records || [], // Complete records data
            presentaciones_by_date: data.summary.presentaciones_by_date,
            most_titles: data.summary.most_titles,
            oldest_escrito: data.summary.oldest_escrito,
            today_date: data.summary.today_date
        }, null, 2)};
        
        // Build lookup maps for quick access
        const recordsByDate = new Map();
        const recordsByTitle = new Map();
        
        // Initialize maps with embedded records
        if (embeddedData.records && embeddedData.records.length > 0) {
            embeddedData.records.forEach(record => {
                // Group by date
                if (!recordsByDate.has(record.fecha)) {
                    recordsByDate.set(record.fecha, []);
                }
                recordsByDate.get(record.fecha).push(record);
                
                // Group by title
                if (!recordsByTitle.has(record.titulo)) {
                    recordsByTitle.set(record.titulo, []);
                }
                recordsByTitle.get(record.titulo).push(record);
            });
        }
        
        // Calendar state
        let currentCalendarDate = new Date(${currentYear}, ${currentMonth}, 1);
        const monthNames = ${JSON.stringify(monthNames)};
        const dayNames = ${JSON.stringify(dayNames)};
        
        // ============================================================================
        // CALENDAR FUNCTIONS
        // ============================================================================
        
        function renderCalendar(date) {
            const year = date.getFullYear();
            const month = date.getMonth();

            // Update month display
            document.getElementById('currentMonthDisplay').textContent = \`\${monthNames[month]} \${year}\`;

            // Get first day of month and total days
            const firstDay = new Date(year, month, 1);
            const lastDay = new Date(year, month + 1, 0);
            const startingDay = firstDay.getDay();
            const totalDays = lastDay.getDate();

            // Adjust for Monday as first day of week
            const adjustedStartDay = startingDay === 0 ? 6 : startingDay - 1;

            // Parse oldest record date for highlighting
            let oldestDate = null;
            if (embeddedData.oldest_escrito) {
                const parts = embeddedData.oldest_escrito.split('/');
                oldestDate = new Date(parts[2], parts[1] - 1, parts[0]);
            }

            const today = new Date();
            today.setHours(0, 0, 0, 0);

            // Build record counts per date
            const recordCounts = {};
            const escritosCounts = {};
            const proyectosCounts = {};
            
            if (embeddedData.presentaciones_by_date) {
                embeddedData.presentaciones_by_date.forEach(item => {
                    recordCounts[item.Fecha] = item.Total;
                    escritosCounts[item.Fecha] = item.Escritos;
                    proyectosCounts[item.Fecha] = item.Proyectos;
                });
            }

            // Build calendar HTML
            let html = '<div class="calendar-grid">';

            // Day headers
            const dayHeaders = ['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'];
            dayHeaders.forEach(day => {
                html += '<div class="calendar-day-header">' + day + '</div>';
            });

            // Empty cells for days before month start
            for (let i = 0; i < adjustedStartDay; i++) {
                html += '<div class="calendar-day empty"></div>';
            }

            // Fill in the days
            for (let day = 1; day <= totalDays; day++) {
                const currentDate = new Date(year, month, day);
                currentDate.setHours(0, 0, 0, 0);
                
                const dateStr = String(day).padStart(2, '0') + '/' + String(month + 1).padStart(2, '0') + '/' + year;
                const total = recordCounts[dateStr] || 0;
                const escritos = escritosCounts[dateStr] || 0;
                const proyectos = proyectosCounts[dateStr] || 0;

                let classes = 'calendar-day';
                if (oldestDate && currentDate.getTime() === oldestDate.getTime()) {
                    classes += ' highlight-first';
                }
                if (currentDate.getTime() === today.getTime()) {
                    classes += ' highlight-today';
                }
                if (total > 0) {
                    classes += ' has-records';
                }

                // Make day clickable if it has records
                const clickHandler = total > 0 ? \`onclick="showDateRecords('\${dateStr}')\"\` : '';

                html += '<div class="' + classes + '" ' + clickHandler + '>';
                html += '<div class="calendar-day-number">' + day + '</div>';
                
                if (total > 0) {
                    html += \`
                        <div class="calendar-day-stats">
                            <div class="d-flex flex-column gap-1 mt-1">
                                <div class="d-flex align-items-center justify-content-between small">
                                    <span><i class="fas fa-file-alt text-primary" style="font-size: 0.7rem;"></i> Escritos:</span>
                                    <span class="fw-semibold text-primary">\${escritos}</span>
                                </div>
                                <div class="d-flex align-items-center justify-content-between small">
                                    <span><i class="fas fa-file text-success" style="font-size: 0.7rem;"></i> Proyectos:</span>
                                    <span class="fw-semibold text-success">\${proyectos}</span>
                                </div>
                                <div class="d-flex align-items-center justify-content-between small border-top mt-1 pt-1">
                                    <span><i class="fas fa-calculator text-muted"></i> Total:</span>
                                    <span class="fw-bold" style="color: #e67e22;">\${total}</span>
                                </div>
                            </div>
                        </div>
                    \`;
                } else {
                    html += '<div class="calendar-day-empty text-muted small mt-2">Sin escritos</div>';
                }
                
                html += '</div>';
            }

            html += '</div>';
            document.getElementById('static-calendar').innerHTML = html;
        }

        // ============================================================================
        // MODAL FUNCTIONS
        // ============================================================================
        
        function showDateRecords(dateStr) {
            const modal = new bootstrap.Modal(document.getElementById('dateRecordsModal'));
            document.getElementById('selected-date-display').textContent = dateStr;

            const tableBody = document.querySelector('#date-records-table tbody');
            tableBody.innerHTML = '';
            
            // Get records for this date
            const records = recordsByDate.get(dateStr) || [];
            
            if (records.length === 0) {
                tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-muted">No hay registros para esta fecha</td></tr>';
            } else {
                records.forEach(record => {
                    const row = document.createElement('tr');
                    const tipo = record.tipo || 'N/A';
                    const isEscrito = tipo.toLowerCase().includes('escrito');
                    
                    row.innerHTML = \`
                        <td>\${record.expte || 'N/A'}</td>
                        <td>\${record.titulo || 'N/A'}</td>
                        <td><span class="badge \${isEscrito ? 'bg-primary' : 'bg-success'}">\${tipo}</span></td>
                        <td>\${record.presentante || 'N/A'}</td>
                    \`;
                    tableBody.appendChild(row);
                });
            }

            modal.show();
        }

        function showTitleRecords(title) {
            const modal = new bootstrap.Modal(document.getElementById('titleRecordsModal'));
            document.getElementById('selected-title-display').textContent = title;

            const tableBody = document.querySelector('#title-records-table tbody');
            tableBody.innerHTML = '';
            
            // Get records with this title
            const records = recordsByTitle.get(title) || [];
            
            if (records.length === 0) {
                tableBody.innerHTML = '<tr><td colspan="4" class="text-center py-4 text-muted">No hay registros con este título</td></tr>';
            } else {
                records.forEach(record => {
                    const row = document.createElement('tr');
                    const tipo = record.tipo || 'N/A';
                    const isEscrito = tipo.toLowerCase().includes('escrito');
                    
                    row.innerHTML = \`
                        <td>\${record.expte || 'N/A'}</td>
                        <td>\${record.fecha || 'N/A'}</td>
                        <td><span class="badge \${isEscrito ? 'bg-primary' : 'bg-success'}">\${tipo}</span></td>
                        <td>\${record.presentante || 'N/A'}</td>
                    \`;
                    tableBody.appendChild(row);
                });
            }

            modal.show();
        }

        // ============================================================================
        // INITIALIZATION AND EVENT LISTENERS
        // ============================================================================
        
        document.addEventListener('DOMContentLoaded', function() {
            renderCalendar(currentCalendarDate);
            
            // Calendar navigation
            document.getElementById('prevMonthBtn').addEventListener('click', function() {
                currentCalendarDate.setMonth(currentCalendarDate.getMonth() - 1);
                renderCalendar(currentCalendarDate);
            });
            
            document.getElementById('nextMonthBtn').addEventListener('click', function() {
                currentCalendarDate.setMonth(currentCalendarDate.getMonth() + 1);
                renderCalendar(currentCalendarDate);
            });
        });

        // Make functions globally available
        window.showDateRecords = showDateRecords;
        window.showTitleRecords = showTitleRecords;
    </script>
</body>
</html>`;
}

// ============================================================================
// MOCK DATA GENERATORS (for development/fallback)
// ============================================================================

function generateMockRecordsForDate(dateStr) {
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
    
    // Add export static button if it exists
    const exportStaticBtn = document.getElementById('exportStaticBtn');
    if (exportStaticBtn) {
        exportStaticBtn.classList.remove('d-none');
    }

    const elements = [
        document.getElementById('exportBtn'),
        document.getElementById('calendarBtn'),
        document.getElementById('clearButton'),
        document.getElementById('proveyentesSection')
    ];
    
    if (exportStaticBtn) elements.push(exportStaticBtn);
    
    elements.forEach(el => {
        if (el) {
            el.style.opacity = 0;
            setTimeout(() => {
                el.style.transition = "opacity 0.5s";
                el.style.opacity = 1;
            }, 50);
        }
    });
}

function hideActions() {
    document.getElementById('exportBtn').classList.add('d-none');
    document.getElementById('calendarBtn').classList.add('d-none');
    document.getElementById('clearButton').classList.add('d-none');
    document.getElementById('proveyentesSection').classList.add('d-none');
    
    const exportStaticBtn = document.getElementById('exportStaticBtn');
    if (exportStaticBtn) {
        exportStaticBtn.classList.add('d-none');
    }

    const elements = [
        document.getElementById('exportBtn'),
        document.getElementById('calendarBtn'),
        document.getElementById('clearButton'),
        document.getElementById('proveyentesSection'),
        exportStaticBtn
    ];
    
    elements.forEach(el => {
        if (el) el.style.opacity = '';
    });
}

// ============================================================================
// NOTIFICATION POPUP
// ============================================================================

let popupCount = 0;

function showDownloadPopup(filePath, fileType = 'archivo') {
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
                <span style="font-weight:600;">${fileType} exportado</span>
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
        container.remove();
        popupCount--;
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
// CREATE EXPORT STATIC BUTTON ON LOAD
// ============================================================================

function createExportStaticButton() {
    if (!document.getElementById('exportStaticBtn')) {
        const mainActions = document.getElementById('mainActions');
        if (mainActions) {
            const exportStaticBtn = document.createElement('button');
            exportStaticBtn.type = 'button';
            exportStaticBtn.className = 'btn btn-primary d-none';
            exportStaticBtn.id = 'exportStaticBtn';
            exportStaticBtn.style.background = 'linear-gradient(135deg, #6f42c1, #6610f2)';
            exportStaticBtn.style.border = 'none';
            exportStaticBtn.innerHTML = '<i class="fas fa-file-code me-2"></i>Exportar Dashboard HTML';
            exportStaticBtn.onclick = exportStaticDashboard;
            mainActions.appendChild(exportStaticBtn);
            console.log('Export static button created');
        }
    }
}

// Create button when DOM is loaded
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', createExportStaticButton);
} else {
    createExportStaticButton();
}

// ============================================================================
// EVENT LISTENERS
// ============================================================================

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

document.getElementById('calendarBtn').onclick = function() {
    openCalendar();
};

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

document.getElementById('exportBtn').onclick = async function() {
    if (window.pywebview) {
        const result = await window.pywebview.api.export_excel();
        if (result && result.status === 'ok') {
            showDownloadPopup(result.path, 'Excel');
        } else if (result && result.status === 'error') {
            alert(result.message);
        }
    } else {
        alert('pywebview API no disponible');
    }
};

document.getElementById('exportPdfBtn').onclick = async function() {
    const nProveyentes = parseInt(document.getElementById('proveyentesInput').value);
    if (!nProveyentes || nProveyentes < 1) {
        alert('Ingrese un número válido de Proveyentes.');
        return;
    }

    if (window.pywebview) {
        const result = await window.pywebview.api.export_pdf(nProveyentes);
        if (result && result.status === 'ok') {
            showDownloadPopup(result.path, 'PDF');
            
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
            showDownloadPopup(result.path, 'PDF');
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
    calendarDataCache = null;
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
window.exportStaticDashboard = exportStaticDashboard;