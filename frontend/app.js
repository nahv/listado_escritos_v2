let pieChart, barChart;

function formatDate(dateString) {
    const date = new Date(dateString);
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function parseSummary(info) {
    // Parse info string from backend to summary object
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
</svg> Período ${summary.period} <br><br>
<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" fill="currentColor" class="bi bi-calendar2-event mb-2" viewBox="0 0 16 16">
  <path d="M11 7.5a.5.5 0 0 1 .5-.5h1a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5h-1a.5.5 0 0 1-.5-.5z"/>
  <path d="M3.5 0a.5.5 0 0 1 .5.5V1h8V.5a.5.5 0 0 1 1 0V1h1a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h1V.5a.5.5 0 0 1 .5-.5M2 2a1 1 0 0 0-1 1v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V3a1 1 0 0 0-1-1z"/>
  <path d="M2.5 4a.5.5 0 0 1 .5-.5h10a.5.5 0 0 1 .5.5v1a.5.5 0 0 1-.5.5H3a.5.5 0 0 1-.5-.5z"/>
</svg>
<strong>${summary.days_difference} días corridos</strong>
<br> del primer escrito (${summary.oldest_escrito})
<br> a la fecha (${summary.today_date})
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

document.getElementById('selectFileBtn').onclick = async function() {
    if (window.pywebview) {
        const result = await window.pywebview.api.read_data('');
        document.getElementById('summary').style.display = 'block';
        const summary = parseSummary(result);
        renderDashboard(summary);
        showActions();
        animateWindowResize(1400, 1000, 400);
    } else {
        alert('pywebview API no disponible');
    }
};

document.getElementById('exportBtn').onclick = async function() {
    if (window.pywebview) {
        const result = await window.pywebview.api.export_excel();
        // Optionally show export result somewhere, e.g. as a toast or alert
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
        await window.pywebview.api.export_pdf(nProveyentes);
        // No alert after export
    } else {
        alert('pywebview API no disponible');
    }
};

document.getElementById('clearButton').onclick = function() {
    document.getElementById('summary').style.display = 'none';
    document.getElementById('summary-text').textContent = '';
    document.getElementById('period-header').textContent = '';
    document.getElementById('top-titles-table').innerHTML = '';
    hideActions();
    if (pieChart) pieChart.destroy();
    if (barChart) barChart.destroy();
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
}
