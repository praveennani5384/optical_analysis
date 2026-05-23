let rawData = [];
let charts = { lineChart: null, barChart: null };

// 1. Load Excel Data
document.getElementById('upload').addEventListener('change', (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        
        populateFilters(rawData);
        alert("File loaded successfully! Adjust filters and click 'Generate Analytics'.");
    };
    reader.readAsArrayBuffer(file);
});

// 2. Dynamic Dropdowns
function populateFilters(data) {
    const filters = ['Category', 'Brand', 'Year', 'Month'];
    filters.forEach(key => {
        const select = document.getElementById(`filter${key}`);
        const uniqueValues = [...new Set(data.map(item => item[key]))].sort();
        select.innerHTML = `<option value="">All ${key}s</option>` + 
            uniqueValues.map(v => `<option value="${v}">${v}</option>`).join('');
    });
}

// 3. Filter & Process
document.getElementById('processBtn').addEventListener('click', () => {
    if (rawData.length === 0) return alert("Please upload an Excel file first!");

    const fCat = document.getElementById('filterCategory').value;
    const fBrand = document.getElementById('filterBrand').value;
    const fYear = document.getElementById('filterYear').value;
    const fMonth = document.getElementById('filterMonth').value;

    const filtered = rawData.filter(d => 
        (!fCat || String(d.Category) === fCat) &&
        (!fBrand || String(d.Brand) === fBrand) &&
        (!fYear || String(d.Year) === fYear) &&
        (!fMonth || String(d.Month) === fMonth)
    );

    if (filtered.length === 0) return alert("No data found for the selected filters!");
    updateDashboard(filtered);
});

// 4. Update Dashboard
function updateDashboard(data) {
    // Labels combine Month and Brand for clarity
    const labels = data.map(d => `${d.Month || ''} - ${d.Brand || ''}`);
    const frames = data.map(d => d.Frames || 0);
    const lenses = data.map(d => d.Lenses || 0);

    // Render both charts
    renderChart('lineChart', 'line', labels, frames, lenses, false);
    renderChart('barChart', 'bar', labels, frames, lenses, true);
}

// 5. Chart.js Core Function
function renderChart(id, type, labels, d1, d2, isStacked) {
    if (charts[id]) charts[id].destroy(); // Clear old chart instance

    const ctx = document.getElementById(id).getContext('2d');
    charts[id] = new Chart(ctx, {
        type: type,
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Frames',
                    data: d1,
                    backgroundColor: 'rgba(26, 115, 232, 0.7)',
                    borderColor: '#1a73e8',
                    borderWidth: 2,
                    fill: !isStacked,
                    tension: 0.3
                },
                {
                    label: 'Lenses',
                    data: d2,
                    backgroundColor: 'rgba(52, 168, 83, 0.7)',
                    borderColor: '#34a853',
                    borderWidth: 2,
                    fill: !isStacked,
                    tension: 0.3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: isStacked },
                y: { stacked: isStacked, beginAtZero: true }
            },
            plugins: {
                legend: { position: 'top' }
            }
        }
    });
}