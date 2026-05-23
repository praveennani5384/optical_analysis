let rawData = [];
let charts = { lineChart: null, barChart: null };

// 1. Listen for File Upload and Parse Excel
document.getElementById('upload').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        
        populateFilters(rawData);
        alert("Excel File Loaded! You can now filter and generate charts.");
    };
    reader.readAsArrayBuffer(file);
});

// 2. Populate Dropdowns dynamically based on unique values in the Excel
function populateFilters(data) {
    const filterKeys = ['Category', 'Brand', 'Year', 'Month'];
    filterKeys.forEach(key => {
        const select = document.getElementById(`filter${key}`);
        const uniqueValues = [...new Set(data.map(item => item[key]))].sort();
        
        select.innerHTML = `<option value="">All ${key}s</option>` + 
            uniqueValues.map(val => `<option value="${val}">${val}</option>`).join('');
    });
}

// 3. Filter data and update charts when button is clicked
document.getElementById('processBtn').addEventListener('click', () => {
    if (rawData.length === 0) return alert("Please upload an Excel file first!");

    // Capture current filter values
    const fCat = document.getElementById('filterCategory').value;
    const fBrand = document.getElementById('filterBrand').value;
    const fYear = document.getElementById('filterYear').value;
    const fMonth = document.getElementById('filterMonth').value;

    // Multi-column filtering logic
    const filteredData = rawData.filter(item => {
        return (!fCat || String(item.Category) === fCat) &&
               (!fBrand || String(item.Brand) === fBrand) &&
               (!fYear || String(item.Year) === fYear) &&
               (!fMonth || String(item.Month) === fMonth);
    });

    if (filteredData.length === 0) return alert("No data matches these filters!");

    // Prep data for Chart.js
    const labels = filteredData.map(d => `${d.Month || ''} ${d.Brand || ''} (${d.Category || ''})`);
    const frames = filteredData.map(d => d.Frames || 0);
    const lenses = filteredData.map(d => d.Lenses || 0);

    // Update both chart visuals
    renderChart('lineChart', 'line', labels, frames, lenses, false);
    renderChart('barChart', 'bar', labels, frames, lenses, true);
});

// 4. Core Chart Rendering Function
function renderChart(id, type, labels, dataset1, dataset2, isStacked) {
    // If chart already exists, destroy it before creating a new one to avoid overlapping
    if (charts[id]) {
        charts[id].destroy();
    }

    const ctx = document.getElementById(id).getContext('2d');
    charts[id] = new Chart(ctx, {
        type: type,
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Frames',
                    data: dataset1,
                    backgroundColor: 'rgba(26, 115, 232, 0.7)',
                    borderColor: '#1a73e8',
                    borderWidth: 2,
                    fill: !isStacked, // Line chart looks better without fill if multiple lines exist
                    tension: 0.3
                },
                {
                    label: 'Lenses',
                    data: dataset2,
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
                x: { 
                    stacked: isStacked,
                    ticks: { font: { size: 10 } } // Smaller font for mobile labels
                },
                y: { 
                    stacked: isStacked, 
                    beginAtZero: true 
                }
            },
            plugins: {
                legend: { position: 'top' }
            }
        }
    });
}
