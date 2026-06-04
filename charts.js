// Build the chart tab for the currently selected sheet.
function renderCharts(sheetName, data) {
    const chartsContainer = document.getElementById('chartsContainer');
    chartsContainer.innerHTML = '';
    
    // 銷毀舊圖表
    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};
    
    currentSheetData = data;
    currentSheetName = sheetName;
    buildChartConfigPanel(data);
    refreshChart();
}

function loadCompareSheetData() {
    // The comparison chart reads the three report sheets directly by Accession.
    sheetDataMap = {};
    compareSheets.forEach(sheetName => {
        if (currentWorkbook.SheetNames.includes(sheetName)) {
            const sheet = currentWorkbook.Sheets[sheetName];
            sheetDataMap[sheetName] = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        }
    });
}

function initializeComparePanel() {
    const panel = document.getElementById('comparePanel');
    const select = document.getElementById('compareAccessionSelect');
    const yAxisSelect = document.getElementById('compareYAxisSelect');
    select.innerHTML = '';
    yAxisSelect.innerHTML = '';

    // Merge available accessions across MAD, MADEV, and MADEVNP.
    const allAccessions = new Set();
    Object.values(sheetDataMap).forEach(data => {
        data.forEach(row => {
            if (row.Accession) {
                allAccessions.add(String(row.Accession).trim());
            }
        });
    });

    const sortedAccessions = Array.from(allAccessions).sort();
    if (sortedAccessions.length === 0) {
        panel.style.display = 'none';
        return;
    }

    panel.style.display = 'block';
    sortedAccessions.forEach(acc => {
        const option = document.createElement('option');
        option.value = acc;
        option.textContent = acc;
        select.appendChild(option);
    });

    // Populate Y-axis options with numeric fields
    const firstSheetData = sheetDataMap[compareSheets[0]];
    if (firstSheetData && firstSheetData.length > 0) {
        const columns = Object.keys(firstSheetData[0]);
        const numericFields = columns.filter(col => isNumericColumn(firstSheetData, col));
        numericFields.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            yAxisSelect.appendChild(option);
        });
        // Set default to 'Relative abundance (%)' if available
        if (columns.includes('Relative abundance (%)')) {
            yAxisSelect.value = 'Relative abundance (%)';
        }
    }

    updateComparisonChart(sortedAccessions[0]);
}

function updateComparisonChart(accession) {
    if (!accession) return;
    const yAxisSelect = document.getElementById('compareYAxisSelect');
    const yField = yAxisSelect.value || 'Relative abundance (%)';
    createCompareChart(accession, yField);
}

function createCompareChart(accession, yField = 'Relative abundance (%)') {
    // Missing sheet data is plotted as 0 but noted in the summary.
    const compareValues = compareSheets.map(sheetName => {
        const row = sheetDataMap[sheetName]?.find(r => String(r.Accession).trim() === String(accession).trim());
        const rawValue = row ? parseFloat(row[yField]) : NaN;
        return {
            sheetName,
            value: Number.isFinite(rawValue) ? rawValue : 0,
            hasData: !!row,
            row,
        };
    });

    const summary = document.getElementById('compareSummary');
    const container = document.getElementById('compareChartContainer');
    container.innerHTML = `
        <div id="compareChartWrapper" style="margin-bottom: 30px;">
            <h3 style="margin-bottom: 15px;">Accession: ${accession} 的 ${yField} 比較</h3>
            <div style="position: relative; height: 320px;">
                <canvas id="compareChart"></canvas>
            </div>
        </div>
    `;
    summary.style.display = 'block';
    summary.innerHTML = compareValues.map(item => {
        const relativeText = item.hasData ? item.value.toFixed(6) : '無資料';
        const abundanceValue = item.hasData ? item.row[`Abundance: ${item.sheetName}`] ?? '-' : '-';
        const countValue = item.hasData ? item.row[`Abundances Count: ${item.sheetName}`] ?? '-' : '-';
        return `<div style="margin-bottom: 6px;"><strong>${item.sheetName}</strong>: Relative abundance (%) = ${relativeText}, Abundance = ${abundanceValue}, Count = ${countValue}</div>`;
    }).join('');

    const ctx = document.getElementById('compareChart').getContext('2d');
    if (charts.compareChart) {
        charts.compareChart.destroy();
    }
    charts.compareChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: compareValues.map(item => item.sheetName),
            datasets: [{
                label: yField,
                data: compareValues.map(item => item.value),
                backgroundColor: ['#4f8cff', '#7dbeff', '#a8d5ff'],
                borderColor: ['#2f6fe5', '#4f9fe5', '#6bb8e5'],
                borderWidth: 1,
            }],
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: yField },
                },
            },
        },
    });
}

function buildChartConfigPanel(data) {
    const panel = document.getElementById('chartConfigPanel');
    const xSelect = document.getElementById('xAxisSelect');
    const ySelect = document.getElementById('yAxisSelect');

    xSelect.innerHTML = '';
    ySelect.innerHTML = '';

    // 排除內部標記欄位
    const columns = Object.keys(data[0] || {}).filter(col => !col.startsWith('_'));
    const numericFields = columns.filter(col => isNumericColumn(data, col));

    columns.forEach(col => {
        const option = document.createElement('option');
        option.value = col;
        option.textContent = col;
        xSelect.appendChild(option);
    });

    numericFields.forEach(col => {
        const option = document.createElement('option');
        option.value = col;
        option.textContent = col;
        ySelect.appendChild(option);
    });

    if (columns.length > 0 && numericFields.length > 0) {
        chartConfig.xAxis = xSelect.value = columns[0];
        chartConfig.yAxis = ySelect.value = numericFields[0];
        panel.style.display = 'block';
    } else {
        panel.style.display = 'none';
    }
}

function updateChartConfig() {
    const xSelect = document.getElementById('xAxisSelect');
    const ySelect = document.getElementById('yAxisSelect');
    chartConfig.xAxis = xSelect.value;
    chartConfig.yAxis = ySelect.value;
}

function refreshChart() {
    if (!currentSheetData.length) return;
    if (!chartConfig.xAxis || !chartConfig.yAxis) return;

    const chartsContainer = document.getElementById('chartsContainer');
    chartsContainer.innerHTML = '';
    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};

    const aggregated = aggregateByField(currentSheetData, chartConfig.xAxis, chartConfig.yAxis);
    createSingleBarChart(aggregated, chartConfig.xAxis, chartConfig.yAxis);
}

// Sum numeric values by category and keep the top 20 for readability.
function aggregateByField(data, xField, yField) {
    const grouped = {};
    data.forEach(row => {
        const category = String(row[xField] ?? '').trim() || '空值';
        const value = parseFloat(row[yField]);
        if (Number.isNaN(value)) return;
        grouped[category] = (grouped[category] || 0) + value;
    });

    return Object.entries(grouped)
        .map(([label, value]) => ({ label, value }))
        .sort((a, b) => b.value - a.value)
        .slice(0, 20);
}

function createSingleBarChart(data, xLabel, yLabel) {
    const container = document.createElement('div');
    container.style.marginBottom = '30px';
    container.innerHTML = `
        <h3 style="margin-bottom: 15px;">${xLabel} vs ${yLabel}</h3>
        <div style="position: relative; height: 360px;">
            <canvas id="singleBarChart"></canvas>
        </div>
    `;
    document.getElementById('chartsContainer').appendChild(container);

    const ctx = container.querySelector('#singleBarChart').getContext('2d');
    charts['singleBar'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: data.map(item => item.label),
            datasets: [{
                label: yLabel,
                data: data.map(item => item.value),
                backgroundColor: '#667eea',
                borderColor: '#764ba2',
                borderWidth: 1,
            }],
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: { display: true, text: yLabel },
                },
                x: {
                    title: { display: true, text: xLabel },
                },
            },
        },
    });
}






