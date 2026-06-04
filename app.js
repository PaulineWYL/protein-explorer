let currentWorkbook = null;
let tables = {};
let charts = {};
let currentSheetData = [];
let currentSheetName = '';
let chartConfig = { xAxis: '', yAxis: '' };
let sheetDataMap = {};
const compareSheets = ['MAD', 'MADEV', 'MADEVNP'];

// 初始化拖曳上傳
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    handleFile(e.dataTransfer.files[0]);
});

fileInput.addEventListener('change', (e) => {
    handleFile(e.target.files[0]);
});

// 處理檔案上傳
function handleFile(file) {
    if (!file) return;
    
    // 檢查檔案類型
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                        'application/vnd.ms-excel', 
                        'text/csv',
                        'application/octet-stream'];
    
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls|csv)$/i)) {
        showError('請選擇有效的 Excel 或 CSV 檔案');
        return;
    }
    
    if (file.size > 50 * 1024 * 1024) {
        showError('檔案大小超過 50MB 限制');
        return;
    }
    
    showLoading('正在讀取檔案...');
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            currentWorkbook = XLSX.read(data, { type: 'array' });
            
            if (!currentWorkbook.SheetNames || currentWorkbook.SheetNames.length === 0) {
                showError('檔案中沒有找到工作表');
                return;
            }
            
            // 顯示檔案資訊
            showSuccess(`成功載入 "${file.name}"（${currentWorkbook.SheetNames.length} 個工作表）`);
            updateFileInfo(file);
            
            // 初始化工作表選擇器
            initializeSheetSelector();
            loadCompareSheetData();
            initializeComparePanel();
            
            // 載入第一個工作表
            switchSheet(currentWorkbook.SheetNames[0]);
            
            // 顯示內容區域
            document.getElementById('contentArea').style.display = 'block';
            
        } catch (error) {
            console.error('Error:', error);
            showError(`解析檔案出錯：${error.message}`);
        }
    };
    
    reader.onerror = () => {
        showError('讀取檔案失敗');
    };
    
    reader.readAsArrayBuffer(file);
}

// 更新檔案資訊顯示
function updateFileInfo(file) {
    const fileInfo = document.getElementById('fileInfo');
    const size = (file.size / 1024).toFixed(2);
    fileInfo.innerHTML = `
        <span class="file-name">📄 ${file.name}</span>
        <span>大小：${size} KB | 修改時間：${new Date(file.lastModified).toLocaleString()}</span>
    `;
    fileInfo.classList.add('show');
}

// 初始化工作表選擇器
function initializeSheetSelector() {
    const select = document.getElementById('sheetSelect');
    select.innerHTML = '';
    
    currentWorkbook.SheetNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        select.appendChild(option);
    });
}

// 切換工作表
function switchSheet(sheetName) {
    const sheet = currentWorkbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    
    // 清理數據：將 'EMPTY' 字符串和其他可能的空值表示轉換為空字符串
    data = data.map(row => {
        const cleanedRow = {};
        for (const [key, value] of Object.entries(row)) {
            if (value === 'EMPTY' || value === null || value === undefined) {
                cleanedRow[key] = '';
            } else {
                cleanedRow[key] = value;
            }
        }
        return cleanedRow;
    });
    
    if (data.length === 0) {
        showError('該工作表沒有資料');
        return;
    }
    
    // 建立表格
    renderTable(sheetName, data);
    
    // 建立統計數據
    renderStats(data);
    
    // 建立圖表
    renderCharts(sheetName, data);
}

// 渲染表格
function renderTable(sheetName, data) {
    const tableContainer = document.getElementById('table');
    
    // 如果已存在表格，銷毀舊表格
    if (tables[sheetName]) {
        tables[sheetName].destroy();
    }
    
    const tableId = 'tabulator-' + sanitizeId(sheetName);
    tableContainer.innerHTML = `<div id="${tableId}" style="margin-top: 20px;"></div>`;
    const tableElement = document.getElementById(tableId);
    if (!tableElement) {
        showError('無法建立表格，請重新整理頁面。');
        return;
    }
    
    // 準備欄位定義
    const columns = Object.keys(data[0]).map(key => ({
        title: key,
        field: key,
        width: 150,
        resizable: true,
        headerFilter: 'input'
    }));
    
    // 建立 Tabulator 表格
    tables[sheetName] = new Tabulator(tableElement, {
        data: data,
        columns: columns,
        layout: 'fitDataFill',
        responsiveLayout: 'collapse',
        pagination: 'local',
        paginationSize: 25,
        movableColumns: true,
        selectable: true,
        clipboard: true,
    });
}

// 渲染統計數據
function renderStats(data) {
    const statsContainer = document.getElementById('stats');
    
    // 計算統計資料
    const rows = data.length;
    const cols = Object.keys(data[0]).length;
    
    // 尋找數字欄位
    let numericCols = 0;
    Object.keys(data[0]).forEach(key => {
        const values = data.map(row => row[key]);
        if (values.some(v => !isNaN(v) && v !== '')) {
            numericCols++;
        }
    });
    
    statsContainer.innerHTML = `
        <div class="stat-box">
            <div class="stat-number">${rows}</div>
            <div class="stat-label">行數</div>
        </div>
        <div class="stat-box">
            <div class="stat-number">${cols}</div>
            <div class="stat-label">欄數</div>
        </div>
        <div class="stat-box">
            <div class="stat-number">${numericCols}</div>
            <div class="stat-label">數字欄位</div>
        </div>
        <div class="stat-box">
            <div class="stat-number">${(rows * cols).toLocaleString()}</div>
            <div class="stat-label">總儲存格數</div>
        </div>
    `;
}

// 渲染圖表
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

    const columns = Object.keys(data[0] || {});
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

// 將工作表名稱轉換為有效的 HTML ID
function sanitizeId(str) {
    return str.replace(/[^a-zA-Z0-9-_]/g, '_');
}

function isNumericColumn(data, col) {
    return data.some(row => {
        const value = row[col];
        return value !== '' && value !== null && !Number.isNaN(parseFloat(value));
    });
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


// 渲染檔案資訊
function renderInfo(sheet, data) {
    const infoContent = document.getElementById('infoContent');
    if (!infoContent) return;
    
    const info = {
        '工作表資訊': {
            '工作表名稱': currentWorkbook.SheetNames.join(', '),
            '目前工作表': sheet.name || 'Sheet',
            '資料行數': data.length,
            '欄數': Object.keys(data[0] || {}).length,
        },
        '欄位詳情': Object.keys(data[0] || {}).map((col, idx) => ({
            '序號': idx + 1,
            '欄名': col,
            '非空值': data.filter(row => row[col] !== '' && row[col] !== null).length,
            '唯一值': new Set(data.map(row => row[col])).size
        }))
    };
    
    infoContent.style.display = 'block';
    infoContent.textContent = JSON.stringify(info, null, 2);
}

// 切換選項卡
function switchTab(tabName) {
    // 隱藏所有選項卡
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    
    // 移除所有按鈕的活動狀態
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    // 顯示選定的選項卡
    document.getElementById(tabName).classList.add('active');
    
    // 突出顯示按鈕
    event.target.classList.add('active');
}

// 錯誤提示
function showError(message) {
    const errorMsg = document.getElementById('errorMsg');
    errorMsg.textContent = '❌ ' + message;
    errorMsg.style.display = 'block';
    
    const successMsg = document.getElementById('successMsg');
    successMsg.style.display = 'none';
}

// 成功提示
function showSuccess(message) {
    const successMsg = document.getElementById('successMsg');
    successMsg.textContent = '✅ ' + message;
    successMsg.style.display = 'block';
    
    const errorMsg = document.getElementById('errorMsg');
    errorMsg.style.display = 'none';
}

// 載入提示
function showLoading(message) {
    const fileInfo = document.getElementById('fileInfo');
    fileInfo.innerHTML = '<div class="loading"><i class="fas fa-spinner fa-spin"></i> ' + message + '</div>';
    fileInfo.classList.add('show');
}
