let currentWorkbook = null;
let tables = {};
let charts = {};
let sheetDataMap = {};
let compareChart = null;
const compareSheets = ['MAD', 'MADEV', 'MADEVNP'];
const compareField = 'Relative abundance (%)';

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
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    
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
    
    // 顯示檔案資訊
    renderInfo(sheet, data);
}

// 渲染表格
function renderTable(sheetName, data) {
    const tableContainer = document.getElementById('table');
    
    // 如果已存在表格，銷毀舊表格
    if (tables[sheetName]) {
        tables[sheetName].destroy();
    }
    
    tableContainer.innerHTML = '<div id="tabulator-' + sheetName + '" style="margin-top: 20px;"></div>';
    
    // 準備欄位定義
    const columns = Object.keys(data[0]).map(key => ({
        title: key,
        field: key,
        width: 150,
        resizable: true,
        headerFilter: 'input'
    }));
    
    // 建立 Tabulator 表格
    tables[sheetName] = new Tabulator('#tabulator-' + sheetName, {
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
    
    // 尋找第一個數字欄位和第一個文字欄位
    const columns = Object.keys(data[0]);
    let numericCol = null;
    let textCol = null;
    
    columns.forEach(col => {
        const values = data.map(row => row[col]);
        if (!numericCol && values.some(v => !isNaN(v) && v !== '' && v !== null)) {
            numericCol = col;
        }
        if (!textCol && !values.every(v => !isNaN(v) && v !== '' && v !== null)) {
            textCol = col;
        }
    });
    
    if (numericCol && textCol) {
        // 建立分組統計圖
        const grouped = groupByAndSum(data, textCol, numericCol);
        
        if (grouped.length > 0 && grouped.length <= 20) {
            createBarChart(grouped, textCol, numericCol);
        }
    }
    
    if (numericCol) {
        // 建立分佈直方圖
        createHistogram(data, numericCol);
    }
}

// 分組統計
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
    const comparePanel = document.getElementById('comparePanel');
    const select = document.getElementById('compareAccessionSelect');
    select.innerHTML = '';
    comparePanel.style.display = 'none';

    const accessions = new Set();
    Object.values(sheetDataMap).forEach(data => {
        data.forEach(row => {
            if (row.Accession) {
                accessions.add(String(row.Accession).trim());
            }
        });
    });

    const sortedAccessions = Array.from(accessions).sort();
    if (sortedAccessions.length === 0) {
        return;
    }

    comparePanel.style.display = 'block';
    sortedAccessions.forEach(acc => {
        const option = document.createElement('option');
        option.value = acc;
        option.textContent = acc;
        select.appendChild(option);
    });

    updateComparisonChart(sortedAccessions[0]);
}

function updateComparisonChart(accession) {
    if (!accession) return;
    createCompareChart(accession);
}

function createCompareChart(accession) {
    const compareValues = compareSheets.map(sheetName => {
        const row = sheetDataMap[sheetName]?.find(r => String(r.Accession).trim() === String(accession).trim());
        const rawValue = row ? parseFloat(row[compareField]) : NaN;
        return {
            sheetName,
            value: Number.isFinite(rawValue) ? rawValue : 0,
            hasData: !!row,
            row,
        };
    });

    const compareChartContainer = document.getElementById('compareChartContainer');
    const summary = document.getElementById('compareSummary');
    compareChartContainer.innerHTML = `
        <div id="compareChartWrapper" style="margin-bottom: 30px;">
            <h3 style="margin-bottom: 15px;">Accession: ${accession} 的 Relative abundance (%) 比較</h3>
            <div style="position: relative; height: 320px;">
                <canvas id="compareChart"></canvas>
            </div>
        </div>
    `;

    if (compareChart) {
        compareChart.destroy();
    }

    const ctx = document.getElementById('compareChart').getContext('2d');
    compareChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: compareValues.map(item => item.sheetName),
            datasets: [{
                label: 'Relative abundance (%)',
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
                    title: { display: true, text: 'Relative abundance (%)' },
                },
            },
        },
    });

    summary.innerHTML = compareValues.map(item => {
        const relativeText = item.hasData ? item.value.toFixed(6) : '無資料';
        const abundanceValue = item.hasData ? item.row[`Abundance: ${item.sheetName}`] ?? '-' : '-';
        const countValue = item.hasData ? item.row[`Abundances Count: ${item.sheetName}`] ?? '-' : '-';
        return `<div style="margin-bottom: 6px;"><strong>${item.sheetName}</strong>: Relative abundance (%) = ${relativeText}, Abundance = ${abundanceValue}, Count = ${countValue}</div>`;
    }).join('');
}

function groupByAndSum(data, groupCol, sumCol) {
    const grouped = {};
    
    data.forEach(row => {
        const groupKey = String(row[groupCol]).substring(0, 20);
        const value = parseFloat(row[sumCol]) || 0;
        
        if (!grouped[groupKey]) {
            grouped[groupKey] = 0;
        }
        grouped[groupKey] += value;
    });
    
    return Object.entries(grouped)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15)
        .map(([key, value]) => ({ label: key, value: value }));
}

// 建立條形圖
function createBarChart(data, xLabel, yLabel) {
    const container = document.createElement('div');
    container.style.marginBottom = '30px';
    container.innerHTML = `
        <h3 style="margin-bottom: 15px;">按 "${xLabel}" 統計 "${yLabel}"</h3>
        <div style="position: relative; height: 300px;">
            <canvas id="barChart"></canvas>
        </div>
    `;
    
    document.getElementById('chartsContainer').appendChild(container);
    
    const ctx = container.querySelector('#barChart').getContext('2d');
    charts['barChart'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: data.map(d => d.label),
            datasets: [{
                label: yLabel,
                data: data.map(d => d.value),
                backgroundColor: '#667eea',
                borderColor: '#764ba2',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: true }
            },
            scales: {
                y: { beginAtZero: true }
            }
        }
    });
}

// 建立直方圖
function createHistogram(data, col) {
    const values = data
        .map(row => parseFloat(row[col]))
        .filter(v => !isNaN(v) && v !== null);
    
    if (values.length === 0) return;
    
    const min = Math.min(...values);
    const max = Math.max(...values);
    const binCount = Math.min(20, Math.ceil(Math.sqrt(values.length)));
    const binSize = (max - min) / binCount || 1;
    
    const bins = Array(binCount).fill(0);
    values.forEach(v => {
        const binIndex = Math.min(binCount - 1, Math.floor((v - min) / binSize));
        bins[binIndex]++;
    });
    
    const binLabels = Array.from({length: binCount}, (_, i) => 
        (min + i * binSize).toFixed(1)
    );
    
    const container = document.createElement('div');
    container.style.marginBottom = '30px';
    container.innerHTML = `
        <h3 style="margin-bottom: 15px;">"${col}" 分佈直方圖</h3>
        <div style="position: relative; height: 300px;">
            <canvas id="histogramChart"></canvas>
        </div>
    `;
    
    document.getElementById('chartsContainer').appendChild(container);
    
    const ctx = container.querySelector('#histogramChart').getContext('2d');
    charts['histogramChart'] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: binLabels,
            datasets: [{
                label: '頻數',
                data: bins,
                backgroundColor: '#764ba2',
                borderColor: '#667eea',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: { beginAtZero: true, title: { display: true, text: '頻數' } }
            }
        }
    });
}

// 渲染檔案資訊
function renderInfo(sheet, data) {
    const infoContent = document.getElementById('infoContent');
    
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
