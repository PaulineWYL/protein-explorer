let currentWorkbook = null;
let tables = {};
let charts = {};
let currentSheetData = [];
let currentSheetName = '';
let chartConfig = { xAxis: '', yAxis: '' };
let sheetDataMap = {};
let detailTable = null;
const compareSheets = ['MAD', 'MADEV', 'MADEVNP'];
const noCollapseSheets = ['樞紐', '總表'];

// 各 Sheet 的表頭行配置
const sheetHeaderConfig = {
    'Description of Column Title': [3, 20],  // row 3 & row 20（0-indexed: 2, 19）
    '樞紐': 4,                               // row 4（0-indexed: 3）
    'MAD': 1,                                 // row 1（0-indexed: 0）
    'MADEV': 1,
    'MADEVNP': 1
    // 其他 sheet 使用自動檢測
};


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
    
    // 取得原始數據（包含所有行）
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    
    if (rawData.length === 0) {
        showError('該工作表沒有資料');
        return;
    }
    
    // 取得該 sheet 的表頭配置
    const headerConfig = sheetHeaderConfig[sheetName];
    let headerRowIndices = [];
    
    if (headerConfig !== undefined) {
        // 使用指定的表頭行
        if (Array.isArray(headerConfig)) {
            headerRowIndices = headerConfig.map(row => row - 1);  // 轉換為 0-indexed
        } else {
            headerRowIndices = [headerConfig - 1];  // 轉換為 0-indexed
        }
    } else {
        // 自動檢測表頭行
        const detected = detectHeaderRow(rawData);
        headerRowIndices = [detected];
    }
    
    // 驗證表頭行有效性
    headerRowIndices = headerRowIndices.filter(idx => idx >= 0 && idx < rawData.length - 1);
    
    if (headerRowIndices.length === 0) {
        showError('無法偵測表頭行');
        return;
    }
    
    // 使用第一個表頭行作為列名
    const headerRowIndex = headerRowIndices[0];
    const columnNames = rawData[headerRowIndex].map((val, idx) => {
        const name = String(val || '').trim();
        return name === '' || name === 'EMPTY' ? `Column ${idx + 1}` : name;
    });
    
    // 提取數據行（第一個表頭行之後的所有行，排除其他表頭行）
    let data = rawData.slice(headerRowIndex + 1).map((row, originalIdx) => {
        const obj = {};
        const indentLevel = detectIndentLevel(row);  // 檢測縮排層級
        
        columnNames.forEach((colName, idx) => {
            const value = row[idx];
            if (value === 'EMPTY' || value === null || value === undefined) {
                obj[colName] = '';
            } else {
                // 確保數值被正確保存為字符串
                obj[colName] = String(value).trim();
            }
        });
        
        // 添加內部縮排層級標記（用於行分組）
        obj._indentLevel = indentLevel;
        obj._isSubHeader = isSubHeader(rawData, headerRowIndex + 1 + originalIdx);
        
        return obj;
    });
    
    // 移除其他表頭行和空行
    data = data.filter(row => !isHeaderRow(row, columnNames) && Object.values(row).some(v => v !== '' && !v.toString().startsWith('_')));
    
    if (data.length === 0) {
        showError('該工作表沒有有效資料');
        return;
    }
    
    // 顯示原始 Excel 表格
    renderRawSheet(sheet);
    
    // 建立表格
    renderTable(sheetName, data, columnNames);
    
    // 建立統計數據
    renderStats(data);
    
    // 建立圖表
    renderCharts(sheetName, data);
}

// 檢測縮排層級（基於第一欄是否為空）
function detectIndentLevel(row) {
    // 檢查前 3 欄是否為空來判斷縮排層級
    for (let i = 0; i < Math.min(1, row.length); i++) {
        const val = String(row[i] || '').trim();
        if (val !== '') {
            return 0;  // 非縮排
        }
    }
    return 1;  // 縮排層級 1
}

// 檢測是否為子表頭（當前 A 欄為空，上一欄 A 有值）
function isSubHeader(rawData, rowIndex) {
    if (rowIndex <= 0) return false;
    
    const currentRowFirstCol = String(rawData[rowIndex][0] || '').trim();
    const prevRowFirstCol = String(rawData[rowIndex - 1][0] || '').trim();
    
    // 當前欄 A 空，上一欄 A 有值 = 子表頭
    return currentRowFirstCol === '' && prevRowFirstCol !== '';
}

// 智能檢測表頭行位置
function detectHeaderRow(rawData) {
    // 找到第一個有最多非空值的行作為表頭
    let maxNonEmptyCount = 0;
    let headerIndex = 0;
    
    for (let i = 0; i < Math.min(5, rawData.length); i++) {
        const nonEmptyCount = rawData[i].filter(val => val !== '' && val !== null && val !== undefined).length;
        if (nonEmptyCount > maxNonEmptyCount) {
            maxNonEmptyCount = nonEmptyCount;
            headerIndex = i;
        }
    }
    
    return headerIndex;
}

// 檢查是否為列名行（整列都是欄位名稱）
function isHeaderRow(row, columnNames) {
    const values = Object.values(row).filter(v => !v.toString().startsWith('_')).map(v => String(v).trim());
    const matchCount = values.filter(v => v !== '' && columnNames.includes(v)).length;
    // 如果匹配的值超過該行70%以上，視為列名行
    return values.length > 0 && matchCount / values.length > 0.7;
}

// 渲染表格
function renderTable(sheetName, data, columnNames) {
    const tableContainer = document.getElementById('table');

    if (tables[sheetName]) {
        tables[sheetName].destroy();
    }
    closeDetailModal();

    const tableId = 'tabulator-' + sanitizeId(sheetName);
    tableContainer.innerHTML = `<div id="${tableId}" class="table-wrapper"></div>`;
    const tableElement = document.getElementById(tableId);
    if (!tableElement) {
        showError('無法建立表格，請重新整理頁面。');
        return;
    }

    const isCompareSheet = compareSheets.includes(sheetName);
    const disableGrouping = noCollapseSheets.includes(sheetName);
    const allKeys = columnNames || Array.from(new Set(data.flatMap(row => Object.keys(row).filter(k => !k.startsWith('_')))));

    let validColumns;
    if (disableGrouping) {
        validColumns = allKeys;
    } else {
        validColumns = allKeys.filter(key => {
            const nonEmptyCount = data.filter(row => {
                const value = row[key];
                return value !== '' && value !== null && value !== undefined && value !== 'EMPTY';
            }).length;
            return nonEmptyCount > 0 && nonEmptyCount >= Math.max(1, Math.ceil(data.length * 0.01));
        });

        if (!isCompareSheet) {
            validColumns = validColumns.sort();
        }
    }

    const columnFieldMap = validColumns.map((key, index) => ({
        title: key,
        field: `field_${index}`,
        originalKey: key,
    }));

    const tabulatorData = data.map((row, rowIndex) => {
        const safeRow = {};
        columnFieldMap.forEach(col => {
            safeRow[col.field] = row[col.originalKey];
        });
        safeRow._indentLevel = row._indentLevel;
        safeRow._isSubHeader = row._isSubHeader;
        safeRow._sourceIndex = rowIndex;
        return safeRow;
    });

    const avgWidth = Math.max(120, Math.min(200, 1400 / Math.max(validColumns.length, 1)));
    const columns = columnFieldMap.map(col => ({
        title: col.title,
        field: col.field,
        width: avgWidth,
        minWidth: 80,
        resizable: true,
        headerFilter: 'input',
        tooltip: true,
        formatter: (cell) => {
            const value = cell.getValue();
            return value === null || value === undefined || value === '' ? '' : String(value);
        },
    }));

    if (isCompareSheet) {
        renderCompareSheetTable(sheetName, tableElement, tabulatorData, columnFieldMap, columns);
        return;
    }

    let groupByField = null;
    if (!disableGrouping && columnFieldMap.length > 0) {
        const firstField = columnFieldMap[0].field;
        const uniqueCount = new Set(tabulatorData.map(r => r[firstField])).size;
        if (uniqueCount < tabulatorData.length * 0.5 && uniqueCount > 1) {
            groupByField = firstField;
        }
    }

    const tabulatorConfig = {
        data: tabulatorData,
        columns: columns,
        layout: 'fitColumns',
        layoutColumnsOnNewData: false,
        responsiveLayout: false,
        pagination: 'local',
        paginationSize: 25,
        movableColumns: true,
        selectable: true,
        clipboard: true,
        virtualDom: true,
        virtualDomBuffer: 10,
        height: '600px',
    };

    if (groupByField) {
        tabulatorConfig.groupBy = groupByField;
        tabulatorConfig.groupStartOpen = false;
    }

    tables[sheetName] = new Tabulator(tableElement, tabulatorConfig);
}

function renderCompareSheetTable(sheetName, tableElement, tabulatorData, columnFieldMap, columns) {
    const firstField = columnFieldMap[0]?.field;
    const mainRows = firstField
        ? tabulatorData.filter(row => String(row[firstField] ?? '').trim() !== '')
        : tabulatorData;

    tables[sheetName] = new Tabulator(tableElement, {
        data: mainRows,
        columns: columns,
        layout: 'fitColumns',
        layoutColumnsOnNewData: false,
        responsiveLayout: false,
        pagination: 'local',
        paginationSize: 25,
        movableColumns: true,
        selectable: true,
        clipboard: true,
        virtualDom: true,
        virtualDomBuffer: 10,
        height: '600px',
        rowFormatter: (row) => {
            row.getElement().classList.add('compare-main-row');
        },
        rowClick: (e, row) => {
            const rowData = row.getData();
            const selectedValue = String(rowData[firstField] ?? '').trim();
            const detailConfig = buildCompareDetailConfig(tabulatorData, rowData, columnFieldMap, firstField);
            showDetailModal(sheetName, selectedValue, detailConfig.rows, detailConfig.columns);
        },
    });
}

function buildCompareDetailConfig(tabulatorData, selectedRow, columnFieldMap, firstField) {
    const startIndex = selectedRow._sourceIndex;
    const nextMainRow = tabulatorData.find(row => {
        return row._sourceIndex > startIndex && String(row[firstField] ?? '').trim() !== '';
    });
    const endIndex = nextMainRow ? nextMainRow._sourceIndex : Infinity;
    const detailBlock = tabulatorData.filter(row => row._sourceIndex > startIndex && row._sourceIndex < endIndex);

    if (detailBlock.length === 0) {
        return { rows: [], columns: [] };
    }

    const headerRow = detailBlock[0];
    const detailFieldMap = columnFieldMap
        .map((col, index) => ({
            sourceField: col.field,
            field: `detail_${index}`,
            title: String(headerRow[col.field] ?? '').trim(),
        }))
        .filter(col => col.title !== '');

    const detailColumns = detailFieldMap.map(col => ({
        title: col.title,
        field: col.field,
        minWidth: 100,
        resizable: true,
        headerFilter: 'input',
        tooltip: true,
        formatter: (cell) => {
            const value = cell.getValue();
            return value === null || value === undefined || value === '' ? '' : String(value);
        },
    }));

    const detailRows = detailBlock.slice(1).map(row => {
        const detailRow = {};
        detailFieldMap.forEach(col => {
            detailRow[col.field] = row[col.sourceField];
        });
        return detailRow;
    });

    return { rows: detailRows, columns: detailColumns };
}
function showDetailModal(sheetName, selectedValue, detailRows, columns) {
    closeDetailModal();

    const modal = document.createElement('div');
    modal.id = 'detailModal';
    modal.className = 'detail-modal';
    modal.innerHTML = `
        <div class="detail-modal-backdrop" onclick="closeDetailModal()"></div>
        <div class="detail-modal-panel" role="dialog" aria-modal="true" aria-labelledby="detailModalTitle">
            <div class="detail-modal-header">
                <h3 id="detailModalTitle">${escapeHtml(sheetName)} - ${escapeHtml(selectedValue)}</h3>
                <button type="button" class="detail-modal-close" aria-label="Close" onclick="closeDetailModal()">&times;</button>
            </div>
            <div class="detail-modal-count">${detailRows.length} 筆明細資料</div>
            <div id="detailTable" class="detail-table-wrapper"></div>
        </div>
    `;

    document.body.appendChild(modal);

    detailTable = new Tabulator('#detailTable', {
        data: detailRows,
        columns: columns,
        layout: 'fitColumns',
        layoutColumnsOnNewData: false,
        responsiveLayout: false,
        pagination: 'local',
        paginationSize: 10,
        movableColumns: true,
        clipboard: true,
        height: '420px',
        placeholder: '沒有符合條件的明細資料',
    });
}

function closeDetailModal() {
    if (detailTable) {
        detailTable.destroy();
        detailTable = null;
    }

    const modal = document.getElementById('detailModal');
    if (modal) {
        modal.remove();
    }
}

function escapeHtml(value) {
    const div = document.createElement('div');
    div.textContent = String(value ?? '');
    return div.innerHTML;
}

function renderRawSheet(sheet) {
    const rawContainer = document.getElementById('rawView');
    if (!rawContainer) return;
    
    const rawHtml = XLSX.utils.sheet_to_html(sheet, {
        editable: false,
        blankrows: true,
        header: ''
    });
    rawContainer.innerHTML = `<div class="raw-sheet-wrapper">${rawHtml}</div>`;
}

// 構建基於縮排層級的分組數據結構
function buildIndentGroups(data) {
    const grouped = [];
    let currentGroupKey = null;
    let groupCounter = 0;
    
    data.forEach((row, idx) => {
        const indentLevel = row._indentLevel || 0;
        const isSubHeader = row._isSubHeader || false;
        
        if (indentLevel === 0 && !isSubHeader) {
            // 主行，創建新分組
            groupCounter++;
            currentGroupKey = `Group_${groupCounter}`;
        } else if (indentLevel === 0 && isSubHeader) {
            // 子表頭，仍屬於當前分組
            if (!currentGroupKey) {
                groupCounter++;
                currentGroupKey = `Group_${groupCounter}`;
            }
        }
        
        // 添加分組鍵
        row._groupKey = currentGroupKey || `Group_${++groupCounter}`;
        grouped.push(row);
    });
    
    return grouped;
}

// 渲染統計數據
function renderStats(data) {
    const statsContainer = document.getElementById('stats');
    
    // 計算統計資料（排除內部標記）
    const rows = data.length;
    const allKeys = data.length > 0 ? Object.keys(data[0]).filter(k => !k.startsWith('_')) : [];
    const cols = allKeys.length;
    
    // 尋找數字欄位
    let numericCols = 0;
    allKeys.forEach(key => {
        const values = data.map(row => row[key]);
        if (values.some(v => {
            const num = parseFloat(v);
            return !isNaN(num) && v !== '';
        })) {
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

