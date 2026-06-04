// File and sheet controls, status messages, and lightweight UI rendering.
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


