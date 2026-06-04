// Parse the current worksheet into normalized rows, then render table, stats,
// and charts from the same data set.
function switchSheet(sheetName) {
    const sheet = currentWorkbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    if (rawData.length === 0) {
        showError('該工作表沒有資料');
        return;
    }

    const headerConfig = sheetHeaderConfig[sheetName];
    let headerRowIndices = [];

    // Known report sheets have fixed header rows; unknown sheets use detection.
    if (headerConfig !== undefined) {
        headerRowIndices = Array.isArray(headerConfig)
            ? headerConfig.map(row => row - 1)
            : [headerConfig - 1];
    } else {
        headerRowIndices = [detectHeaderRow(rawData)];
    }

    headerRowIndices = headerRowIndices.filter(idx => idx >= 0 && idx < rawData.length - 1);

    if (headerRowIndices.length === 0) {
        showError('無法偵測表頭行');
        return;
    }

    const headerRowIndex = headerRowIndices[0];
    const columnNames = rawData[headerRowIndex].map((val, idx) => {
        const name = String(val || '').trim();
        return name === '' || name === 'EMPTY' ? `Column ${idx + 1}` : name;
    });

    let data = rawData.slice(headerRowIndex + 1).map((row, originalIdx) => {
        const obj = {};
        const indentLevel = detectIndentLevel(row);

        // Normalize cells to strings so filters, tables, and charts see
        // consistent values regardless of Excel cell type.
        columnNames.forEach((colName, idx) => {
            const value = row[idx];
            obj[colName] = value === 'EMPTY' || value === null || value === undefined
                ? ''
                : String(value).trim();
        });

        obj._indentLevel = indentLevel;
        obj._isSubHeader = isSubHeader(rawData, headerRowIndex + 1 + originalIdx);
        return obj;
    });

    // Drop repeated header rows and rows with no visible values.
    data = data.filter(row => {
        const hasVisibleValue = Object.entries(row).some(([key, value]) => !key.startsWith('_') && value !== '');
        return hasVisibleValue && !isHeaderRow(row, columnNames);
    });

    if (data.length === 0) {
        showError('該工作表沒有有效資料');
        return;
    }

    renderRawSheet(sheet);
    renderTable(sheetName, data, columnNames);
    renderStats(data);
    renderCharts(sheetName, data);
}

// In the report sheets, detail rows are indicated by an empty column A.
function detectIndentLevel(row) {
    const firstCell = String(row[0] || '').trim();
    return firstCell === '' ? 1 : 0;
}

function isSubHeader(rawData, rowIndex) {
    if (rowIndex <= 0) return false;

    const currentRowFirstCol = String(rawData[rowIndex][0] || '').trim();
    const prevRowFirstCol = String(rawData[rowIndex - 1][0] || '').trim();
    return currentRowFirstCol === '' && prevRowFirstCol !== '';
}

// Pick the densest row near the top as a pragmatic default for unknown sheets.
function detectHeaderRow(rawData) {
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

// Exported reports can repeat header rows inside the data body.
function isHeaderRow(row, columnNames) {
    const values = Object.entries(row)
        .filter(([key]) => !key.startsWith('_'))
        .map(([, value]) => String(value).trim());
    const matchCount = values.filter(value => value !== '' && columnNames.includes(value)).length;
    return values.length > 0 && matchCount / values.length > 0.7;
}

function renderRawSheet(sheet) {
    const rawContainer = document.getElementById('rawView');
    if (!rawContainer) return;

    const rawHtml = XLSX.utils.sheet_to_html(sheet, {
        editable: false,
        blankrows: true,
        header: '',
    });
    rawContainer.innerHTML = `<div class="raw-sheet-wrapper">${rawHtml}</div>`;
}








