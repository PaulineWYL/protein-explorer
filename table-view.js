// Render the selected sheet with Tabulator. MAD/MADEV/MADEVNP use a custom
// master/detail view; other sheets use the standard table view.
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
        // Hide columns that are effectively empty so wide reports stay readable.
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
    // Blue rows: protein/master rows have a value in Excel column A.
    const mainRows = firstField
        ? tabulatorData.filter(row => String(row[firstField] ?? '').trim() !== '')
        : tabulatorData;

    const table = new Tabulator(tableElement, {
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
    });

    table.on('rowClick', (e, row) => {
        const rowData = row.getData();
        const selectedValue = String(rowData[firstField] ?? '').trim();
        // Orange rows: take the block directly below the clicked blue row,
        // stopping at the next blue row to mirror the original Excel layout.
        const detailConfig = buildCompareDetailConfig(tabulatorData, rowData, columnFieldMap, firstField);
        showDetailModal(sheetName, selectedValue, detailConfig.rows, detailConfig.columns);
    });

    tables[sheetName] = table;
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

    // The first orange row is the detail table header.
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
        <div class="detail-modal-panel" role="dialog" aria-labelledby="detailModalTitle">
            <div class="detail-modal-header">
                <h3 id="detailModalTitle">${escapeHtml(sheetName)} - ${escapeHtml(selectedValue)}</h3>
                <button type="button" class="detail-modal-close" aria-label="Close" onclick="closeDetailModal()">&times;</button>
            </div>
            <div class="detail-modal-count">${detailRows.length} 筆明細資料</div>
            <div id="detailTable" class="detail-table-wrapper"></div>
        </div>
    `;

    document.body.appendChild(modal);
    makeDetailModalDraggable(modal);

    // The floating panel must be in the DOM before Tabulator can measure columns.
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

    setTimeout(() => detailTable?.redraw(true), 0);
}

function makeDetailModalDraggable(modal) {
    const header = modal.querySelector('.detail-modal-header');
    if (!header) return;

    let startX = 0;
    let startY = 0;
    let initialLeft = 0;
    let initialTop = 0;
    let isDragging = false;

    const moveModal = (clientX, clientY) => {
        if (!isDragging) return;

        // Keep the floating panel inside the viewport while dragging.
        const maxLeft = window.innerWidth - modal.offsetWidth;
        const maxTop = window.innerHeight - modal.offsetHeight;
        const nextLeft = Math.min(Math.max(0, initialLeft + clientX - startX), Math.max(0, maxLeft));
        const nextTop = Math.min(Math.max(0, initialTop + clientY - startY), Math.max(0, maxTop));

        modal.style.left = `${nextLeft}px`;
        modal.style.top = `${nextTop}px`;
        modal.style.right = 'auto';
        modal.style.bottom = 'auto';
    };

    header.addEventListener('pointerdown', (event) => {
        if (event.target.closest('button')) return;

        isDragging = true;
        startX = event.clientX;
        startY = event.clientY;
        initialLeft = modal.offsetLeft;
        initialTop = modal.offsetTop;
        header.setPointerCapture(event.pointerId);
        modal.classList.add('is-dragging');
    });

    header.addEventListener('pointermove', (event) => {
        moveModal(event.clientX, event.clientY);
    });

    header.addEventListener('pointerup', (event) => {
        isDragging = false;
        header.releasePointerCapture(event.pointerId);
        modal.classList.remove('is-dragging');
    });

    header.addEventListener('pointercancel', () => {
        isDragging = false;
        modal.classList.remove('is-dragging');
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








