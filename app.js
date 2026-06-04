// Entry point: wire upload interactions and hand parsed workbooks to the
// parser/rendering modules loaded before this file.
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');

uploadArea.addEventListener('dragover', (event) => {
    event.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (event) => {
    event.preventDefault();
    uploadArea.classList.remove('dragover');
    handleFile(event.dataTransfer.files[0]);
});

fileInput.addEventListener('change', (event) => {
    handleFile(event.target.files[0]);
});

function handleFile(file) {
    if (!file) return;

    // Browser file.type is not always reliable for Excel files, so keep the
    // extension check below as a fallback.
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel',
        'text/csv',
        'application/octet-stream',
    ];

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
    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            currentWorkbook = XLSX.read(data, { type: 'array' });

            if (!currentWorkbook.SheetNames || currentWorkbook.SheetNames.length === 0) {
                showError('檔案中沒有找到工作表');
                return;
            }

            showSuccess(`成功載入 "${file.name}"（${currentWorkbook.SheetNames.length} 個工作表）`);
            updateFileInfo(file);
            // Populate controls before rendering the first sheet.
            initializeSheetSelector();
            loadCompareSheetData();
            initializeComparePanel();
            switchSheet(currentWorkbook.SheetNames[0]);
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




