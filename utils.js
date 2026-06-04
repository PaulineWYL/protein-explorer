// Small shared helpers used by multiple modules.
function sanitizeId(str) {
    return str.replace(/[^a-zA-Z0-9-_]/g, '_');
}

function escapeHtml(value) {
    const div = document.createElement('div');
    div.textContent = String(value ?? '');
    return div.innerHTML;
}

function isNumericColumn(data, col) {
    return data.some(row => {
        const value = row[col];
        return value !== '' && value !== null && !Number.isNaN(parseFloat(value));
    });
}


