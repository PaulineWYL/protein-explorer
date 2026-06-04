// Shared app state. Files are loaded as classic scripts, so these globals are
// intentionally shared across parser, table, chart, and UI modules.
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

// Sheet-specific header rows. Values are 1-based to match Excel row numbers.
const sheetHeaderConfig = {
    'Description of Column Title': [3, 20],
    '樞紐': 4,
    'MAD': 1,
    'MADEV': 1,
    'MADEVNP': 1,
};



