/**
 * 簡單的 Express 伺服器（可選用）
 * 如果已安裝 Node.js，可以使用此方式啟動
 * 
 * 使用方式:
 * 1. npm install
 * 2. npm start
 * 3. 打開 http://localhost:8000
 */

const express = require('express');
const path = require('path');
const app = express();

const PORT = 8000;

// 提供靜態檔案
app.use(express.static(path.join(__dirname)));

// 主頁路由
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// 啟動伺服器
app.listen(PORT, () => {
    console.log('\n╔════════════════════════════════════════╗');
    console.log('║  Excel 視覺化工具 - 伺服器已啟動      ║');
    console.log('╠════════════════════════════════════════╣');
    console.log(`║  🌐 http://localhost:${PORT}              ║`);
    console.log('║  按 Ctrl+C 停止伺服器                  ║');
    console.log('╚════════════════════════════════════════╝\n');
});
