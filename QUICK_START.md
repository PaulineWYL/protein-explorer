# 快速設置指南 ⚡ (5 分鐘內)

## 您已獲得的檔案

```
protein/
├── 📄 index.html          # 主程式（UI 介面）
├── 📜 app.js              # 前端邏輯
├── 🚀 start_server.bat    # Windows 快速啟動（推薦）
├── 🌐 server.js           # Node.js 伺服器（可選）
├── 📦 package.json        # Node.js 配置（可選）
├── 📋 README.md           # 完整說明文檔
└── 📊 26042903-Protein Peptide Report.xlsx  # 您的 Excel 檔案
```

---

## 🎯 最快啟動方式（推薦）

### ✅ 方式 1：只需 2 個步驟（Python）

**第 1 步：確認已安裝 Python**

```bash
python --version
```

如果出現版本號，表示已安裝；如果出現錯誤，[下載安裝 Python](https://www.python.org/downloads/)

**第 2 步：雙擊 `start_server.bat` 啟動**

- 檔案會自動在瀏覽器中打開
- 您現在可以上傳和查看 Excel 檔案！

---

## 📦 替代方式（如果您已安裝 Node.js）

### ✅ 方式 2：使用 Node.js + Express

**第 1 步：安裝依賴**

```bash
cd c:\Users\PaulineWang\Desktop\protein
npm install
```

**第 2 步：啟動伺服器**

```bash
npm start
```

**第 3 步：打開瀏覽器**

訪問 `http://localhost:8000`

---

## 🌐 方式 3：直接打開（最簡單，但功能受限）

您也可以直接在瀏覽器中打開 `index.html` 檔案：

- 在 Windows 資源管理器中找到 `index.html`
- 右鍵選擇「開啟方式」→「瀏覽器」
- 或將 `index.html` 拖曳到瀏覽器中

⚠️ **注意**：由於瀏覽器安全限制，某些功能在本地檔案協議（`file://`）中可能無法正常工作。建議使用上面的伺服器方式。

---

## 📊 如何使用

1. **上傳檔案**
   - 點擊「選擇檔案」或將 Excel 檔案拖曳到上傳區域
   - 支持 `.xlsx`、`.xls` 和 `.csv` 格式

2. **查看數據**
   - **表格檢視**：搜尋、排序、篩選您的數據
   - **圖表分析**：自動生成統計圖表
   - **檔案資訊**：查看詳細的列統計

3. **互動操作**
   - 在表格標題下搜尋
   - 點擊欄進行排序
   - 拖曳欄邊界調整寬度
   - 選擇多行進行批量操作

---

## 🔧 故障排除

### 問題 1：雙擊 `.bat` 檔案無反應

**解決方案**：

```bash
# 手動執行
cd c:\Users\PaulineWang\Desktop\protein
python -m http.server 8000
```

然後在瀏覽器中訪問：`http://localhost:8000`

### 問題 2："已開啟埠 8000"

**解決方案**：改用其他埠

```bash
python -m http.server 8001
# 訪問 http://localhost:8001
```

### 問題 3：某些功能無法在本地檔案中工作

**解決方案**：使用伺服器方式（`start_server.bat` 或 `npm start`）而不是直接打開 `index.html`

---

## ✨ 功能亮點

- 🚀 **零設置** - 無需複雜配置
- 🔒 **隱私安全** - 檔案不上傳到任何伺服器
- 📱 **響應式設計** - 支援各種螢幕大小
- 🎨 **現代 UI** - 漂亮的界面設計
- 📊 **智能分析** - 自動檢測和分析數據
- 🔍 **強大搜尋** - 快速查找數據

---

## 💡 進階技巧

### 在其他電腦上共享

同一網路上的其他電腦可以訪問您的工具：

```bash
# 查看您的 IP 地址
ipconfig

# 其他電腦訪問
# http://YOUR_IP:8000
```

### 自訂埠號

編輯 `start_server.bat`，將 `8000` 改為其他埠號：

```batch
python -m http.server 9000  # 改為 9000
```

---

## 📚 相關資源

- [Tabulator.js 文檔](http://tabulator.info/) - 表格庫
- [Chart.js 文檔](https://www.chartjs.org/) - 圖表庫
- [SheetJS 文檔](https://sheetjs.com/) - Excel 解析

---

## 🎉 準備好了？

現在就試試看：

**Windows 使用者**：雙擊 `start_server.bat` 🚀

**macOS/Linux 使用者**：

```bash
cd protein
python3 -m http.server 8000
```

**或使用 Node.js**：

```bash
npm install && npm start
```

祝您使用愉快！ 🌟
