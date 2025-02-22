# AutoReturn 自動退款系統

## 專案介紹
AutoReturn 是一個基於 .NET 8 Windows Forms 開發的自動化退款處理系統。該系統可以批量處理 Excel 文件中的訂單退款請求，並自動與資料庫和 API 進行交互，大大提高了退款處理的效率。

## 主要功能
- 📑 Excel 訂單資料讀取與解析
- 🔄 批量退款處理
- 📊 資料庫訂單驗證
- 🔒 安全的 API 整合
- 📝 完整的日誌記錄
- ⚡ 自動重試機制
- 🛡️ 異常處理和錯誤提示

## 系統需求
- .NET 8.0 SDK
- Windows 10/11
- MySQL 資料庫
- 有效的 API 存取權限

## 開發環境設置
1. 安裝 Visual Studio 2022 (17.8 或更高版本)
2. 安裝 .NET 8.0 SDK
3. 安裝以下 NuGet 套件：
   - EPPlus
   - MySql.Data
   - System.Text.Json

## 安裝說明
1. 從 Release 頁面下載最新版本
2. 解壓縮檔案到本地目錄
3. 配置 `setting.json` 檔案：
   ```json
   {
     "ConnectionString": "your_database_connection_string",
     "ApiUrl": "your_api_url",
     "Token": "your_token"
   }
   ```

## 使用方法
1. 啟動應用程序
2. 點擊選擇 Excel 文件（檔案必須包含「訂單號」欄位）
3. 點擊開始處理按鈕
4. 系統將自動：
   - 讀取 Excel 中的訂單資料
   - 驗證訂單在資料庫中的狀態
   - 提交退款請求
   - 記錄處理結果

## 技術特點
- 使用 .NET 8 最新特性
- 現代化的 Windows Forms 應用程序
- 非同步操作處理
- HTTP/3 支援
- 原生 JSON 序列化
- 改進的垃圾回收機制

## 專案結構
```
AutoReturn/
├── Forms/
│   └── Form1.cs
├── Models/
│   └── Settings.cs
├── Services/
│   ├── ExcelService.cs
│   ├── DatabaseService.cs
│   └── ApiService.cs
├── Utils/
│   └── LogHelper.cs
└── Program.cs
```

## 錯誤處理
- 資料庫鎖定時自動重試（最多3次）
- 網絡超時自動重試
- 詳細的錯誤提示和日誌記錄

## 日誌記錄
系統會自動在應用程序目錄下的 Logs 資料夾中生成日誌文件：
- 檔名：`refund_log_YYYYMMDD.txt`
- 格式：`[時間戳] [日誌級別] 訂單編號: {訂單號} - {訊息內容}`

## 建置說明
```bash
# 恢復 NuGet 套件
dotnet restore

# 建置專案
dotnet build

# 發布專案
dotnet publish -c Release -r win-x64 --self-contained true
```

## 貢獻指南
歡迎提交 Pull Request 或 Issue。在提交之前，請確保：
1. 代碼遵循現有的風格規範
2. 新功能包含適當的測試
3. 更新相關文檔

## 授權協議
本專案採用 MIT 授權協議。詳細內容請參見 [LICENSE](LICENSE) 文件。

## 聯絡方式
如有任何問題或建議，請提交 Issue 或聯繫專案維護者。

## 更新日誌
### v1.0.0 (2024-02-22)
- 初始版本發布
- 基於 .NET 8 開發
- 實現基本的退款處理功能
- 添加日誌記錄系統
- 實現自動重試機制
