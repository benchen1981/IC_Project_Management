# 快速參考 - Office 365 專案管理方案

## 📋 專案資訊

**專案名稱**: I&C Project Management  
**Microsoft 365 帳號**: benchen1981@smail.nchu.edu.tw  
**建立日期**: 2025-12-30  
**狀態**: 文件完成，待執行實施

---

## 📁 文件結構

```
I&C Project Management/
├── README.md                          # 專案概覽
├── CHANGE_LOG.md                      # 變更記錄
├── EXECUTION_PLAN.md                  # ⭐ 執行計畫（從這裡開始）
├── EXCEL_DASHBOARD_GUIDE.md           # Excel 建立指南
├── USER_GUIDE.md                      # 完整使用手冊
├── QUICK_REFERENCE.md                 # 本文件
├── .gitignore                         # Git 忽略規則
└── openspec/
    ├── prompts/
    │   └── 02-office365_project_management_template.md  # 需求文件
    └── changes/
        ├── archive_01_project_initialization.md         # Archive 01
        └── 02-office365_project_management_template.md  # Archive 02
```

---

## 🚀 快速開始（3 步驟）

### Step 1: 閱讀文件（15 分鐘）

1. **EXECUTION_PLAN.md** - 了解整體實施步驟
2. **USER_GUIDE.md** - 第 2 章「快速開始指南」

### Step 2: 建立 Excel 模板（60 分鐘）

1. 開啟 **EXCEL_DASHBOARD_GUIDE.md**
2. 依照指南建立 `Dashboard.xlsx`
3. 建立其他 Excel 模板

### Step 3: 執行實施（4.5 小時）

1. 依照 **EXECUTION_PLAN.md** 執行
2. 階段 1-6 逐步完成
3. 勾選檢查清單

---

## 🎯 5 大核心流程

| 流程              | 主要工具                  | 時間         | 關鍵活動                         |
| ----------------- | ------------------------- | ------------ | -------------------------------- |
| **1. 專案啟動**   | Teams + Planner + OneNote | 1.5 小時     | 建立團隊、撰寫章程、建立看板     |
| **2. 任務分配**   | Planner + Excel           | 2 小時       | 排序優先級、分配負責人、設定日期 |
| **3. 日常追蹤**   | To Do + Planner           | 每日 15 分鐘 | 站立會議、更新狀態、記錄進度     |
| **4. 週檢討**     | Excel + Teams             | 每週 60 分鐘 | 檢視進度、分析問題、調整計畫     |
| **5. 結案總報告** | Excel + OneNote           | 3 小時       | 統計成果、經驗總結、文件歸檔     |

---

## 🛠️ 工具職責速查

### Microsoft Planner
- ✅ 團隊任務看板（Kanban）
- ✅ 拖曳改變狀態
- ✅ 任務拆解（步驟功能）
- ✅ 負責人指派
- ✅ 截止日期管理

### Microsoft To Do
- ✅ 個人任務清單
- ✅ 我的一天規劃
- ✅ 提醒設定
- ✅ 與 Planner 自動同步

### Microsoft OneNote
- ✅ 專案文件
- ✅ 會議紀錄
- ✅ 設計文件
- ✅ 知識庫

### Microsoft Excel
- ✅ Dashboard 儀表板
- ✅ 甘特圖
- ✅ 進度追蹤
- ✅ 統計分析
- ✅ 週報生成

### Microsoft Teams
- ✅ 團隊溝通中樞
- ✅ 整合所有工具
- ✅ 檔案共享
- ✅ 會議管理

### Microsoft Outlook
- ✅ 行事曆同步
- ✅ 任務提醒
- ✅ 郵件通知

---

## ✅ 11 項功能需求檢查清單

- [x] **需求 1**: Kanban 看板，拖曳改變狀態
- [x] **需求 2**: 任務拆解，記錄目標與需求
- [x] **需求 3**: 分配負責人、截止日期、設定提醒
- [x] **需求 4**: 管理每日任務，更新狀態，會議紀錄
- [x] **需求 5**: Emoji 標籤表示優先級、進度
- [x] **需求 6**: 通知提醒、截止日、重複任務
- [x] **需求 7**: Dashboard Sheet，甘特圖，進度追蹤
- [x] **需求 8**: 完成狀態統計，甘特圖分析，瓶頸識別
- [x] **需求 9**: 條件格式、圖表、公式視覺化
- [x] **需求 10**: 匯出、共享、整合功能
- [x] **需求 11**: 進階視覺化、大量圖表

---

## 📊 Excel Dashboard 結構

### Dashboard.xlsx（5 個 Sheets）

| Sheet        | 用途       | 主要內容                      |
| ------------ | ---------- | ----------------------------- |
| **專案總覽** | 儀表板     | 4 個 KPI 卡片 + 4 個圖表      |
| **任務清單** | 資料管理   | 10 欄位 + 條件格式 + 資料驗證 |
| **甘特圖**   | 進度視覺化 | 條件格式甘特圖 + 里程碑       |
| **週報**     | 週報生成   | 自動統計 + 圖表 + 手動輸入區  |
| **統計分析** | 數據分析   | 6+ 圖表 + 樞紐分析表          |

### 其他 Excel 模板

- **Task_Tracker.xlsx** - 個人任務追蹤
- **Weekly_Report_Template.xlsx** - 週報模板

---

## 🎨 Planner 標籤系統

| 標籤       | 顏色 | 用途         |
| ---------- | ---- | ------------ |
| 🔴 高優先級 | 紅色 | 緊急且重要   |
| 🟡 中優先級 | 黃色 | 重要但不緊急 |
| 🟢 低優先級 | 綠色 | 可延後處理   |
| 💼 管理任務 | 藍色 | 管理類工作   |
| 🔧 技術任務 | 紫色 | 技術類工作   |
| 📄 文件任務 | 橙色 | 文件撰寫     |

---

## 📅 每日/每週工作流程

### 每日流程（15 分鐘）

**早上 9:00（5 分鐘）**
1. 開啟 To Do
2. 規劃「我的一天」（3-5 個任務）
3. 參加每日站立會議

**工作中（隨時）**
1. 完成任務時更新 Planner 狀態
2. 在 OneNote 記錄重要資訊

**下班前（10 分鐘）**
1. 更新 Excel 進度
2. 規劃明日任務

### 每週流程（60 分鐘）

**週五下午 3:00**
1. 週檢討會議（60 分鐘）
2. 更新 Excel 週報（15 分鐘）
3. 調整 Planner 任務（10 分鐘）
4. 發送週報（5 分鐘）

---

## 🔑 關鍵公式參考

### Excel 常用公式

```excel
// 整體進度
=COUNTIF(任務清單!F:F,"已完成")/COUNTA(任務清單!F:F)*100

// 逾期任務數
=COUNTIFS(任務清單!E:E,"<"&TODAY(),任務清單!F:F,"<>已完成")

// 本週完成任務
=COUNTIFS(任務清單!G:G,">="&TODAY()-7,任務清單!F:F,"已完成")

// 甘特圖條件格式
=AND($D2<=E$1,E$1<=$E2)
```

---

## 📞 常見問題快速解答

### Q: Planner 任務沒同步到 To Do？
**A**: 重新整理 To Do，或登出後重新登入

### Q: Excel 公式錯誤？
**A**: 檢查資料來源範圍和欄位名稱

### Q: 如何刪除任務？
**A**: 建議移至「已封存」而非刪除

### Q: 如何匯出 Planner？
**A**: 點擊「...」> 匯出計畫至 Excel

### Q: Teams 會議錄影在哪？
**A**: OneDrive 或「會議紀錄」頻道的檔案

---

## 📚 文件閱讀順序

### 第一次使用（建議順序）

1. **QUICK_REFERENCE.md** (本文件) - 5 分鐘
2. **EXECUTION_PLAN.md** - 15 分鐘
3. **USER_GUIDE.md** 第 1-2 章 - 10 分鐘
4. **EXCEL_DASHBOARD_GUIDE.md** - 依需要參考

### 日常使用

- **USER_GUIDE.md** 第 9 章 - 流程操作指南
- **EXCEL_DASHBOARD_GUIDE.md** - Excel 操作參考
- **USER_GUIDE.md** 第 10 章 - 常見問題

### 進階學習

- **USER_GUIDE.md** 第 11 章 - 最佳實踐
- **USER_GUIDE.md** 第 12 章 - 疑難排解
- **openspec/prompts/02-office365_project_management_template.md** - 完整需求規格

---

## ⏱️ 時間規劃

### 初次設定（4.5 小時）

| 階段   | 時間    | 活動                  |
| ------ | ------- | --------------------- |
| 階段 1 | 30 分鐘 | 環境準備與設定        |
| 階段 2 | 45 分鐘 | Planner 看板設計      |
| 階段 3 | 60 分鐘 | Excel Dashboard 建立  |
| 階段 4 | 30 分鐘 | To Do 與 Outlook 整合 |
| 階段 5 | 45 分鐘 | 整合與自動化          |
| 階段 6 | 60 分鐘 | 文件撰寫              |

### 日常維護

- **每日**: 15 分鐘（站立會議 + 狀態更新）
- **每週**: 90 分鐘（週檢討 + 週報）
- **每月**: 2 小時（月度檢討 + 調整）

---

## 🎯 下一步行動

### 立即可做

1. ✅ 閱讀本快速參考（已完成）
2. ⏳ 閱讀 EXECUTION_PLAN.md
3. ⏳ 建立 Excel Dashboard

### 需要登入 Microsoft 365

1. ⏳ 登入 https://www.office.com
2. ⏳ 建立 Teams 團隊
3. ⏳ 設定 Planner 看板
4. ⏳ 建立 OneNote 筆記本

### 驗證與優化

1. ⏳ 測試工作流程
2. ⏳ 團隊培訓
3. ⏳ 收集反饋並調整

---

## 📖 更多資源

### 官方文件

- [Microsoft Planner 說明](https://support.microsoft.com/planner)
- [Microsoft To Do 說明](https://support.microsoft.com/todo)
- [Microsoft Teams 說明](https://support.microsoft.com/teams)
- [Excel 函數參考](https://support.microsoft.com/excel)

### 內部文件

- `USER_GUIDE.md` - 完整使用手冊（12 章節，33KB）
- `EXECUTION_PLAN.md` - 執行計畫（6 階段，18KB）
- `EXCEL_DASHBOARD_GUIDE.md` - Excel 指南（11KB）

---

## ✨ 成功關鍵

### 3 個最重要的習慣

1. **每日更新** - 每天花 15 分鐘更新任務狀態
2. **週檢討** - 每週五檢視進度並調整計畫
3. **文件記錄** - 重要決策和會議記錄至 OneNote

### 3 個常見陷阱

1. ❌ 任務狀態長時間不更新
2. ❌ 會議沒有記錄行動項目
3. ❌ Excel 數據與 Planner 不同步

### 3 個最佳實踐

1. ✅ 使用「我的一天」專注於重要任務
2. ✅ 定期檢視 Dashboard 數據做決策
3. ✅ 在 Teams 保持透明溝通

---

**版本**: 1.0  
**最後更新**: 2025-12-30  
**維護者**: 專案經理

**提示**: 將本文件加入書籤，作為日常參考！
