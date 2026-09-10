# 🏭 Industrial-Excel-System-Logic-Suite
> **工業級試算表邏輯建模與生產管理控制系統 (Industrial Engineering Excel Models)**

本專案庫包含 4 個基於工業工程（IE）與生產管理（Production Management）標準建構的 Excel 高階邏輯模組。專案擺脫傳統單純試算表計算思維，導入**參數解耦（Parameter Decoupling）**、**防呆驗證（Error-Trapping）**與**權限控管（RBAC）**架構，展示將工業控制與數據邏輯落地於 Excel 的系統化設計能力。

---

## 📌 核心模組與系統實作畫面 (Core Modules & System Logic)

### 1. SCADA 設備動態監控與階層式告警矩陣

![SCADA 設備動態監控與警報邏輯模組](./備運轉動態監控與警報邏輯模組.png)

* **系統痛點**：傳統單一門檻值容易造成「極端異常降級漏報」與告警遮蔽（Alarm Masking）。
* **邏輯架構**：建立 HH/H (High-High / High) 雙階層警報機制，整合 `IFS` 與 `AND`/`OR` 多條件布林邏輯算式。
* **工程設計**：採用參數解耦（`$E$2:$G$3` 絕對引用），將動態門檻值抽出至獨立參數區，實現全模組免修改算式即可動態調整告警條件。

---

### 2. EMS 時間電價稽核與四階降載預警模型

![EMS 廠房時間電價計算與契約容量超標預警](./廠房時間電價計算與契約容量超標預警.png)

* **系統痛點**：離尖峰電價計算複雜，且超過契約容量時缺乏即時預警機制。
* **邏輯架構**：運用 `XLOOKUP` 動態比對時間電價（TOU Rate）進行電費精準稽核；透過需量比率計算觸發 Level 1~3 四階降載預警。
* **工程設計**：將契約容量與電費基準參數化，確保費率調整時只需更新右側對照表即可即時連動全廠稽核邏輯。

---

### 3. MES/QA SOP 規格驗證與 RBAC 權限控管

#### 前台輸入與 QA 診斷校對
![MES QA 邏輯校對結果](./MES%20後台%20RBAC%20權限與%20SOP%20規格QA%20驗證模組-1.png)

#### 後台 SOP 主規格與角色授權表 (Master Table)
![MES SOP 規格與 RBAC 權限主表](./MES%20後台%20RBAC%20權限與%20SOP%20規格QA%20驗證模組-2.png)

* **系統痛點**：人為誤輸入超出 SOP 規格之數據，或未授權人員隨意更動生產參數。
* **邏輯架構**：結合 `XLOOKUP` 自動檢核 SOP 上下限規格（USL/LSL），並嵌入角色權限控制（RBAC）防呆機制。
* **工程設計**：利用進階巢狀函數，自動攔截「非法寫入」、「空值」與「越權操作」，明確輸出診斷訊息（如：`錯誤：權限不足 (需 Admin)`、`錯誤：超出 SOP 規格`）。

---

### 4. TPM 設備六大損失與帕雷托 (Pareto) 數據分析儀表板

#### 原始停機日誌資料表 (Raw Data Log)
![TPM 設備損失分類原始資料表](./TPM%20設備損失分類.png)

#### 帕雷托樞紐分析表與互動式戰情儀表板 (Dashboard)
![TPM 設備損失分類樞紐分析表與雙軸圖](./TPM%20設備損失分類樞紐分析表.png)

* **系統痛點**：設備停機日誌缺乏系統化歸因，導致改善資源無法精準投入。
* **邏輯架構**：依據 TPM 六大損失分類歸納數據，利用樞紐分析表（Pivot Table）設定「按某一欄位彙總的百分比（累計百分比）」。
* **工程設計**：建構雙軸柏拉圖（Pareto Chart），精確定位造成 80% 停機時間的 20% 關鍵故障源（如：設備故障占 72.5%），並搭配互動式交叉篩選器（Slicer）提供跨機台與班別之多維度分析。

---

## 🛠️ 核心 Excel 技術棧與設計模式 (Tech Stack & Architecture)

| 技術類別 | 關鍵函數 / 功能 | 工業應用情境 |
| :--- | :--- | :--- |
| **進階邏輯算式** | `XLOOKUP`, `IFS`, `SWITCH`, `AND`/`OR` | SCADA 多階告警判定、MES 規格自動檢核、EMS 費率查表 |
| **架構設計** | 絕對引用 (`$`)、參數解耦 (Decoupling) | 動態門檻維護、契約容量設定、系統低耦合維護 |
| **數據防呆與權限** | Data Validation, RBAC, Error-Trapping | 未授權輸入攔截、空值異常處理、超標自動警報 |
| **數據分析與視覺化** | Pivot Table (% Running Total), Pareto Chart, Slicer | TPM 設備六大損失歸因、80/20 停機主因洞察儀表板 |

---

## 📂 專案檔案結構 (Repository Structure)

```text
.
├── images/
│   ├── 備運轉動態監控與警報邏輯模組.png
│   ├── 廠房時間電價計算與契約容量超標預警.png
│   ├── MES 後台 RBAC 權限與 SOP 規格QA 驗證模組-1.png
│   ├── MES 後台 RBAC 權限與 SOP 規格QA 驗證模組-2.png
│   ├── TPM 設備損失分類.png
│   └── TPM 設備損失分類樞紐分析表.png
├── 01_SCADA_Alarm_Matrix.xlsx            # SCADA 告警矩陣與參數解耦模組
├── 02_EMS_TOU_LoadShedding.xlsx          # EMS 時間電價稽核與降載預警模組
├── 03_MES_SOP_RBAC_Validation.xlsx      # MES SOP 規格驗證與 RBAC 防呆模組
├── 04_TPM_Pareto_Downtime_Analysis.xlsx    # TPM 六大損失與帕雷托樞紐分析表
└── README.md                             # 專案說明文件
