# 🏭 Industrial-Excel-System-Logic-Suite
> **工業級試算表邏輯建模與生產管理控制系統 (Industrial Engineering Excel Models)**

本專案庫包含 4 個基於工業工程（IE）與生產管理（Production Management）標準建構的 Excel 高階邏輯模組。專案擺脫傳統單純試算表計算思維，導入**參數解耦（Parameter Decoupling）**、**防呆驗證（Error-Trapping）**與**權限控管（RBAC）**架構，展示將工業控制與數據邏輯落地於 Excel 的系統化設計能力。

---

## 📌 核心模組一覽 (Core Modules)

### 1. SCADA 設備動態監控與階層式告警矩陣 (`01_SCADA_Alarm_Matrix.xlsx`)
* **系統痛點**：傳統單一門檻值容易造成「極端異常降級漏報」與告警遮蔽（Alarm Masking）。
* **邏輯架構**：建立 HH/H (High-High / High) 雙階層警報機制，整合 `IFS` 與 `AND`/`OR` 多條件布林邏輯算式。
* **工程設計**：採用參數解耦（$F$2:$G$3 絕對引用），將動態門檻值抽出至獨立參數區，實現全模組免修改算式即可動態調整告警條件。
  

### 2. EMS 時間電價稽核與四階降載預警模型 (`02_EMS_TOU_LoadShedding.xlsx`)
* **系統痛點**：離尖峰電價計算複雜，且超過契約容量時缺乏即時預警機制。
* **邏輯架構**：運用 `XLOOKUP` 動態比對時間電價（TOU Rate）進行電費精準稽核；透過需量比率計算觸發 Level 1~3 四階降載預警。
* **工程設計**：將契約容量與電費基準參數化，確保費率調整時只需更新參數表即可即時連動全廠稽核邏輯。

### 3. MES/QA SOP 規格驗證與 RBAC 權限控管 (`03_MES_SOP_RBAC_Validation.xlsx`)
* **系統痛點**：人為誤輸入超出 SOP 規格之數據，或未授權人員隨意更動生產參數。
* **邏輯架構**：結合 `XLOOKUP` 自動檢核 SOP 上下限規格（USL/LSL），並嵌入角色權限控制（RBAC）防呆機制。
* **工程設計**：利用「資料驗證（Data Validation）」與條件式格式化，自動攔截非法寫入、空值與越權操作，落實現場品質防呆。

### 4. TPM 設備六大損失與帕雷托 (Pareto) 數據分析儀表板 (`04_TPM_Pareto_Downtime_Analysis.xlsx`)
* **系統痛點**：設備停機日誌缺乏系統化歸因，導致改善資源無法精準投入。
* **邏輯架構**：依據 TPM 六大損失分類歸納數據，利用樞紐分析表（Pivot Table）設定「按某一欄位彙總的百分比（累計百分比）」。
* **工程設計**：建構雙軸柏拉圖（Pareto Chart），精確定位造成 80% 停機時間的 20% 關鍵故障源，並搭配互動式交叉篩選器（Slicer）提供跨機台/班別之多維度分析。

---

## 🛠️ 核心 Excel 技術棧與設計模式 (Tech Stack & Architecture)

| 技術類別 | 關鍵函數 / 功能 | 工業應用情境 |
| :--- | :--- | :--- |
| **進階邏輯算式** | `XLOOKUP`, `IFS`, `SWITCH`, `AND`/`OR` | SCADA 多階告警判定、MES 規格自動檢核、EMS 費率查表 |
| **架構設計** | 絕對引用 (`$`)、參數解耦 (Decoupling) | 動態門檻維護、契約容量設定、系統低耦合維護 |
| **數據防呆與權限** | Data Validation, RBAC, Error-Trapping | 未授權輸入攔截、空值異常處理、超標自動警報 |
| **數據分析與視覺化** | Pivot Table (% Running Total), Pareto Chart, Slicer | TPM 設備六大損失歸因、80/20 停機主因洞察儀表板 |

---

