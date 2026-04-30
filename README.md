***

# AutoCAD 「坵塊法」網格坡度與坡向計算工具 (Grid Method Slope Calculator)

本專案是一個專為 AutoCAD .NET 平台開發的擴充套件，旨在自動化執行水土保持計畫中「坵塊法」的大範圍地形坡度運算與坡向分析。本工具專注於符合台灣相關法規標準，提供使用者容易操作、結果精確的坡度分級與方向判定，並能一鍵產出完整的統計報表與原生 Excel 報表。

## 核心特色 (Features)

- **自動化網格劃分與運算**：支援自動化劃分 10m/25m 網格，並運用交點計算法則自動求得網格平均坡度。
- **智慧型方向內插 (k-NN IDW)**：對於等高線交點不足的平坦區域，系統能自動以反距離權重法，從周圍 8 個最近的有效網格內插地形方向。
- **即時圖面同步更新 (DOM-based Update)**：支援設計師手動在 AutoCAD 內修改交點數或方向文字，系統能自動掃描圖層變更、校正箭頭指向，並即時重繪網底與更新數據表。
- **原生 XLSX 報表匯出**：無須安裝 Excel，直接將分析數據匯出為帶有動態公式的 `.xlsx` 檔案。

## 專案結構 (Project Structure)

本專案採行「使用與開發分離」的設計：

| 檔案 / 資料夾 | 說明 |
|---------------|------|
| `GridMethodSlopeCalculator.dll` | 供一般使用者直接載入 AutoCAD 的編譯完成檔 |
| `GridMethodSlopeCalculator_Edit/` | 開發者專用目錄，內含原始碼 (`.cs`)、編譯腳本 (`build.bat`) 與相關 API 組件 |
| `UNBLOCK.bat` | 解除 Windows 安全鎖定（ZIP 下載後必須執行） |
| `README.md` | 使用說明書 |
| `DetailedInfo_細節說明.md` | 相關細節說明 |

## 系統需求 (Prerequisites)

- **環境**：AutoCAD 2021 或更新版本 (相容於 Civil 3D)
- **框架**：.NET Framework 4.x
- **編譯器**：若需自行編譯，需使用內建的 `csc.exe` (C# 5)

***
***

## 快速開始 (Getting Started)

### 載入外掛套件

> [!WARNING]
> **從 GitHub 下載 ZIP 的使用者請注意：** Windows 會自動封鎖從網路下載的 DLL 檔案，導致 AutoCAD `NETLOAD` 失敗（錯誤代碼 `0x80131515`）。  
> 請在載入前先雙擊執行 **`UNBLOCK.bat`** 以解除鎖定。  
> 若使用 `git clone` 方式取得專案則不受此影響。

1. 開啟 AutoCAD 並載入您的 DWG 地形圖及計畫範圍(可視情況選擇是否預先建立網格)。
2. 在命令列輸入 `NETLOAD`。
3. 選取本專案目錄下的 `GridMethodSlopeCalculator.dll`。

### 指令操作流程

本工具提供三個主要的指令操作流程：

> [!TIP]
> 建議依序使用指令，以達到最佳的使用者體驗與分析精準度。

1. **初始計算 (`GM1_Calc`)**
   啟動主要運算流程。系統會提示您建立方格、輸入等高線間距、選取圖面特徵，並最終產出視覺化的網格標示與統計摘要表。

   ```text
   GM1_Calc
   │
   ├─ 「方格是否已建立？」 Has grid already been built? [Y/N]
   │
   ├─ [N] ── 自動產生方格 (Auto Grid Generation)
   │   ├─ 1. 選擇計畫範圍 (Select project boundary)
   │   ├─ 2. 確認範圍選取 (Confirm boundary — highlight + Y/N)
   │   ├─ 3. 選擇方格尺寸 [25m / 10m] (default 25m)
   │   └─ 4. 自動建立 UCS 對齊方格，最佳化偏移防止邊界重疊
   │
   ├─ [Y] ── 輸入方格邊長 L (Enter grid side length)
   │
   ├─ 輸入等高線間距 Δh (Enter contour interval)
   ├─ 是否匯出 XLSX？ (Export to XLSX? [Y/N])
   │
   ├─ 選取方格/等高線/計畫範圍 (Select grids, contours, boundary)
   ├─ 指定報表插入點 (Pick table insertion point)
   │
   └─ 計算完成 (Process):
       ├─ 計算坡度與方向 (Calculate slope & direction)
       ├─ k-NN IDW 方向內插 (Interpolate missing directions)
       ├─ 繪製標註與箭頭 (Draw annotations & arrows)
       ├─ 產生統計摘要表 (Generate statistical summary tables)
       └─ 匯出 XLSX (Export to XLSX)
   ```
   
2. **手動微調與更新 (`GM2_Update`)**
   若您手動編輯了網格內的「交點數」或「方向文字」，請執行此指令。系統將會更新快取資料、重新產生坡度網底，並可選擇覆寫更新畫面中的統計表。

   ```text
   GM2_Update
   │
   ├─ 手動修改圖面 (User edits on the drawing):
   │   ├─ 修改交點數 (Edit intersection count)
   │   └─ 修改方向文字 (Edit direction text)
   │
   ├─ 執行指令 (Execute GM2_Update)
   │
   ├─ 讀取快取 (Read NOD Cache)
   ├─ 掃描圖面變更 (Scan DOM for text modifications)
   │   └─ 若發現方向變更 → 自動重繪中心箭頭 (Arrow auto-correction)
   │
   ├─ 重新產生坡度網底 (Regenerate slope hatch)
   │
   ├─ 詢問更新報表 (Regenerate Summary Table? [Y/N])
   │   ├─ [Y] ── 原地覆蓋新統計表 (Replace table in place)
   │   └─ [N] ── 保留舊表 (Keep existing table)
   │
   └─ 更新完成 (Update Complete)
       └─ 自動匯出遞增檔名之新 XLSX (Export new XLSX)
   ```

3. **獨立報表匯出 (`GM3_Export`)**
   當您僅需要將現有圖面上的分析結果輸出成報表，而不需要重新計算時，執行此指令即可直接生成 `.xlsx` 檔案。

   ```text
   GM3_Export
   │
   ├─ 執行指令 (Execute GM3_Export)
   │
   ├─ 讀取快取 (Read NOD Cache)
   ├─ 掃描現有圖面資料 (Scan current drawing grid data)
   │
   └─ 匯出完成 (Export Complete)
       └─ 不觸發重新運算，直接產生最新的 .xlsx 檔案
   ```

## 其他細節說明

###  關於坵塊法 (About the Grid Method)

**一、目的 (Purpose)**  
坵塊法主要用於有實測地形圖時的**坡度分析**，藉由量測並計算出特定方格內的平均坡度，作為**山坡地範圍劃定、檢討變更**及**土地可利用限度查定**時的法定基準。  
*The Grid Method calculates average grid slopes to legally delineate and review slopeland boundaries or determine slopeland utilization limits based on topographic maps.*

**二、計算方式 (Calculation Method)**
1. **劃設方格 (Draw a Grid):** 在實測地形圖上，每 10 公尺或 25 公尺畫一方格。
2. **計算交點 (Count Intersections):** 算出每方格各邊與等高線相交的交點數量總和 (**n** 值)。
3. **代入公式 (Apply Formula):** 利用專用公式求得坵塊內平均坡度：
   **S(%) = (n × π × Δh / 8L) × 100**
   * **S**: 方格內平均坡度 (%) / *Average slope*
   * **n**: 等高線與方格四邊交點總數 / *Total intersections*
   * **Δh**: 等高線間距 (公尺) / *Contour interval*
   * **L**: 方格邊長 (公尺) / *Grid side length*

**三、法規出處 (Regulatory Sources)**  
根據《110水土保持相關法規彙編》 (Taiwan Soil and Water Conservation Regulations):
1. 《水土保持技術規範》第二十五條 
2. 《山坡地範圍劃定及檢討變更作業要點》第七條
3. 《山坡地土地可利用限度查定工作要點》第四點

###  開發與編譯 (Development & Compile)

本專案採行「使用與開發分離」結構設計：
- **一般使用者**：直接使用外層的 `GridMethodSlopeCalculator.dll` 載入 AutoCAD 即可。
- **開發者**：若需修改功能，請進入 `GridMethodSlopeCalculator_Edit` 資料夾。該資料夾包含完整的原始碼、AutoCAD 參考庫與編譯環境。

修改原始碼後，請在 `GridMethodSlopeCalculator_Edit` 資料夾內執行 `build.bat`，系統會自動編譯並覆蓋原有的 DLL 檔案。  
*If you wish to edit the source code, navigate to the `GridMethodSlopeCalculator_Edit` folder and run `build.bat` to recompile the plugin.*

```text
> cd GridMethodSlopeCalculator_Edit
> build.bat
Compiling Grid Method Slope Calculator...
SUCCESS: GridMethodSlopeCalculator.dll created.
```



