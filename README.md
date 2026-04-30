# AutoCAD Grid Method Slope Calculator
**AutoCAD 「坵塊法」網格坡度與坡向計算工具 (V6)**

A production-ready C# AutoCAD .NET plugin for rigorous topography slope analysis and True-North aspect calculations using the authoritative Grid Method — featuring **k-NN IDW direction interpolation**, **DOM-based manual override scraping**, **arrow auto-correction on update**, and **native XLSX exports**.  
本工具為 AutoCAD .NET 擴充功能，利用自動化「坵塊法」快速完成大範圍地形坡度運算、正北坡向分析，並一鍵產出完整統計報表與數據報表。V6 版新增 **k-NN IDW 方向內插**、**箭頭自動校正**與**統計摘要表**功能。

---

### 📌 關於坵塊法 (About the Grid Method)

**一、目的 (Purpose)**  
坵塊法主要用於有實測地形圖時的**坡度分析**，藉由量測並計算出特定方格內的平均坡度，作為**山坡地範圍劃定、檢討變更**及**土地可利用限度查定**時的法定基準。  
*The Grid Method calculates average grid slopes to legally delineate and review slopeland boundaries or determine slopeland utilization limits based on topographic maps.*

**二、計算方式 (Calculation Method)**
1. **劃設方格 (Draw a Grid):** 在實測地形圖上，每 10 公尺或 25 公尺畫一方格。
2. **計算交點 (Count Intersections):** 算出每方格各邊與等高線相交的交點數量總和 (**n** 值)。
3. **代入公式 (Apply Formula):** 利用專用公式求得坵塊內平均坡度：**S(%) = (n × π × Δh / 8L) × 100**
   * **S**: 方格內平均坡度 (%) / *Average slope*
   * **n**: 等高線與方格四邊交點總數 / *Total intersections*
   * **Δh**: 等高線間距 (公尺) / *Contour interval*
   * **L**: 方格邊長 (公尺) / *Grid side length*

**三、法規出處 (Regulatory Sources)**  
根據《110水土保持相關法規彙編》 (Taiwan Soil and Water Conservation Regulations):
1. 《水土保持技術規範》第二十五條 
2. 《山坡地範圍劃定及檢討變更作業要點》第七條
3. 《山坡地土地可利用限度查定工作要點》第四點

---

### 🚀 安裝與執行 (Install & Run)

**1. 載入 (Load Plugin)**  
開啟 DWG 地形圖。在 AutoCAD 輸入指令 `NETLOAD`，並選取 `GridSlopeCalculatorV6_1.dll` 檔案。  
*In AutoCAD, use `NETLOAD` and select the `GridSlopeCalculatorV6_1.dll` file.*

**2. 執行 (Execute Command)**  
在 AutoCAD 命令列輸入：  
*Type the commands:*
```text
CalcGridSlopeCSV6
UpdateGridSlopeCSV6
```

---

### 💻 使用流程 (Usage Workflow)

```
CalcGridSlopeCSV6
│
├─ 「方格是否已建立？」 Has grid already been built? [Y/N]
│
├─ [N] ── 自動產生方格 (Auto Grid Generation)
│   ├─ 1. 選擇計畫範圍 (Select project boundary)
│   ├─ 2. 確認範圍選取 (Confirm boundary — highlight + Y/N)
│   ├─ 3. 選擇方格尺寸 [25m / 10m] (default 25m)
│   └─ 4. 自動建立 UCS 對齊方格，最佳化偏移防止邊界重疊
│        → "Generated 42 grid cells (25m) on layer [GRID]"
│
├─ [Y] ── 輸入方格邊長 L (Enter grid side length)
│
├─ 輸入等高線間距 Δh (Enter contour interval)
├─ 是否匯出 XLSX？ (Export to XLSX? [Y/N])
│
├─ 選取方格樣本 (Auto-select grids) 或自動選取已產生方格
├─ 選取等高線樣本 (Select contour samples)
├─ 確認計畫範圍 (Confirm boundary)
├─ 指定報表插入點 (Pick table insertion point)
│
└─ 計算完成 (Process):
    ├─ 計算坡度與方向 (Calculate slope & direction)
    ├─ k-NN IDW 方向內插 (Interpolate missing directions)
    ├─ 繪製標註與箭頭 (Draw annotations & arrows)
    ├─ 產生統計摘要表 (Generate statistical summary tables)
    └─ 匯出 XLSX (Export to XLSX)
```

**更新流程 (Update Workflow):**
若使用者在圖面上**手動修改了交點數或坡向文字** (Manual Edits)：
1. 執行 `UpdateGridSlopeCSV6`
2. 系統讀取 NOD 暫存數據，掃描方格內的使用者文字 (DOM Scraping by Layer)
3. 自動偵測方向文字變更並重繪對應箭頭 (Arrow auto-correction)
4. 自動消除警告 (NOMUTT) 並重新產生坡度網底 (Hatch)
5. 問答 `Regenerate Summary Table? [Y/N]`，若選擇 `Y` 則自動在原地覆蓋新表格
6. 自動遞增檔名 (Increment filename) 防止 Excel 檔案鎖死，並輸出新 XLSX

---

### 📊 成果輸出 (Generated Outputs)

* **網格視覺標示 (Visual Grid Markers):**  
  每格網格內標示方向 (top)、坡度 (%)、級別分類、交點數 n= 及面積 A (bottom)。  
  使用精緻的 37 頂點輪廓箭頭標示下坡方向，箭頭置於方格正中心。

* **方向內插 (Direction Interpolation):**  
  等高線交點不足 3 個的方格，自動以 k-NN (k=8) 反距離權重法 (IDW, w=1/d²) 從最近的有效鄰居方格內插方向。適用於大面積平坦地形 (如 20 公頃、平均坡度 2% 的平台)。  
  *Grid cells with < 3 contour intersections automatically receive an interpolated direction from the 8 nearest valid neighbors via Inverse Distance Weighting.*

* **統計摘要表 (Statistical Summary Tables):**  
  自動產出兩張 AutoCAD 原生表格：  
  1. **坡度分級統計** — 七級坡分類的方格數、面積與百分比，含合計與加權平均坡度。  
  2. **坡向分佈統計** — 八方位的方格數、面積與百分比，含合計與主要坡向。

* **分析總表 (Grid Data Summary Table):**  
  每格網格的 ID、交點數、坡度、分級、方向、面積彙整表。更新時支援透過 ObjectID 追蹤，原地自動替換舊表。

* **原生 XLSX 匯出 (Native XLSX Data Export):**  
  使用 `System.IO.Packaging` 零套件直接生成 OpenXML `(.xlsx)`。支援中文，且直接在 Excel 中夾帶動態公式，每次開啟即自動重新計算 (FullCalcOnLoad)。

---

### 🆕 V6 新增功能 (What's New in V6)

| Feature | Description |
|---------|-------------|
| **k-NN IDW 方向內插** | 等高線不足的方格自動從 8 個最近有效鄰居內插方向 |
| **箭頭置中** | 方向箭頭改為方格正中心，取代舊版上偏位置 |
| **箭頭自動校正** | `UpdateGridSlopeCSV6` 偵測方向文字變更後自動刪除舊箭頭並重繪 |
| **Layer-based 方向抓取** | 更新指令改用圖層篩選 (`Grid_Outputs_DirText`) 提升文字識別可靠性 |
| **統計摘要表** | 新增坡度分級 + 坡向分佈兩張 AutoCAD 表格，取代舊版圖例與指南針 |
| **獨立 NOD Key** | V6 使用 `GridSlopeParamsV6`，與 V5 資料互不干擾 |

---

### 🛠️ 開發與編譯 (Development & Compile)

若是修改了 `.cs` 原始碼，請執行 `buildV6.bat`，系統會自動編譯更新 DLL。  
*If you edit the source code, run `buildV6.bat` to recompile the plugin automatically.*

```text
> buildV6.bat
Compiling Civil 3D Grid Slope Tool V6_1...
SUCCESS: GridSlopeCalculatorV6_1.dll created.
AutoCAD commands: CalcGridSlopeCSV6, UpdateGridSlopeCSV6
```

---

### 📁 檔案結構 (File Structure)

| File | Purpose |
|------|---------|
| `GridSlopeCalculatorV6_1.cs` | V6 主程式原始碼 (Main source code — latest) |
| `buildV6.bat` | V6 編譯腳本 (Build script) |
| `GridSlopeCalculatorV6_1.dll` | 編譯輸出 (Compiled plugin) |
| `V5/GridSlopeCalculatorV5.cs` | V5 存檔 (Archived V5 source) |
| `acmgd.dll` / `acdbmgd.dll` / `accoremgd.dll` | AutoCAD .NET API 參考組件 |

**版本命名規則 (Version Naming):**  
`V6_1` → `V6_2` → `V6_3` ... 每次功能迭代遞增子版本號。

---

### ⚙️ 系統需求 (Requirements)

* **AutoCAD** 2021 或以上 (Civil 3D 相容)
* **.NET Framework** 4.x
* **編譯器**: .NET Framework 內建 `csc.exe` (C# 5)
