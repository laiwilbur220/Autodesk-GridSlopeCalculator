# AutoCAD Grid Method Slope Calculator
**AutoCAD 「坵塊法」網格坡度與坡向計算工具 (V6)**

A production-ready C# AutoCAD .NET plugin for rigorous topography slope analysis and True-North aspect calculations using the authoritative Grid Method.  
本工具為 AutoCAD .NET 擴充功能，利用自動化「坵塊法」快速完成大範圍地形坡度運算、正北坡向分析，並一鍵產出完整統計報表與數據報表。

---

###  關於坵塊法 (About the Grid Method)

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

###  安裝與執行 (Install & Run)

**1. 載入 (Load Plugin)**  
開啟 DWG 地形圖。在 AutoCAD 輸入指令 `NETLOAD`，並選取 `GridMethodSlopeCalculator.dll` 檔案。  
*In AutoCAD, use `NETLOAD` and select the `GridMethodSlopeCalculator.dll` file.*

**2. 執行 (Execute Command)**  
在 AutoCAD 命令列輸入：  
*Type the commands:*
```text
GM1_Calc
GM2_Update
GM3_Export
```

---

###  使用流程 (Usage Workflow)

```
GM1_Calc
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

**使用者手動修改後更新流程 (User-Initiated Update Workflow):**
若使用者在圖面上**手動修改了交點數或坡向文字** (Manual Edits)：
1. 執行 `GM2_Update`
2. 系統讀取 NOD 暫存數據，掃描方格內的使用者文字 (DOM Scraping by Layer)
3. 自動偵測方向文字變更並重繪對應箭頭 (Arrow auto-correction)
4. 自動消除警告 (NOMUTT) 並重新產生坡度網底 (Hatch)
5. 問答 `Regenerate Summary Table? [Y/N]`，若選擇 `Y` 則自動在原地覆蓋新表格
6. 自動遞增檔名 (Increment filename) 防止 Excel 檔案鎖死，並輸出新 XLSX

**獨立匯出 (Standalone Export):**
若只需將圖面現有數據輸出至 Excel，無須重新計算或更新畫面：
1. 執行 `GM3_Export`
2. 系統讀取 NOD 參數並掃描畫面資料，直接輸出最新的 `.xlsx` 報表

---

###  成果輸出 (Generated Outputs)

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

---

###  檔案結構 (File Structure)

| File / Folder | Purpose |
|---------------|---------|
| `README.md` | 本使用說明書 (Documentation) |
| `GridMethodSlopeCalculator.dll` | 供使用者直接載入的編譯完成檔 (Ready-to-use plugin) |
|  `GridMethodSlopeCalculator_Edit/` | **原始碼編輯區 (Developer workspace)** |
| ├── `GridMethodSlopeCalculator.cs` | 主程式 C# 原始碼 (Source code) |
| ├── `build.bat` | 一鍵自動編譯腳本 (Build script) |
| └── `acmgd.dll`... | AutoCAD .NET API 編譯專用參考組件 |


###  系統需求 (Requirements)

* **AutoCAD** 2021 或以上 (Civil 3D 相容)
* **.NET Framework** 4.x
* **編譯器**: .NET Framework 內建 `csc.exe` (C# 5)
