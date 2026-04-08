# AutoCAD Grid Method Slope Calculator
**AutoCAD 「坵塊法」網格坡度與坡向計算工具 (V3)**

A fast, native C# AutoCAD .NET plugin for rigorous topography slope analysis and True-North aspect calculations using the authoritative Grid Method. 
本工具為 AutoCAD .NET 擴充功能，利用自動化「坵塊法」快速完成大範圍地形坡度運算、正北坡向分析，並一鍵產出完整圖例與數據報表。

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
1. 《水土保持技術規範》第二十五條 (Article 25).
2. 《山坡地範圍劃定及檢討變更作業要點》第七條 (Point 7).
3. 《山坡地土地可利用限度查定工作要點》第四點 (Point 4).

---

### 🚀 安裝與執行 (Install & Run)

**1. 修改與編譯 (Compile & Modify - Optional)**
若需自訂運算邏輯，請修改 `GridSlopeCalculatorV3.cs`，然後執行資料夾內的 `buildV3.bat`，即可自動編譯出全新的 DLL！
*Edit the `.cs` file if needed, then run `buildV3.bat` to instantly generate a custom `.dll`.*

**2. 載入 (Load Plugin)**
開啟 DWG 地形圖。在 AutoCAD 輸入指令 `NETLOAD`，並選取 `GridSlopeCalculatorV3.dll` 載入。
*In AutoCAD, use the `NETLOAD` command and select `GridSlopeCalculatorV3.dll`.*

**3. 執行 (Execute Command)**
在 AutoCAD 命令列輸入主程式指令：
*Type the command:*
```text
CalcGridSlopeCSV3
```

---

### 💻 使用流程 (Usage Workflow)

1. **參數輸入 (Input Parameters)**：輸入方格邊長 $L$ (例如: 25) 與 等高線間距 $\Delta h$ (例如: 1.0)。
2. **條件選取 (Select Objects)**：依據畫面提示，依序點擊：
   - 一個「方格 (Grid)」樣本 *(Auto-selects all grids)*
   - 一條「等高線 (Contour)」樣本 *(Auto-selects all contours)*
   - 圖面區的「專案總邊界 (Project Boundary)」 *(Filters calculation area)*
3. **輸出定位 (Set Anchor)**：在空白處點選一點，放置圖例與分析報表！

---

### 📊 成果輸出 (Generated Outputs)

* **網格視覺標示 (Visual Grid Markers):** 於每格網格內生成面積 ($m^2$)、坡度 (%) 以及精準度貼合絕對正北 (True North) 的下坡箭頭。
* **分析總表與圖例 (Native AutoCAD Tables):** 自動產出所有網格的彙整成果表、0~100% 級別參考圖例，以及一組具備視覺防呆功能的動態 3x3 指南針 (Compass Legend)。
* **匯出 CSV (CSV Data Export):** 支援 UTF-8 (BOM) 格式直接寫出分析矩陣，確保匯入 Excel 時中文字元完美顯示！
