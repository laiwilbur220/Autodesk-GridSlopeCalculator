# AutoCAD Grid Slope & Aspect Calculator (V3)
*(Scroll down for Traditional Chinese version / 向下捲動以查看繁體中文版本)*

A powerful, native C# AutoCAD .NET automation utility designed to rapidly evaluate engineering contour slopes, calculate True-North directional azimuths, and automatically generate comprehensive AutoCAD-native geometric visualizations, Summary Tables, and Compass Legends.

## 🌟 Key Features

- **Area-Weighted Terrain Analysis**: Extracts raw intersection parameters explicitly measuring overlapping boolean region areas to evaluate the physical geographic mean slope over overlapping non-uniform lot boundaries.
- **True North Azimuth Calculus**: Generates downhill geographic Aspect Vectors evaluated across planar Least-Squares elevation vectors against True Absolute World North (WCS).
- **Dynamic Geometric Visualizations**: Automatically maps geometric compass legends and dynamic internal arrow layouts specifically aligned with your User Coordinate System (UCS) while strictly orienting direction logic mapped from Absolute Gravity Flow vectors.
- **Native Table Generation**: Produces dynamic Summary Layout matrices and explicit mathematically-configured Statistical Arrays detailing grids by ID, Slopes, Directions, and Classifications.
- **Export to CSV**: Streams generated data matrices directly into an integrated CSV export properly encoding pure UTF-8 BOM ensuring explicit TrueType Chinese String compatibility natively into Microsoft Excel pipelines.

## 🛠 Prerequisites

No complex Civil 3D mapping extensions are required! This utility runs directly natively upon baseline Core AutoCAD arrays utilizing pure underlying DB/Geometry frameworks!
- Any Modern version of standard **Autodesk AutoCAD**.
- Base `.NET` runtime framework matching your target AutoCAD dependencies (usually `.NET Framework 4.7+` depending on AutoCAD version constraints).

## 🚀 Installation & Loading

1. **Compile the DLL**: If you wish to build the DLL manually, you can execute the localized `buildV3.bat` script provided. (Ensure your `acmgd.dll`, `acdbmgd.dll` library references are routed successfully dependent on your explicit AutoCAD framework directory constraints).
2. **Launch AutoCAD**: Open your relevant mapping project `DWG`.
3. **NETLOAD**: In the AutoCAD command prompt, type explicitly:
   ```text
   NETLOAD
   ```
4. Map the browser exactly to the generated `./GridSlopeCalculatorV3.dll` compiled application and select to load it explicitly!

## 💻 Usage Instructions

With the tool formally loaded onto your layout, kick off the core routine using:

1. **Start the Command**: Track your UI cursor physically rendering layouts using:
   ```text
   CalcGridSlopeCSV3
   ```
2. **Parameters**: Define the exact dimension for `sideLength (L)` corresponding directly identically to the side measurements of your Polylines. (e.g., `25`).
3. **Contour Intersect Mapping**: Feed the mathematical boundaries `interval (\u0394h)` parameters.
4. **Target Selection Pipeline**: Use the active UI interactions to physically click:
   - A single Sample Grid Boundary natively mapping identical layer extraction.
   - A single Sample Contour Line mapping native Layer sets organically.
   - The primary Project Layout Boundary mapping.
5. **Geometry Instantiation**: Define your structural Master Point placing the target Output Legend Anchor!

## 📊 Outputs Generated

- **Grid Visual Markers**: Generates explicit arrow limits mapping aspects accurately bounded recursively into each Grid's relative UCS bounds tracking true North calculations flawlessly.
- **Analysis Table Tracker**: Detailed Grid summaries.
- **Reference Table**: Complete formula constraint bounds physically mapping values exactly against bounds from `n=0` scaling natively structurally to `n=100`.
- **Rotatable Geometric Compass Legend**: Dynamically drafts an identical 3x3 geometric compass legend array structurally mapping visual identical rotations parallel to your viewport's specific tracking limits dynamically tracking your vectors explicitly!

---

# AutoCAD 網格坡度與坡向計算工具 (V3)

這是一個強大的原生 C# AutoCAD .NET 自動化工具，專為快速評估工程等高線坡度、計算真實正北坡向方位角而設計，並能自動產生綜合的 AutoCAD 原生幾何視覺化圖形、統計摘要表以及指南針圖例。

## 🌟 主要功能

- **面積加權地形分析**：提取網格與專案邊界重疊的布林運算 (Boolean) 面積，以準確計算非規則網格幾何區域內真實的平均坡度。
- **正北方位角計算**：利用最小平方法建構等高線三維交點向量，計算並產生地理真實正北 (Absolute WCS) 基準的下坡方位向量。
- **動態幾何視覺化**：自動產生方向圖例與動態內部箭頭標示，箭頭圖形邊界完美貼齊您目前的使用者座標系統 (UCS)，同時確保內部箭頭的指向邏輯嚴格對應真實的絕對重力流向。
- **原生表格產生**：自動產生動態總結矩陣與統計報表，並建立 AutoCAD 原生表格，詳細列出每個網格的 ID、坡度、坡向與法規級別。
- **匯出為 CSV**：將生成的統計數據矩陣直接寫入 CSV 檔案，採用純 UTF-8 BOM 編碼串流，確保中文字元及 `m²` 等符號能在 Microsoft Excel 環境中完美無瑕地顯示。

## 🛠 執行環境需求

此工具**不**依賴複雜的 Civil 3D 專有圖元功能！本程式完全基於底層的 AutoCAD 核心資料庫 (DB) 與純幾何運算架構運行！
- 任何支援 .NET 架構的現代版本 **Autodesk AutoCAD**。
- 符合您 AutoCAD 版本的基礎 `.NET` 執行階段框架 (通常取決於 AutoCAD 版本限制，建議使用 `.NET Framework 4.7+`)。

## 🚀 安裝與載入

1. **編譯 DLL 檔**：若您欲自行從原始碼編譯 DLL，可執行資料夾內提供的 `buildV3.bat` 指令碼（請確保腳本內的 `acmgd.dll`、`acdbmgd.dll` 參考路徑正確對應您的 AutoCAD 安裝目錄）。
2. **啟動 AutoCAD**：開啟您相關的專案地形 `DWG` 圖檔。
3. **載入程式 (NETLOAD)**：在 AutoCAD 底部指令列中直接輸入：
   ```text
   NETLOAD
   ```
4. 在彈出的瀏覽視窗中，找到剛編譯完成的 `./GridSlopeCalculatorV3.dll` 檔案並選取載入！

## 💻 使用說明

當工具成功載入至您的圖面後，請依照下列步驟執行核心計算運作：

1. **啟動指令**：在指令列中輸入以下指令以追蹤 UI 提示：
   ```text
   CalcGridSlopeCSV3
   ```
2. **網格參數**：輸入與您繪製的方格對應的「網格邊長 (L)」（例如輸入 `25`）。
3. **等高線距離**：輸入您地形圖的「等高線間距 (\u0394h)」高差參數。
4. **目標選取流程**：依照滑鼠十字游標旁的介面提示點選目標：
   - 任意選取一個樣本「網格 (Grid)」，系統會自動萃取並全選該圖層。
   - 任意選取一條樣本「等高線 (Contour)」，系統同樣會自動鎖定該圖層所有線段。
   - 選取您的「專案邊界 (Project Boundary)」（請確保是單一封閉的多段線）。
5. **幾何輸出定位**：在畫面上點選一個插入點，做為「摘要表格」及「指南針圖例」的錨點位置！

## 📊 輸出報表與圖元

- **網格視覺標示**：在每個網格內產生精確的下坡箭頭、坡度結果、重疊面積與坡向標籤，並準確貼合目前 UCS。
- **數據分析表**：詳細的坡度計算成果報表。
- **法規級別參考表**：按照指定的邊長大小與等高線間隔，自動建構從 `n=0` 延伸至 `n=100` 的理論參考矩陣表格。
- **可旋轉的幾何指南針圖例**：動態繪製 $3 \times 3$ 陣列的幾何指南針圖例，外框動態平行於您的 UCS 視埠旋轉角度，而內部向量明確指向您的真實北向！

---

**本腳本嚴格利用 AutoCAD 原生 `.NET` 幾何框架開發，無縫接軌真實世界地形座標限制！**
