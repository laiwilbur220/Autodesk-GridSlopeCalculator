# Context
我正在開發一個 AutoCAD .NET 外掛 (C#)，原始碼檔案為 `GridMethodSlopeCalculator.cs`。目前的 `GM1_Calc` 指令在處理數千個方格與等高線的交點時非常緩慢，因為它對每一個方格的 4 個邊與每一條等高線都執行了 `IntersectWith`。

# Task
請幫我優化計算交點的雙重迴圈效能，並加入 UI 回饋機制。

# Constraints & Implementation Details
1. 在執行 `segment.IntersectWith(contourCurve, ...)` 之前，請先取得 `segment` 的 Bounding Box (包絡盒) 與 `contourCurve` 的包絡盒 (`Extents3d`)。
2. 撰寫一個輕量級的交集判斷，只有當這兩個包絡盒在 X 與 Y 軸上有重疊時，才執行後續的 `IntersectWith` 運算，否則直接 continue 略過。
3. 在主要的網格處理迴圈 (迴圈遍歷 `preValidGrids` 時) 中，實作 `Application.StatusBar.MeterProgress` 來顯示運算進度 (0% 到 100%)。
4. 在進度條更新時，適時加入 `System.Windows.Forms.Application.DoEvents()`，以防止 AutoCAD 視窗在長時間運算時顯示「沒有回應」。

# Context
專案中的 `CalculateOverlapArea` 與 `CreateFlatRegionFromCurve` 方法，使用了 AutoCAD 原生的 `Region.BooleanOperation` 來計算方格與計畫邊界的多邊形交集面積。但在複雜或微小容差的地形邊界下，Region 轉換經常拋出 Exception 或失敗。

# Task
請將重疊面積計算邏輯重構，替換為穩定的 2D 幾何運算。假設我已經透過 NuGet 安裝了 `Clipper2` 函式庫。

# Constraints & Implementation Details
1. 請移除 `CreateFlatRegionFromCurve` 中依賴 `Region.CreateFromCurves` 的相關程式碼。
2. 重寫 `CalculateOverlapArea(Polyline pline, Curve boundaryCurve, Editor ed)` 方法。
3. 實作一個輔助方法，將 AutoCAD 的 Polyline/Polyline2d/Polyline3d 的 2D 頂點座標 (X, Y) 轉換為 Clipper2 支援的 `Path64` 或 `PathD` 格式。
4. 使用 Clipper2 執行兩個多邊形的交集運算 (Intersection)。
5. 取得交集後的多邊形，計算其精確面積並回傳 (double)。請確保運算過程不會因為共線、自我交集或微小碎面而拋出錯誤。