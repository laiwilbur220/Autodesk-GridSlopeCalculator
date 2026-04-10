using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Civil3DGridMethod
{
    /// <summary>
    /// V4 Grid Slope Calculator — Automated terrain slope analysis for AutoCAD Civil 3D.
    /// Features:
    ///   • Optional automatic grid generation aligned to UCS (25m or 10m cells)
    ///   • Slope calculation via Horton grid method: S = nπΔh / 8L × 100
    ///   • Downhill aspect via least-squares plane fitting
    ///   • Detailed outline direction arrows with center-point positioning
    ///   • Area-weighted summary statistics and compass legend
    ///   • Optional CSV export (UTF-8 BOM for Excel compatibility)
    /// </summary>
    public class GridSlopeCalculatorV4
    {
        #region Constants

        /// <summary>Tolerance for deduplicating intersection points (meters).</summary>
        private const double IntersectTolerance = 0.02;

        /// <summary>Fractional tolerance for validating grid area against expected square area.</summary>
        private const double AreaValidationTolerance = 0.05;

        /// <summary>Minimum overlap area (m²) to consider a grid inside the boundary.</summary>
        private const double MinOverlapArea = 0.001;

        /// <summary>Minimum total weight for weighted average (prevents division by near-zero).</summary>
        private const double MinWeightThreshold = 0.0001;

        /// <summary>Determinant threshold for least-squares singularity check.</summary>
        private const double SingularityThreshold = 1e-8;

        /// <summary>Text height as a fraction of grid side length.</summary>
        private const double TextHeightRatio = 0.15;

        /// <summary>Row sorting snap tolerance (meters).</summary>
        private const double SortingTolerance = 1.0;

        // Layout positioning ratios (fraction of sideLength from centroid)
        private const double IdOffsetRatio = 0.45;              // grid ID, top-left corner
        private const double AreaOffsetYRatio = 0.44;           // direction text, top-center
        private const double ArrowUpwardShift = 0.15;           // arrow center shifted up
        private const double SlopeTextYRatio = 0.12;            // slope % text, upper zone
        private const double ClassTextYRatio = 0.09;            // classification, below slope
        private const double IntersectsOffsetYRatio = 0.15;     // n= count, mid-lower zone
        private const double DirectionTextOffsetYRatio = 0.40;  // overlap area text, bottom

        // Arrow geometry (legacy ratios, kept for reference; V4 uses outline template)
        private const double ArrowLengthRatio = 0.4;
        private const double ArrowShaftWidthRatio = 0.015;
        private const double ArrowHeadWidthRatio = 0.1;

        // Table defaults
        private const double TableRowHeight = 4.0;
        private const double TableTextHeight = 1.8;
        private const int MaxSummaryRowsPerChunk = 20;
        private const int MaxLegendRowsPerChunk = 26;

        #endregion

        #region Data Structures

        /// <summary>
        /// Holds per-cell results: centroid, slope, classification, direction, and overlap area.
        /// </summary>
        public class GridData
        {
            /// <summary>ObjectId of the grid polyline in the drawing.</summary>
            public ObjectId Id { get; set; }

            /// <summary>Cell centroid in World Coordinate System.</summary>
            public Point3d CentroidWCS { get; set; }

            /// <summary>Cell centroid in User Coordinate System.</summary>
            public Point3d CentroidUCS { get; set; }

            /// <summary>UCS X-coordinate for left-to-right sorting.</summary>
            public double SortX { get; set; }

            /// <summary>UCS Y-coordinate for top-to-bottom sorting.</summary>
            public double SortY { get; set; }

            /// <summary>Total unique contour intersection count on this cell's edges.</summary>
            public int TotalIntersects { get; set; }

            /// <summary>Computed slope percentage via Horton grid method.</summary>
            public double SlopePercent { get; set; }

            /// <summary>Slope classification category (一級坡 through 七級坡).</summary>
            public string Classification { get; set; }

            /// <summary>Compass direction label for the downhill aspect (e.g. "NE", "S").</summary>
            public string Direction { get; set; }

            public GridData()
            {
                Classification = "";
                Direction = "";
            }

            /// <summary>Area of intersection between grid cell and project boundary (m²).</summary>
            public double LappingArea { get; set; }
        }

        /// <summary>
        /// Computed weighted-average summary across all valid grids.
        /// </summary>
        private struct WeightedSummary
        {
            public double MeanSlope;
            public string ModeDirection;
        }

        #endregion

        #region Main AutoCAD Command

        /// <summary>
        /// Main entry point — AutoCAD command: CalcGridSlopeCSV4
        /// Workflow: grid existence check → (optional grid generation) → slope calculation → output.
        /// </summary>
        [CommandMethod("CalcGridSlopeCSV4")]
        public void CalculateGridSlope()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // ===== STEP 0: Check if grids need to be built =====
                PromptKeywordOptions pkoGrid = new PromptKeywordOptions(
                    "\nHas grid already been built? [Yes/No] ", "Yes No");
                pkoGrid.Keywords.Default = "Yes";
                pkoGrid.AllowNone = true;
                PromptResult prGrid = ed.GetKeywords(pkoGrid);
                if (prGrid.Status != PromptStatus.OK && prGrid.Status != PromptStatus.None) return;
                bool gridAlreadyBuilt = (prGrid.StringResult == "Yes");

                double sideLength = 25.0;
                ObjectId preselectedBoundaryId = ObjectId.Null;
                string generatedGridLayer = "GRID";

                if (!gridAlreadyBuilt)
                {
                    // ===== GRID GENERATION PHASE =====

                    // 1. Select project boundary
                    PromptEntityOptions peoBound = new PromptEntityOptions(
                        "\nSelect the Project Boundary (Polyline): ");
                    peoBound.SetRejectMessage("\nBoundary must be a Polyline!");
                    peoBound.AddAllowedClass(typeof(Polyline), false);
                    peoBound.AddAllowedClass(typeof(Polyline2d), false);
                    peoBound.AddAllowedClass(typeof(Polyline3d), false);
                    PromptEntityResult perBound = ed.GetEntity(peoBound);
                    if (perBound.Status != PromptStatus.OK) return;
                    preselectedBoundaryId = perBound.ObjectId;

                    // Verify boundary selection
                    using (Transaction trVerify = db.TransactionManager.StartTransaction())
                    {
                        ObjectId[] verifyBoundIds = new ObjectId[] { preselectedBoundaryId };
                        HighlightObjects(trVerify, verifyBoundIds, true);
                        PromptKeywordOptions pkoVerify = new PromptKeywordOptions(
                            "\nConfirm selected project boundary? [Yes/No] ", "Yes No");
                        pkoVerify.Keywords.Default = "Yes";
                        PromptResult prVerify = ed.GetKeywords(pkoVerify);
                        HighlightObjects(trVerify, verifyBoundIds, false);
                        trVerify.Commit();
                        if (prVerify.Status != PromptStatus.OK || prVerify.StringResult != "Yes") return;
                    }

                    // 2. Ask grid size
                    PromptKeywordOptions pkoSize = new PromptKeywordOptions(
                        "\nSelect grid cell size [25m/10m] ", "25m 10m");
                    pkoSize.Keywords.Default = "25m";
                    pkoSize.AllowNone = true;
                    PromptResult prSize = ed.GetKeywords(pkoSize);
                    if (prSize.Status != PromptStatus.OK && prSize.Status != PromptStatus.None) return;
                    sideLength = (prSize.StringResult == "10m") ? 10.0 : 25.0;

                    // 3. Generate grids in a separate transaction (committed before slope calc)
                    Matrix3d ucsForGrid = ed.CurrentUserCoordinateSystem;
                    using (Transaction trGrid = db.TransactionManager.StartTransaction())
                    {
                        BlockTable btGrid = (BlockTable)trGrid.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btrGrid = (BlockTableRecord)trGrid.GetObject(
                            btGrid[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        Curve boundaryCurve = trGrid.GetObject(preselectedBoundaryId, OpenMode.ForRead) as Curve;
                        if (boundaryCurve == null)
                        {
                            ed.WriteMessage("\nFailed to load boundary polyline.");
                            return;
                        }

                        int gridCount = BuildGridPattern(trGrid, btrGrid, db, boundaryCurve,
                            sideLength, ucsForGrid, ed, generatedGridLayer);

                        if (gridCount == 0)
                        {
                            ed.WriteMessage("\nNo valid grids generated within boundary.");
                            return;
                        }

                        trGrid.Commit();
                        ed.WriteMessage(string.Format(
                            "\nGenerated {0} grid cells ({1}m) on layer [{2}].",
                            gridCount, sideLength, generatedGridLayer));
                    }
                }
                else
                {
                    // Existing grids — ask side length manually
                    sideLength = PromptForDouble(ed, "\nEnter grid side length (L) in meters: ", 25.0);
                    if (sideLength <= 0) return;
                }

                // ===== SLOPE CALCULATION PHASE =====
                double contourInterval = PromptForDouble(ed, "\nEnter contour interval (Δh) in meters: ", 1.0);
                if (contourInterval <= 0) return;

                PromptKeywordOptions pkoCsv = new PromptKeywordOptions(
                    "\nExport data to CSV? [Yes/No] ", "Yes No");
                pkoCsv.AllowNone = true;
                pkoCsv.Keywords.Default = "No";
                PromptResult prExport = ed.GetKeywords(pkoCsv);
                if (prExport.Status != PromptStatus.OK && prExport.Status != PromptStatus.None) return;
                bool shouldExportCsv = (prExport.StringResult == "Yes");

                // ------------- STEP 2: CAD SELECTION PROCESS -------------
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 2a. Grid selection
                    PromptSelectionResult psrAllGrids;
                    if (!gridAlreadyBuilt)
                    {
                        // Auto-select grids from the generated layer
                        SelectionFilter gridFilter = new SelectionFilter(new TypedValue[] {
                            new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                            new TypedValue((int)DxfCode.LayerName, generatedGridLayer)
                        });
                        psrAllGrids = ed.SelectAll(gridFilter);
                        if (psrAllGrids.Status != PromptStatus.OK)
                        {
                            ed.WriteMessage("\nFailed to auto-select generated grids.");
                            return;
                        }
                        ed.WriteMessage(string.Format("\nAuto-selected {0} grids from [{1}] layer.",
                            psrAllGrids.Value.Count, generatedGridLayer));
                    }
                    else
                    {
                        psrAllGrids = GetSelectionByLayer(ed, tr, "GRID(s)", "LWPOLYLINE,POLYLINE");
                        if (psrAllGrids == null) return;
                    }

                    // 2b. Select contours
                    PromptSelectionResult psrAllContours = GetSelectionByLayer(ed, tr, "CONTOUR(s)",
                        "LWPOLYLINE,POLYLINE,3DPOLYLINE");
                    if (psrAllContours == null) return;

                    // 2c. Boundary selection / confirmation
                    ObjectId boundaryId;
                    if (!gridAlreadyBuilt && preselectedBoundaryId != ObjectId.Null)
                    {
                        // Re-use boundary from grid generation
                        boundaryId = preselectedBoundaryId;
                        ObjectId[] boundIds = new ObjectId[] { boundaryId };
                        HighlightObjects(tr, boundIds, true);
                        PromptKeywordOptions pkoBound = new PromptKeywordOptions(
                            "\nUse same boundary from grid generation? [Yes/No] ", "Yes No");
                        pkoBound.Keywords.Default = "Yes";
                        PromptResult prBound = ed.GetKeywords(pkoBound);
                        HighlightObjects(tr, boundIds, false);
                        if (prBound.Status != PromptStatus.OK || prBound.StringResult != "Yes") return;
                    }
                    else
                    {
                        PromptEntityOptions peo = new PromptEntityOptions(
                            "\nSelect the Project Boundary (Polyline): ");
                        peo.SetRejectMessage("\nBoundary must be a Polyline!");
                        peo.AddAllowedClass(typeof(Polyline), false);
                        peo.AddAllowedClass(typeof(Polyline2d), false);
                        peo.AddAllowedClass(typeof(Polyline3d), false);
                        PromptEntityResult perBoundary = ed.GetEntity(peo);
                        if (perBoundary.Status != PromptStatus.OK) return;
                        boundaryId = perBoundary.ObjectId;

                        ObjectId[] boundIds = new ObjectId[] { boundaryId };
                        HighlightObjects(tr, boundIds, true);
                        PromptKeywordOptions pkoBound = new PromptKeywordOptions(
                            "\nConfirm selected project boundary? [Yes/No] ", "Yes No");
                        pkoBound.Keywords.Default = "Yes";
                        PromptResult prBound = ed.GetKeywords(pkoBound);
                        HighlightObjects(tr, boundIds, false);
                        if (prBound.Status != PromptStatus.OK || prBound.StringResult != "Yes") return;
                    }

                    // 2d. Coordinate system setup
                    Matrix3d ucs = ed.CurrentUserCoordinateSystem;
                    CoordinateSystem3d cs = ucs.CoordinateSystem3d;

                    PromptPointOptions ppo = new PromptPointOptions(
                        "\nSelect insertion point for the Summary Table: ");
                    PromptPointResult ppr = ed.GetPoint(ppo);
                    if (ppr.Status != PromptStatus.OK) return;
                    Point3d tableInsertPtWCS = ppr.Value.TransformBy(ucs);

                    // ------------- STEP 3: PREPARE LAYERS & GEOMETRY -------------
                    ObjectId layerId = GetOrCreateLayer(db, tr, "Grid_Outputs_ID", 1);
                    ObjectId layerEdges = GetOrCreateLayer(db, tr, "Grid_Outputs_Edges", 2);
                    ObjectId layerTotal = GetOrCreateLayer(db, tr, "Grid_Outputs_Total", 3);
                    ObjectId layerSlope = GetOrCreateLayer(db, tr, "Grid_Outputs_Slope", 4);
                    ObjectId layerDirection = GetOrCreateLayer(db, tr, "Grid_Outputs_Direction", 4);
                    ObjectId layerDirText = GetOrCreateLayer(db, tr, "Grid_Outputs_DirText", 6);
                    ObjectId layerArea = GetOrCreateLayer(db, tr, "Grid_Outputs_Area", 160);
                    ObjectId layerTable = GetOrCreateLayer(db, tr, "Grid_Outputs_Table", 7);

                    List<Curve> allContourCurves = GetCurvesFromSelection(tr, psrAllContours);
                    if (allContourCurves.Count == 0)
                    {
                        ed.WriteMessage("\nNo valid contour curves found.");
                        return;
                    }

                    Curve baseBoundaryCurve = tr.GetObject(boundaryId, OpenMode.ForRead) as Curve;
                    if (baseBoundaryCurve == null)
                    {
                        ed.WriteMessage("\nFailed to load boundary polyline.");
                        return;
                    }

                    Region boundaryRegion = CreateFlatRegionFromCurve(baseBoundaryCurve, ed);
                    if (boundaryRegion == null)
                    {
                        ed.WriteMessage("\nERROR: Failed to create planar boundary region.");
                        return;
                    }

                    try
                    {
                        // ------------- STEP 4: PROCESS GRID CELLS -------------
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        Plane flatPlane = new Plane(Point3d.Origin, Vector3d.ZAxis);
                        HashSet<string> processedEdges = new HashSet<string>();
                        double textHeight = sideLength * TextHeightRatio;

                        List<GridData> preValidGrids = GetValidGrids(tr, psrAllGrids, sideLength, ucs, ed);
                        if (preValidGrids.Count == 0) return;

                        preValidGrids = preValidGrids
                            .OrderByDescending(g => Math.Round(g.SortY / SortingTolerance) * SortingTolerance)
                            .ThenBy(g => g.SortX)
                            .ToList();

                        List<GridData> validGrids = new List<GridData>();

                        for (int i = 0; i < preValidGrids.Count; i++)
                        {
                            GridData currentGrid = preValidGrids[i];
                            Polyline pline = tr.GetObject(currentGrid.Id, OpenMode.ForRead) as Polyline;

                            if (pline == null || pline.NumberOfVertices < 4)
                            {
                                ed.WriteMessage(string.Format("\nSkipping grid {0}: need at least 4 vertices, found {1}.",
                                    i + 1, pline == null ? 0 : pline.NumberOfVertices));
                                continue;
                            }

                            currentGrid.LappingArea = CalculateOverlapArea(pline, boundaryRegion, ed);
                            if (currentGrid.LappingArea <= MinOverlapArea) continue;

                            List<Point3d> allGridIntersects = new List<Point3d>();

                            for (int j = 0; j < 4; j++)
                            {
                                using (Line segment = new Line(pline.GetPoint3dAt(j), pline.GetPoint3dAt((j + 1) % 4)))
                                {
                                    List<Point3d> uniqueSegmentPts = new List<Point3d>();

                                    foreach (Curve contourCurve in allContourCurves)
                                    {
                                        using (Point3dCollection rawPts = new Point3dCollection())
                                        {
                                            segment.IntersectWith(contourCurve, Intersect.OnBothOperands,
                                                flatPlane, rawPts, IntPtr.Zero, IntPtr.Zero);
                                            foreach (Point3d pt in rawPts)
                                            {
                                                if (!ContainsPointWithTolerance(uniqueSegmentPts, pt, IntersectTolerance))
                                                {
                                                    Point3d ptOnCurve = contourCurve.GetClosestPointTo(pt, false);
                                                    Point3d ptWithZ = new Point3d(pt.X, pt.Y, ptOnCurve.Z);
                                                    uniqueSegmentPts.Add(ptWithZ);
                                                }
                                            }
                                        }
                                    }
                                    allGridIntersects.AddRange(uniqueSegmentPts);

                                    int segmentIntersects = uniqueSegmentPts.Count;
                                    Point3d midPt = new Point3d(
                                        (segment.StartPoint.X + segment.EndPoint.X) / 2,
                                        (segment.StartPoint.Y + segment.EndPoint.Y) / 2, 0);

                                    string edgeKey = string.Format("{0}_{1}",
                                        Math.Round(midPt.X, 3), Math.Round(midPt.Y, 3));

                                    if (processedEdges.Add(edgeKey))
                                    {
                                        AddTextToBTR(tr, btr, segmentIntersects.ToString(), midPt,
                                            textHeight * 0.8, layerEdges, cs, AttachmentPoint.MiddleCenter);
                                    }
                                }
                            }

                            List<Point3d> finalGridPts = new List<Point3d>();
                            foreach (Point3d pt in allGridIntersects)
                            {
                                if (!ContainsPointWithTolerance(finalGridPts, pt, IntersectTolerance))
                                    finalGridPts.Add(pt);
                            }

                            Vector3d slopeDir;
                            currentGrid.Direction = CalculateAspectAndDirection(finalGridPts, out slopeDir);
                            currentGrid.TotalIntersects = finalGridPts.Count;
                            currentGrid.SlopePercent = ((currentGrid.TotalIntersects * Math.PI * contourInterval)
                                / (8 * sideLength)) * 100;
                            currentGrid.Classification = GetSlopeClassification(currentGrid.SlopePercent);

                            Point3d cUCS = currentGrid.CentroidUCS;
                            Point3d cWCS = currentGrid.CentroidWCS;

                            // Draw direction arrow shifted upward
                            Point3d arrowCenter = new Point3d(cWCS.X, cWCS.Y, 0)
                                + new Vector3d(cs.Yaxis.X, cs.Yaxis.Y, 0) * (sideLength * ArrowUpwardShift);
                            if (!string.IsNullOrEmpty(currentGrid.Direction))
                            {
                                DrawDirectionArrow(tr, btr, arrowCenter, slopeDir, sideLength, layerDirection);
                            }

                            // --- Per-cell annotation labels (all centered horizontally) ---
                            int displayId = validGrids.Count + 1;

                            Point3d idPt = new Point3d(cUCS.X - (sideLength * IdOffsetRatio),
                                cUCS.Y + (sideLength * IdOffsetRatio), 0).TransformBy(ucs);
                            AddTextToBTR(tr, btr, displayId.ToString(), idPt, textHeight,
                                layerId, cs, AttachmentPoint.TopLeft);

                            // Direction text — top-center (swapped with area)
                            if (!string.IsNullOrEmpty(currentGrid.Direction))
                            {
                                Point3d dirTextPt = new Point3d(cUCS.X,
                                    cUCS.Y + (sideLength * AreaOffsetYRatio), 0).TransformBy(ucs);
                                AddTextToBTR(tr, btr, currentGrid.Direction, dirTextPt,
                                    textHeight * 0.8, layerDirText, cs, AttachmentPoint.TopCenter);
                            }

                            Point3d slopePt = new Point3d(cUCS.X,
                                cUCS.Y + (sideLength * SlopeTextYRatio), 0).TransformBy(ucs);
                            AddTextToBTR(tr, btr, string.Format("{0:F2}%", currentGrid.SlopePercent),
                                slopePt, textHeight, layerSlope, cs, AttachmentPoint.BottomCenter);

                            Point3d classPt = new Point3d(cUCS.X,
                                cUCS.Y + (sideLength * ClassTextYRatio), 0).TransformBy(ucs);
                            string classText = string.Format("({0})", currentGrid.Classification);
                            AddTextToBTR(tr, btr, classText, classPt, textHeight * 0.8,
                                layerSlope, cs, AttachmentPoint.TopCenter);

                            Point3d nPt = new Point3d(cUCS.X,
                                cUCS.Y - (sideLength * IntersectsOffsetYRatio), 0).TransformBy(ucs);
                            AddTextToBTR(tr, btr, string.Format("n={0}", currentGrid.TotalIntersects),
                                nPt, textHeight * 0.8, layerTotal, cs, AttachmentPoint.MiddleCenter);

                            // Overlap area — bottom of cell (swapped with direction)
                            {
                                Point3d areaPt = new Point3d(cUCS.X,
                                    cUCS.Y - (sideLength * DirectionTextOffsetYRatio), 0).TransformBy(ucs);
                                string areaText = string.Format("A={0:F0} m\u00B2", currentGrid.LappingArea);
                                AddTextToBTR(tr, btr, areaText, areaPt, textHeight * 0.8,
                                    layerArea, cs, AttachmentPoint.BottomCenter);
                            }

                            validGrids.Add(currentGrid);
                        }

                        // ------------- STEP 5: OUTPUT GENERATION -------------
                        GenerateSummaryTable(tr, btr, db, cs, validGrids, tableInsertPtWCS, layerTable);

                        int mainChunks = (validGrids.Count == 0) ? 1
                            : ((validGrids.Count - 1) / MaxSummaryRowsPerChunk) + 1;
                        double summaryTableWidth = mainChunks * 88.0;

                        Point3d legendInsertPt = tableInsertPtWCS + (cs.Xaxis * (summaryTableWidth + 40.0));
                        GenerateLegendTable(tr, btr, db, cs, legendInsertPt, layerTable,
                            contourInterval, sideLength);

                        double legendChunkWidth = 16.0 + 15.0 + 18.0;
                        int legendChunks = ((51 - 1) / MaxLegendRowsPerChunk) + 1;
                        double legendTotalWidth = legendChunks * legendChunkWidth;
                        Point3d compassCenterPt = legendInsertPt
                            + (cs.Xaxis * (legendTotalWidth + 20.0 + (sideLength * 1.5)))
                            - (cs.Yaxis * (sideLength * 2.0));

                        DrawCompassLegend(tr, btr, cs, ucs, compassCenterPt, sideLength,
                            layerEdges, layerDirection, layerDirText);

                        if (shouldExportCsv) ExportToCSV(doc, validGrids);

                        tr.Commit();
                        ed.WriteMessage(string.Format("\nSuccess! Processed {0} grids.", validGrids.Count));
                    }
                    finally
                    {
                        if (boundaryRegion != null)
                            boundaryRegion.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(string.Format("\nError: {0}\n{1}", ex.Message, ex.StackTrace));
            }
        }

        #endregion

        #region User Interactive Core (CLI)

        /// <summary>
        /// Prompts for a positive double value with an optional default.
        /// Returns -1.0 if cancelled.
        /// </summary>
        private double PromptForDouble(Editor ed, string promptMessage, double defaultValue = 0.0)
        {
            PromptDoubleOptions pdo = new PromptDoubleOptions(promptMessage);
            pdo.AllowNegative = false;
            pdo.AllowZero = false;
            if (defaultValue > 0.0)
            {
                pdo.DefaultValue = defaultValue;
                pdo.UseDefaultValue = true;
            }
            PromptDoubleResult pdr = ed.GetDouble(pdo);
            return pdr.Status == PromptStatus.OK ? pdr.Value : -1.0;
        }

        /// <summary>
        /// Lets the user pick sample entities, auto-selects all entities on those layers,
        /// highlights them for confirmation, and returns the final selection.
        /// </summary>
        private PromptSelectionResult GetSelectionByLayer(Editor ed, Transaction tr, string itemType, string dxfFilter)
        {
            while (true)
            {
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = string.Format("\nSelect sample {0} to auto-select their layers: ", itemType);
                PromptSelectionResult psrSample = ed.GetSelection(pso);
                if (psrSample.Status != PromptStatus.OK) return null;

                string layers = GetUniqueLayersFromSelection(tr, psrSample.Value);
                SelectionFilter filter = new SelectionFilter(new TypedValue[] {
                    new TypedValue((int)DxfCode.Start, dxfFilter),
                    new TypedValue((int)DxfCode.LayerName, layers)
                });

                PromptSelectionResult psrAll = ed.SelectAll(filter);
                if (psrAll.Status != PromptStatus.OK) return null;

                ObjectId[] ids = psrAll.Value.GetObjectIds();
                HighlightObjects(tr, ids, true);

                PromptKeywordOptions pko = new PromptKeywordOptions(
                    string.Format("\nFound {0} items on layer(s) [{1}]. Is this correct? [Yes/No] ", ids.Length, layers), "Yes No");
                pko.Keywords.Default = "Yes";

                PromptResult pr = ed.GetKeywords(pko);
                HighlightObjects(tr, ids, false);

                if (pr.Status != PromptStatus.OK) return null;
                if (pr.StringResult == "Yes") return psrAll;
            }
        }

        /// <summary>
        /// Extracts Curve objects from a selection result.
        /// </summary>
        private List<Curve> GetCurvesFromSelection(Transaction tr, PromptSelectionResult psr)
        {
            List<Curve> curves = new List<Curve>();
            foreach (SelectedObject so in psr.Value)
            {
                Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                if (curve != null) curves.Add(curve);
            }
            return curves;
        }

        #endregion

        #region Core Mathematical Engine

        /// <summary>
        /// Validates grid polylines: must be closed, approximately square with expected area.
        /// Computes centroids in both WCS and UCS.
        /// </summary>
        private List<GridData> GetValidGrids(Transaction tr, PromptSelectionResult psr,
            double sideLength, Matrix3d ucs, Editor ed)
        {
            List<GridData> valid = new List<GridData>();
            double targetArea = sideLength * sideLength;

            foreach (SelectedObject so in psr.Value)
            {
                Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                if (curve != null && curve.Closed)
                {
                    if (Math.Abs(curve.Area - targetArea) < (targetArea * AreaValidationTolerance))
                    {
                        GridData gd = new GridData();
                        gd.Id = so.ObjectId;

                        Point3d p1 = curve.GetPointAtParameter(0);
                        Point3d p3 = curve.GetPointAtParameter(2);
                        gd.CentroidWCS = new Point3d((p1.X + p3.X) / 2, (p1.Y + p3.Y) / 2, 0);

                        gd.CentroidUCS = gd.CentroidWCS.TransformBy(ucs.Inverse());
                        gd.SortX = gd.CentroidUCS.X;
                        gd.SortY = gd.CentroidUCS.Y;

                        valid.Add(gd);
                    }
                }
            }
            return valid;
        }

        /// <summary>
        /// Generates a grid pattern aligned to the UCS that covers the project boundary.
        /// Grids that do not intersect/touch the boundary are erased.
        /// Returns the count of surviving grids.
        /// </summary>
        private int BuildGridPattern(Transaction tr, BlockTableRecord btr, Database db,
            Curve boundaryCurve, double sideLength, Matrix3d ucs, Editor ed, string gridLayerName)
        {
            // Get boundary WCS extents
            Extents3d? extents = boundaryCurve.Bounds;
            if (!extents.HasValue)
            {
                ed.WriteMessage("\nCannot determine boundary extents.");
                return 0;
            }

            // Transform all 4 WCS AABB corners to UCS to find correct UCS bounding box
            // (needed when UCS is rotated relative to WCS)
            Point3d minWCS = extents.Value.MinPoint;
            Point3d maxWCS = extents.Value.MaxPoint;
            Matrix3d ucsInv = ucs.Inverse();

            Point3d[] wcsCorners = new Point3d[] {
                new Point3d(minWCS.X, minWCS.Y, 0),
                new Point3d(maxWCS.X, minWCS.Y, 0),
                new Point3d(maxWCS.X, maxWCS.Y, 0),
                new Point3d(minWCS.X, maxWCS.Y, 0)
            };

            double ucsMinX = double.MaxValue, ucsMinY = double.MaxValue;
            double ucsMaxX = double.MinValue, ucsMaxY = double.MinValue;
            foreach (Point3d corner in wcsCorners)
            {
                Point3d ucsCorner = corner.TransformBy(ucsInv);
                if (ucsCorner.X < ucsMinX) ucsMinX = ucsCorner.X;
                if (ucsCorner.Y < ucsMinY) ucsMinY = ucsCorner.Y;
                if (ucsCorner.X > ucsMaxX) ucsMaxX = ucsCorner.X;
                if (ucsCorner.Y > ucsMaxY) ucsMaxY = ucsCorner.Y;
            }

            // Snap grid start to multiples of sideLength with one-cell buffer
            double startX = Math.Floor(ucsMinX / sideLength) * sideLength - sideLength;
            double startY = Math.Floor(ucsMinY / sideLength) * sideLength - sideLength;
            double endX = Math.Ceiling(ucsMaxX / sideLength) * sideLength + sideLength;
            double endY = Math.Ceiling(ucsMaxY / sideLength) * sideLength + sideLength;

            // Create boundary region for overlap testing
            Region boundaryRegion = CreateFlatRegionFromCurve(boundaryCurve, ed);
            if (boundaryRegion == null) return 0;

            ObjectId gridLayerId = GetOrCreateLayer(db, tr, gridLayerName, 180);
            int validCount = 0;

            try
            {
                int cols = (int)Math.Round((endX - startX) / sideLength);
                int rows = (int)Math.Round((endY - startY) / sideLength);
                ed.WriteMessage(string.Format("\nBuilding {0} x {1} grid pattern...", cols, rows));

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        double x = startX + col * sideLength;
                        double y = startY + row * sideLength;

                        // Grid corners in UCS → transformed to WCS
                        Point3d p1 = new Point3d(x, y, 0).TransformBy(ucs);
                        Point3d p2 = new Point3d(x + sideLength, y, 0).TransformBy(ucs);
                        Point3d p3 = new Point3d(x + sideLength, y + sideLength, 0).TransformBy(ucs);
                        Point3d p4 = new Point3d(x, y + sideLength, 0).TransformBy(ucs);

                        Polyline pl = new Polyline();
                        pl.AddVertexAt(0, new Point2d(p1.X, p1.Y), 0, 0, 0);
                        pl.AddVertexAt(1, new Point2d(p2.X, p2.Y), 0, 0, 0);
                        pl.AddVertexAt(2, new Point2d(p3.X, p3.Y), 0, 0, 0);
                        pl.AddVertexAt(3, new Point2d(p4.X, p4.Y), 0, 0, 0);
                        pl.Closed = true;
                        pl.LayerId = gridLayerId;

                        btr.AppendEntity(pl);
                        tr.AddNewlyCreatedDBObject(pl, true);

                        // Check overlap with boundary — erase if no intersection
                        double overlap = CalculateOverlapArea(pl, boundaryRegion, ed);
                        if (overlap <= MinOverlapArea)
                        {
                            pl.Erase();
                        }
                        else
                        {
                            validCount++;
                        }
                    }
                }
            }
            finally
            {
                boundaryRegion.Dispose();
            }

            return validCount;
        }

        /// <summary>
        /// Creates a flat (Z=0) Region from a curve, handling Polyline, Polyline2d, and Polyline3d.
        /// Returns null if the region cannot be created.
        /// </summary>
        private Region CreateFlatRegionFromCurve(Curve sourceCurve, Editor ed)
        {
            try
            {
                using (Curve clone = sourceCurve.Clone() as Curve)
                {
                    // Flatten the curve to Z=0 for planar region creation
                    if (clone is Polyline)
                    {
                        ((Polyline)clone).Closed = true;
                        ((Polyline)clone).Elevation = 0.0;
                    }
                    else if (clone is Polyline2d)
                    {
                        ((Polyline2d)clone).Closed = true;
                        ((Polyline2d)clone).Elevation = 0.0;
                    }
                    else if (clone is Polyline3d)
                    {
                        // §3.1 — Polyline3d has no Elevation property; vertices are 3D.
                        // Region.CreateFromCurves may fail if non-planar — catch handles with clear message.
                        ((Polyline3d)clone).Closed = true;
                    }

                    DBObjectCollection curveCollection = new DBObjectCollection();
                    curveCollection.Add(clone);
                    DBObjectCollection regionCollection = Region.CreateFromCurves(curveCollection);
                    if (regionCollection.Count > 0)
                    {
                        Region result = regionCollection[0] as Region;
                        // Dispose any extra regions (unlikely but defensive)
                        for (int i = 1; i < regionCollection.Count; i++)
                            regionCollection[i].Dispose();
                        return result;
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception aCADEx)
            {
                ed.WriteMessage(string.Format(
                    "\nERROR creating boundary Region: {0}. Ensure the boundary is planar and non-self-intersecting.",
                    aCADEx.Message));
            }
            return null;
        }

        /// <summary>
        /// Computes the overlap area between a grid polyline and the project boundary via Boolean intersection.
        /// Falls back to the raw polyline area with a warning if the operation fails.
        /// </summary>
        private double CalculateOverlapArea(Polyline pline, Region boundaryRegion, Editor ed)
        {
            // §2.3 — fixed: no double-dispose; gridRegion lifecycle is explicit
            Region gridRegion = null;
            try
            {
                DBObjectCollection gridCol = new DBObjectCollection();
                using (Curve gridClone = pline.Clone() as Curve)
                {
                    gridCol.Add(gridClone);
                    DBObjectCollection regionCol = Region.CreateFromCurves(gridCol);
                    if (regionCol.Count > 0)
                    {
                        gridRegion = regionCol[0] as Region;

                        // Dispose any extra region objects (not index 0, which we're using)
                        for (int i = 1; i < regionCol.Count; i++)
                            regionCol[i].Dispose();

                        using (Region boundClone = boundaryRegion.Clone() as Region)
                        {
                            gridRegion.BooleanOperation(BooleanOperationType.BoolIntersect, boundClone);
                            return gridRegion.Area;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                // §2.2 — log instead of silently swallowing
                ed.WriteMessage(string.Format("\nWARNING: Boolean overlap failed for grid, using raw area. [{0}]", ex.Message));
                return pline.Area;
            }
            finally
            {
                if (gridRegion != null)
                    gridRegion.Dispose();
            }
            return 0.0;
        }

        /// <summary>
        /// Checks if any existing point is within tolerance of the candidate point.
        /// </summary>
        private bool ContainsPointWithTolerance(List<Point3d> points, Point3d newPt, double tolerance)
        {
            foreach (Point3d uPt in points)
            {
                double dx = newPt.X - uPt.X;
                double dy = newPt.Y - uPt.Y;
                if (Math.Sqrt(dx * dx + dy * dy) < tolerance) return true;
            }
            return false;
        }

        /// <summary>
        /// Determines downhill direction via least-squares plane fit: z = a(x−x̄) + b(y−ȳ) + z̄.
        /// Returns the compass label (N/NE/E/SE/S/SW/W/NW) and the unit direction vector.
        /// Requires at least 3 points with elevation variation.
        /// </summary>
        private string CalculateAspectAndDirection(List<Point3d> pts, out Vector3d direction)
        {
            direction = Vector3d.XAxis;
            if (pts.Count < 3) return "";

            double sumX = 0, sumY = 0, sumZ = 0;
            foreach (var p in pts)
            {
                sumX += p.X; sumY += p.Y; sumZ += p.Z;
            }
            int n = pts.Count;
            double avgX = sumX / n;
            double avgY = sumY / n;
            double avgZ = sumZ / n;

            double sxx = 0, syy = 0, sxy = 0, sxz = 0, syz = 0;
            foreach (var p in pts)
            {
                double ux = p.X - avgX;
                double vy = p.Y - avgY;
                double wz = p.Z - avgZ;

                sxx += ux * ux;
                syy += vy * vy;
                sxy += ux * vy;
                sxz += ux * wz;
                syz += vy * wz;
            }

            double D = sxx * syy - sxy * sxy;
            if (Math.Abs(D) < SingularityThreshold) return "";

            double a = (sxz * syy - syz * sxy) / D;
            double b = (syz * sxx - sxz * sxy) / D;

            if (Math.Abs(a) < SingularityThreshold && Math.Abs(b) < SingularityThreshold) return "";

            Vector3d rawDir = new Vector3d(-a, -b, 0).GetNormal();
            double angle = Math.Atan2(rawDir.Y, rawDir.X);
            double deg = angle * 180.0 / Math.PI;

            if (deg < 0) deg += 360.0;

            if (deg >= 337.5 || deg < 22.5)  { direction = new Vector3d(1, 0, 0); return "E"; }
            if (deg >= 22.5  && deg < 67.5)  { direction = new Vector3d(1, 1, 0).GetNormal(); return "NE"; }
            if (deg >= 67.5  && deg < 112.5) { direction = new Vector3d(0, 1, 0); return "N"; }
            if (deg >= 112.5 && deg < 157.5) { direction = new Vector3d(-1, 1, 0).GetNormal(); return "NW"; }
            if (deg >= 157.5 && deg < 202.5) { direction = new Vector3d(-1, 0, 0); return "W"; }
            if (deg >= 202.5 && deg < 247.5) { direction = new Vector3d(-1, -1, 0).GetNormal(); return "SW"; }
            if (deg >= 247.5 && deg < 292.5) { direction = new Vector3d(0, -1, 0); return "S"; }
            if (deg >= 292.5 && deg < 337.5) { direction = new Vector3d(1, -1, 0).GetNormal(); return "SE"; }

            return "";
        }

        /// <summary>
        /// Computes area-weighted average slope and area-weighted mode direction.
        /// §6.2 — extracted from duplicated logic in GenerateSummaryTable and ExportToCSV.
        /// §7.1 — mode direction now weighted by overlap area, not by count.
        /// </summary>
        private WeightedSummary ComputeWeightedSummary(List<GridData> grids)
        {
            WeightedSummary result = new WeightedSummary();

            if (grids.Count == 0) return result;

            double totalWeights = grids.Sum(g => g.LappingArea);
            if (totalWeights > MinWeightThreshold)
            {
                result.MeanSlope = grids.Sum(g => g.SlopePercent * g.LappingArea) / totalWeights;
            }
            else
            {
                result.MeanSlope = grids.Average(g => g.SlopePercent);
            }

            // §7.1 — area-weighted mode direction instead of simple count
            var dirGroups = grids
                .Where(g => !string.IsNullOrEmpty(g.Direction))
                .GroupBy(g => g.Direction)
                .Select(g => new { Direction = g.Key, TotalArea = g.Sum(x => x.LappingArea) })
                .OrderByDescending(g => g.TotalArea);

            if (dirGroups.Any()) result.ModeDirection = dirGroups.First().Direction;

            return result;
        }

        #endregion

        #region AutoCAD Drawing Handlers

        /// <summary>
        /// Draws a 3×3 compass legend grid with directional arrows in each cardinal/intercardinal cell.
        /// </summary>
        private void DrawCompassLegend(Transaction tr, BlockTableRecord btr, CoordinateSystem3d cs,
            Matrix3d ucs, Point3d centerPtWCS, double sideLength, ObjectId layerEdges,
            ObjectId layerDirection, ObjectId layerDirText)
        {
            double textHeight = sideLength * TextHeightRatio;

            string[,] cardinalNames = new string[3, 3] {
                { "NW", "N", "NE" },
                { "W",  "COMPASS", "E" },
                { "SW", "S", "SE" }
            };

            Vector3d[,] cardinalDirs = new Vector3d[3, 3] {
                { new Vector3d(-1, 1, 0).GetNormal(), new Vector3d(0, 1, 0), new Vector3d(1, 1, 0).GetNormal() },
                { new Vector3d(-1, 0, 0),             new Vector3d(0, 0, 0), new Vector3d(1, 0, 0)             },
                { new Vector3d(-1, -1, 0).GetNormal(),new Vector3d(0, -1, 0),new Vector3d(1, -1, 0).GetNormal()}
            };

            for (int y = -1; y <= 1; y++)
            {
                for (int x = -1; x <= 1; x++)
                {
                    Point3d cWCS = centerPtWCS + (cs.Xaxis * (x * sideLength)) - (cs.Yaxis * (y * sideLength));
                    Point3d cUCS = cWCS.TransformBy(ucs.Inverse());

                    Polyline pl = new Polyline();
                    pl.LayerId = layerEdges;

                    Point3d p13d = cWCS - (cs.Xaxis * (sideLength / 2)) + (cs.Yaxis * (sideLength / 2));
                    Point3d p23d = cWCS + (cs.Xaxis * (sideLength / 2)) + (cs.Yaxis * (sideLength / 2));
                    Point3d p33d = cWCS + (cs.Xaxis * (sideLength / 2)) - (cs.Yaxis * (sideLength / 2));
                    Point3d p43d = cWCS - (cs.Xaxis * (sideLength / 2)) - (cs.Yaxis * (sideLength / 2));

                    pl.AddVertexAt(0, new Point2d(p13d.X, p13d.Y), 0, 0, 0);
                    pl.AddVertexAt(1, new Point2d(p23d.X, p23d.Y), 0, 0, 0);
                    pl.AddVertexAt(2, new Point2d(p33d.X, p33d.Y), 0, 0, 0);
                    pl.AddVertexAt(3, new Point2d(p43d.X, p43d.Y), 0, 0, 0);
                    pl.Closed = true;
                    pl.Elevation = 0;
                    btr.AppendEntity(pl);
                    tr.AddNewlyCreatedDBObject(pl, true);

                    int arrY = y + 1;
                    int arrX = x + 1;

                    if (x != 0 || y != 0)
                    {
                        Vector3d direction = cardinalDirs[arrY, arrX];
                        DrawDirectionArrow(tr, btr, cWCS, direction, sideLength, layerDirection);

                        Point3d textPt = new Point3d(cUCS.X, cUCS.Y - (sideLength * 0.40), 0).TransformBy(ucs);
                        AddTextToBTR(tr, btr, cardinalNames[arrY, arrX], textPt, textHeight * 0.8,
                            layerDirText, cs, AttachmentPoint.BottomCenter);
                    }
                    else
                    {
                        Point3d centerLabelPt = new Point3d(cUCS.X, cUCS.Y, 0).TransformBy(ucs);
                        AddTextToBTR(tr, btr, cardinalNames[arrY, arrX], centerLabelPt, textHeight * 0.8,
                            layerDirText, cs, AttachmentPoint.MiddleCenter);
                    }
                }
            }
        }

        /// <summary>
        /// Draws a detailed outline arrow from ArrowCode.txt template.
        /// Template points upward (+Y), scaled by cellSide, rotated to match slope direction,
        /// then translated to the specified centroid.
        /// </summary>
        private void DrawDirectionArrow(Transaction tr, BlockTableRecord btr, Point3d centroid,
            Vector3d dir, double cellSide, ObjectId layerId)
        {
            if (Math.Abs(dir.X) < 1e-8 && Math.Abs(dir.Y) < 1e-8) return;

            // Arrow template vertices (pointing +Y, tip at origin, tail at -0.625)
            double[][] template = new double[][] {
                new double[] { 0, 0 },
                new double[] { -0.0562, -0.325 },
                new double[] { -0.0125, -0.3125 },
                new double[] { -0.0125, -0.625 },
                new double[] { 0.0125, -0.625 },
                new double[] { 0.0125, -0.3125 },
                new double[] { 0.0563, -0.325 },
                new double[] { 0, 0 },
                new double[] { 0, -0.0586 },
                new double[] { 0.0437, -0.311 },
                new double[] { 0.0033, -0.2993 },
                new double[] { 0.0025, -0.615 },
                new double[] { -0.0025, -0.615 },
                new double[] { -0.0025, -0.2992 },
                new double[] { -0.0436, -0.311 },
                new double[] { 0, -0.0586 },
                new double[] { 0, -0.1173 },
                new double[] { 0.0312, -0.297 },
                new double[] { 0.0018, -0.2913 },
                new double[] { -0.0314, -0.2967 },
                new double[] { 0, -0.1173 },
                new double[] { 0.0004, -0.1704 },
                new double[] { 0.0207, -0.2859 },
                new double[] { -0.0184, -0.2846 },
                new double[] { 0.0004, -0.1664 },
                new double[] { 0.0004, -0.2161 },
                new double[] { 0.0134, -0.2792 },
                new double[] { -0.0141, -0.2792 },
                new double[] { 0.0004, -0.2147 },
                new double[] { 0.0004, -0.2483 },
                new double[] { 0.0062, -0.2738 },
                new double[] { -0.0054, -0.2738 },
                new double[] { 0.0004, -0.247 },
                new double[] { 0.0004, -0.2671 },
                new double[] { 0.0033, -0.2698 },
                new double[] { -0.0039, -0.2698 },
                new double[] { 0.0004, -0.2658 }
            };

            // Scale factor: template is ~0.625 tall, scale to fit cellSide
            double scale = cellSide * 0.65;

            // Rotation: template points +Y (90°), rotate to match dir
            double targetAngle = Math.Atan2(dir.Y, dir.X);
            double templateAngle = Math.PI / 2.0; // +Y
            double rotation = targetAngle - templateAngle;
            double cosR = Math.Cos(rotation);
            double sinR = Math.Sin(rotation);

            Polyline pl = new Polyline();
            pl.LayerId = layerId;

            // Center offset: place arrow by its center, not its tip
            double centerX = 0.0004;
            double centerY = -0.2483;

            for (int i = 0; i < template.Length; i++)
            {
                double sx = (template[i][0] - centerX) * scale;
                double sy = (template[i][1] - centerY) * scale;
                double rx = sx * cosR - sy * sinR;
                double ry = sx * sinR + sy * cosR;
                pl.AddVertexAt(i, new Point2d(centroid.X + rx, centroid.Y + ry), 0, 0, 0);
            }

            pl.Elevation = 0;

            btr.AppendEntity(pl);
            tr.AddNewlyCreatedDBObject(pl, true);
        }

        /// <summary>
        /// Highlights or unhighlights entities for interactive confirmation prompts.
        /// </summary>
        private void HighlightObjects(Transaction tr, ObjectId[] ids, bool highlight)
        {
            foreach (ObjectId id in ids)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null)
                {
                    if (highlight)
                        ent.Highlight();
                    else
                        ent.Unhighlight();
                }
            }
        }

        /// <summary>
        /// Returns a comma-separated list of unique layer names from a selection set.
        /// </summary>
        private string GetUniqueLayersFromSelection(Transaction tr, SelectionSet ss)
        {
            HashSet<string> layers = new HashSet<string>();
            foreach (SelectedObject so in ss)
            {
                Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                if (ent != null) layers.Add(ent.Layer);
            }
            return string.Join(",", layers);
        }

        /// <summary>
        /// Gets or creates a named layer with the specified ACI color index.
        /// </summary>
        private ObjectId GetOrCreateLayer(Database db, Transaction tr, string layerName, short colorIndex)
        {
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            if (lt.Has(layerName))
            {
                return lt[layerName];
            }
            else
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord();
                ltr.Name = layerName;
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
                    Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
                ObjectId newLayerId = lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
                return newLayerId;
            }
        }

        /// <summary>
        /// Creates an MText entity with Arial TrueType font, positioned and aligned in the current UCS.
        /// </summary>
        private void AddTextToBTR(Transaction tr, BlockTableRecord btr, string text, Point3d pos,
            double height, ObjectId layerId, CoordinateSystem3d cs, AttachmentPoint align)
        {
            MText mText = new MText();
            mText.Contents = "{\\fArial|b0|i0|c0|p0;" + text + "}";
            mText.TextHeight = height;
            mText.LayerId = layerId;
            try { mText.Attachment = align; }
            catch { mText.Attachment = AttachmentPoint.BottomLeft; }

            mText.Normal = cs.Zaxis;
            mText.Direction = cs.Xaxis;
            mText.Location = pos;

            btr.AppendEntity(mText);
            tr.AddNewlyCreatedDBObject(mText, true);
        }

        #endregion

        #region Data Table & CSV Export

        /// <summary>
        /// Returns the slope classification category for a given slope percentage.
        /// </summary>
        private string GetSlopeClassification(double slope)
        {
            if (slope <= 5.0) return "一級坡";
            if (slope <= 15.0) return "二級坡";
            if (slope <= 30.0) return "三級坡";
            if (slope <= 40.0) return "四級坡";
            if (slope <= 55.0) return "五級坡";
            if (slope <= 100.0) return "六級坡";
            return "七級坡";
        }

        /// <summary>
        /// Generates the reference legend table showing slope categories for n = 0 to 100 (even values).
        /// </summary>
        private void GenerateLegendTable(Transaction tr, BlockTableRecord btr, Database db,
            CoordinateSystem3d cs, Point3d insertPt, ObjectId layerTable,
            double contourInterval, double sideLength)
        {
            int totalItems = 51; // n = 0, 2, 4, ... 100
            int chunks = ((totalItems - 1) / MaxLegendRowsPerChunk) + 1;
            int totalCols = chunks * 3;
            int totalRows = MaxLegendRowsPerChunk + 2; // +1 title, +1 headers

            Table tb = new Table();
            tb.TableStyle = db.Tablestyle;
            tb.LayerId = layerTable;

            tb.Position = insertPt;
            tb.Normal = cs.Zaxis;
            tb.Direction = cs.Xaxis;

            tb.SetSize(totalRows, totalCols);

            for (int r = 0; r < totalRows; r++)
            {
                tb.Rows[r].Height = TableRowHeight;
                for (int c = 0; c < totalCols; c++)
                {
                    tb.Cells[r, c].TextHeight = TableTextHeight;
                    tb.Cells[r, c].Alignment = CellAlignment.MiddleCenter;
                }
            }

            tb.MergeCells(CellRange.Create(tb, 0, 0, 0, totalCols - 1));
            tb.Cells[0, 0].TextString = "Slope Categories Reference Legend (n = 0 to 100)";

            for (int c = 0; c < chunks; c++)
            {
                int baseCol = c * 3;
                tb.Columns[baseCol].Width = 16.0;
                tb.Columns[baseCol + 1].Width = 15.0;
                tb.Columns[baseCol + 2].Width = 18.0;

                tb.Cells[1, baseCol].TextString = "Intersects (n)";
                tb.Cells[1, baseCol + 1].TextString = "Theoretical Slope (%)";
                tb.Cells[1, baseCol + 2].TextString = "Class";
            }

            int index = 0;
            for (int n = 0; n <= 100; n += 2)
            {
                int chunkIndex = index / MaxLegendRowsPerChunk;
                int localRow = (index % MaxLegendRowsPerChunk) + 2;
                int baseCol = chunkIndex * 3;

                double slopePercent = ((n * Math.PI * contourInterval) / (8.0 * sideLength)) * 100.0;
                string classification = GetSlopeClassification(slopePercent);

                tb.Cells[localRow, baseCol].TextString = n.ToString();
                tb.Cells[localRow, baseCol + 1].TextString = string.Format("{0:F2}%", slopePercent);
                tb.Cells[localRow, baseCol + 2].TextString = classification;

                index++;
            }

            tb.GenerateLayout();
            btr.AppendEntity(tb);
            tr.AddNewlyCreatedDBObject(tb, true);
        }

        /// <summary>
        /// Generates the summary results table with per-grid data and a weighted-average footer.
        /// </summary>
        private void GenerateSummaryTable(Transaction tr, BlockTableRecord btr, Database db,
            CoordinateSystem3d cs, List<GridData> validGrids, Point3d insertPt, ObjectId layerTable)
        {
            // Guard: if no grids, still produce a 1-chunk table
            int gridCount = Math.Max(validGrids.Count, 1);
            int chunks = ((gridCount - 1) / MaxSummaryRowsPerChunk) + 1;
            int totalCols = chunks * 6;
            int totalRows = Math.Min(validGrids.Count, MaxSummaryRowsPerChunk) + 3; // +2 headers, +1 summary

            Table tb = new Table();
            tb.TableStyle = db.Tablestyle;
            tb.LayerId = layerTable;

            tb.Position = insertPt;
            tb.Normal = cs.Zaxis;
            tb.Direction = cs.Xaxis;

            tb.SetSize(totalRows, totalCols);

            for (int r = 0; r < totalRows; r++)
            {
                tb.Rows[r].Height = TableRowHeight;
                for (int c = 0; c < totalCols; c++)
                {
                    tb.Cells[r, c].TextHeight = TableTextHeight;
                    tb.Cells[r, c].Alignment = CellAlignment.MiddleCenter;
                }
            }

            tb.MergeCells(CellRange.Create(tb, 0, 0, 0, totalCols - 1));
            tb.Cells[0, 0].TextString = "Slope Calculation Summary";

            for (int c = 0; c < chunks; c++)
            {
                int baseCol = c * 6;

                tb.Columns[baseCol].Width = 10.0;
                tb.Columns[baseCol + 1].Width = 16.0;
                tb.Columns[baseCol + 2].Width = 15.0;
                tb.Columns[baseCol + 3].Width = 18.0;
                tb.Columns[baseCol + 4].Width = 14.0;
                tb.Columns[baseCol + 5].Width = 15.0;

                tb.Cells[1, baseCol].TextString = "Grid ID";
                tb.Cells[1, baseCol + 1].TextString = "Intersects (n)";
                tb.Cells[1, baseCol + 2].TextString = "Slope (%)";
                tb.Cells[1, baseCol + 3].TextString = "Class";
                tb.Cells[1, baseCol + 4].TextString = "Direction";
                tb.Cells[1, baseCol + 5].TextString = "Area (m\u00B2)";
            }

            for (int i = 0; i < validGrids.Count; i++)
            {
                int chunkIndex = i / MaxSummaryRowsPerChunk;
                int localRow = (i % MaxSummaryRowsPerChunk) + 2;
                int baseCol = chunkIndex * 6;

                tb.Cells[localRow, baseCol].TextString = (i + 1).ToString();
                tb.Cells[localRow, baseCol + 1].TextString = validGrids[i].TotalIntersects.ToString();
                tb.Cells[localRow, baseCol + 2].TextString = string.Format("{0:F2}%", validGrids[i].SlopePercent);
                tb.Cells[localRow, baseCol + 3].TextString = validGrids[i].Classification;
                tb.Cells[localRow, baseCol + 4].TextString = validGrids[i].Direction;
                tb.Cells[localRow, baseCol + 5].TextString = string.Format("{0:F2}", validGrids[i].LappingArea);
            }

            // §6.2 — uses extracted helper
            WeightedSummary summary = ComputeWeightedSummary(validGrids);

            int summaryRow = totalRows - 1;
            tb.MergeCells(CellRange.Create(tb, summaryRow, 0, summaryRow, totalCols - 1));
            tb.Cells[summaryRow, 0].TextString = string.Format(
                "Area-Weighted Average Slope: {0:F2}%   |   Primary Direction: {1}",
                summary.MeanSlope, summary.ModeDirection);

            tb.GenerateLayout();
            btr.AppendEntity(tb);
            tr.AddNewlyCreatedDBObject(tb, true);
        }

        /// <summary>
        /// Exports grid results to a UTF-8 BOM CSV file alongside the current drawing.
        /// </summary>
        private void ExportToCSV(Document doc, List<GridData> grids)
        {
            try
            {
                string filePath;
                if (!string.IsNullOrEmpty(doc.Name) && Path.IsPathRooted(doc.Name))
                {
                    string dir = Path.GetDirectoryName(doc.Name);
                    string name = Path.GetFileNameWithoutExtension(doc.Name);
                    filePath = Path.Combine(dir, name + "_slopedata.csv");
                }
                else
                {
                    filePath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "slopedata.csv");
                }

                // §6.2 — uses extracted helper
                WeightedSummary summary = ComputeWeightedSummary(grids);

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Grid ID,Intersects (n),Slope (%),Classification,Direction,Area (sq m)");
                for (int i = 0; i < grids.Count; i++)
                {
                    GridData g = grids[i];
                    sb.AppendLine(string.Format("{0},{1},{2:F2},{3},{4},{5:F2}",
                        i + 1, g.TotalIntersects, g.SlopePercent, g.Classification, g.Direction, g.LappingArea));
                }

                sb.AppendLine();
                sb.AppendLine(string.Format("Area-Weighted Average Slope (%),{0:F2}", summary.MeanSlope));
                sb.AppendLine(string.Format("Primary Direction,{0}", summary.ModeDirection));

                File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(true));
                doc.Editor.WriteMessage(string.Format("\n[+] CSV exported to: {0}", filePath));
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage(string.Format("\n[-] CSV export failed: {0}", ex.Message));
            }
        }

        #endregion
    }
}
