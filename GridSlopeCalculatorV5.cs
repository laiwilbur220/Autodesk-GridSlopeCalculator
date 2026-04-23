using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Civil3DGridMethod
{
    /// <summary>
    /// V5 Grid Slope Calculator ??Automated terrain slope analysis for AutoCAD Civil 3D.
    /// Features:
    ///   ??Optional automatic grid generation aligned to UCS (25m or 10m cells)
    ///   ??Slope calculation via Horton grid method: S = n??h / 8L ? 100
    ///   ??Downhill aspect via least-squares plane fitting
    ///   ??Detailed outline direction arrows with center-point positioning
    ///   ??Transparent heatmap hatching by slope classification
    ///   ??NOD-based parameter persistence for update command
    ///   ??Native XLSX export with live Excel formulas
    /// </summary>
    public class GridSlopeCalculatorV5
    {
        #region Constants

        /// <summary>Tolerance for deduplicating intersection points (meters).</summary>
        private const double IntersectTolerance = 0.02;

        /// <summary>Fractional tolerance for validating grid area against expected square area.</summary>
        private const double AreaValidationTolerance = 0.05;

        /// <summary>Minimum overlap area (m蝪? to consider a grid inside the boundary.</summary>
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

        // NOD (Named Object Dictionary) key for persistent parameter storage
        private const string NOD_KEY = "GridSlopeParams";

        // Slope classification layers (class name, layer name, ACI color index)
        private static readonly string[] SlopeClassNames = { "\u4E00\u7D1A\u5761", "\u4E8C\u7D1A\u5761", "\u4E09\u7D1A\u5761", "\u56DB\u7D1A\u5761", "\u4E94\u7D1A\u5761", "\u516D\u7D1A\u5761", "\u4E03\u7D1A\u5761" };
        private static readonly string[] SlopeClassLayerNames = {
            "SLOPE_CLASS_01_Green", "SLOPE_CLASS_02_YGreen", "SLOPE_CLASS_03_Yellow",
            "SLOPE_CLASS_04_Orange", "SLOPE_CLASS_05_ROrange", "SLOPE_CLASS_06_Red",
            "SLOPE_CLASS_07_DarkRed"
        };
        private static readonly short[] SlopeClassACI = { 82, 52, 42, 30, 20, 10, 14 };

        private static readonly string[] SlopeClassRanges = {
            "S\u22665%", "5%<S\u226615%", "15%<S\u226630%", "30%<S\u226640%",
            "40%<S\u226655%", "55%<S\u2266100%", "S>100%"
        };

        private static readonly string[] DirLabels = { "\u6771(E)", "\u5317(N)", "\u897F(W)", "\u5357(S)", "\u6771\u5317(NE)", "\u6771\u5357(SE)", "\u897F\u5317(NW)", "\u897F\u5357(SW)" };
        private static readonly string[] DirKeys = { "E", "N", "W", "S", "NE", "SE", "NW", "SW" };

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

            /// <summary>Slope classification category (??城豰?through ????.</summary>
            public string Classification { get; set; }

            /// <summary>Compass direction label for the downhill aspect (e.g. "NE", "S").</summary>
            public string Direction { get; set; }

            public GridData()
            {
                Classification = "";
                Direction = "";
            }

            /// <summary>Area of intersection between grid cell and project boundary (m蝪?.</summary>
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
        /// Main entry point ??AutoCAD command: CalcGridSlopeCSV5
        /// Workflow: grid existence check ??(optional grid generation) ??slope calculation ??output.
        /// </summary>
        [CommandMethod("CalcGridSlopeCSV5")]
        public void CalculateGridSlope()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // ===== STEP 0: Check if grids need to be built =====
                PromptKeywordOptions pkoGrid = new PromptKeywordOptions(
                    "\nHas grid already been built? [Y/N] ", "Y N");
                pkoGrid.Keywords.Default = "Y";
                pkoGrid.AllowNone = true;
                PromptResult prGrid = ed.GetKeywords(pkoGrid);
                if (prGrid.Status != PromptStatus.OK && prGrid.Status != PromptStatus.None) return;
                bool gridAlreadyBuilt = (prGrid.StringResult == "Y");

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
                            "\nConfirm selected project boundary? [Y/N] ", "Y N");
                        pkoVerify.Keywords.Default = "Y";
                        PromptResult prVerify = ed.GetKeywords(pkoVerify);
                        HighlightObjects(trVerify, verifyBoundIds, false);
                        trVerify.Commit();
                        if (prVerify.Status != PromptStatus.OK || prVerify.StringResult != "Y") return;
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

                    PromptKeywordOptions pkoEdit = new PromptKeywordOptions(
                        "\nDo you want to edit the generated grids before calculating slope? [Y/N] ", "Y N");
                    pkoEdit.Keywords.Default = "N";
                    pkoEdit.AllowNone = true;
                    PromptResult prEdit = ed.GetKeywords(pkoEdit);
                    if (prEdit.Status == PromptStatus.OK && prEdit.StringResult == "Y")
                    {
                        ed.WriteMessage("\nPlease edit the generated grids. When finished, re-run the command and choose 'Yes' for 'Has grid already been built?'.");
                        return;
                    }
                }
                else
                {
                    // Existing grids ??ask side length manually
                    sideLength = PromptForDouble(ed, "\nEnter grid side length (L) in meters: ", 25.0);
                    if (sideLength <= 0) return;
                }

                // ===== SLOPE CALCULATION PHASE =====
                double contourInterval = PromptForDouble(ed, "\nEnter contour interval (?h) in meters: ", 1.0);
                if (contourInterval <= 0) return;

                PromptKeywordOptions pkoCsv = new PromptKeywordOptions(
                    "\nExport data to XLSX? [Y/N] ", "Y N");
                pkoCsv.AllowNone = true;
                pkoCsv.Keywords.Default = "N";
                PromptResult prExport = ed.GetKeywords(pkoCsv);
                if (prExport.Status != PromptStatus.OK && prExport.Status != PromptStatus.None) return;
                bool shouldExportXlsx = (prExport.StringResult == "Y");

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
                            "\nUse same boundary from grid generation? [Y/N] ", "Y N");
                        pkoBound.Keywords.Default = "Y";
                        PromptResult prBound = ed.GetKeywords(pkoBound);
                        HighlightObjects(tr, boundIds, false);
                        if (prBound.Status != PromptStatus.OK || prBound.StringResult != "Y") return;
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
                            "\nConfirm selected project boundary? [Y/N] ", "Y N");
                        pkoBound.Keywords.Default = "Y";
                        PromptResult prBound = ed.GetKeywords(pkoBound);
                        HighlightObjects(tr, boundIds, false);
                        if (prBound.Status != PromptStatus.OK || prBound.StringResult != "Y") return;
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
                    ObjectId layerArea = GetOrCreateLayer(db, tr, "Grid_Outputs_Area", 7);
                    ObjectId layerTable = GetOrCreateLayer(db, tr, "Grid_Outputs_Table", 7);

                    // Create slope classification heatmap layers
                    ObjectId[] slopeClassLayerIds = new ObjectId[7];
                    for (int sci = 0; sci < 7; sci++)
                    {
                        slopeClassLayerIds[sci] = GetOrCreateLayer(db, tr, SlopeClassLayerNames[sci], SlopeClassACI[sci]);
                    }

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

                        // Use 50% of the side length to safely cluster rows even if hand-drawn grids are slightly rotated
                        double snapTol = sideLength * 0.5;
                        preValidGrids = preValidGrids
                            .OrderByDescending(g => Math.Round(g.SortY / snapTol) * snapTol)
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

                            // Direction text ??top-center (swapped with area)
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

                            // Overlap area ??bottom of cell (swapped with direction)
                            {
                                Point3d areaPt = new Point3d(cUCS.X,
                                    cUCS.Y - (sideLength * DirectionTextOffsetYRatio), 0).TransformBy(ucs);
                                string areaText = string.Format("A={0:F0} m\u00B2", currentGrid.LappingArea);
                                AddTextToBTR(tr, btr, areaText, areaPt, textHeight * 0.8,
                                    layerArea, cs, AttachmentPoint.BottomCenter);
                            }

                            validGrids.Add(currentGrid);

                            // Create heatmap hatch for this grid cell
                            if (currentGrid.TotalIntersects >= 0)
                            {
                                int classIdx = GetSlopeClassIndex(currentGrid.SlopePercent);
                                try
                                {
                                    Hatch hatch = new Hatch();
                                    hatch.SetDatabaseDefaults();
                                    hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                                    hatch.LayerId = slopeClassLayerIds[classIdx];
                                    hatch.Transparency = new Transparency(127); // 50% transparent
                                    hatch.Normal = Vector3d.ZAxis;
                                    hatch.Elevation = 0.0;
                                    btr.AppendEntity(hatch);
                                    tr.AddNewlyCreatedDBObject(hatch, true);
                                    hatch.Associative = true;
                                    hatch.AppendLoop(HatchLoopTypes.Default,
                                        new ObjectIdCollection(new ObjectId[] { currentGrid.Id }));
                                    hatch.EvaluateHatch(true);

                                    // Move hatch to bottom of draw order
                                    DrawOrderTable dot = (DrawOrderTable)tr.GetObject(
                                        btr.DrawOrderTableId, OpenMode.ForWrite);
                                    dot.MoveToBottom(new ObjectIdCollection(new ObjectId[] { hatch.ObjectId }));
                                }
                                catch { /* hatch may fail on degenerate cells */ }
                            }
                        }

                        // ------------- STEP 5: OUTPUT GENERATION -------------
                        ObjectId summaryTableId = GenerateSummaryTable(tr, btr, db, cs, validGrids, tableInsertPtWCS, layerTable);

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

                        if (shouldExportXlsx) ExportToXLSX(doc, validGrids, sideLength, contourInterval);

                        // Save params to NOD for UpdateGridSlopeCSV5
                        SaveParamsToNOD(db, tr, sideLength, contourInterval, summaryTableId);

                        tr.Commit();
                        ed.WriteMessage(string.Format("\nSuccess! Processed {0} grids.", validGrids.Count));
                        ed.WriteMessage("\ngrid data has correctly updated\n");
                        
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
                    string.Format("\nFound {0} items on layer(s) [{1}]. Is this correct? [Y/N] ", ids.Length, layers), "Y N");
                pko.Keywords.Default = "Y";

                PromptResult pr = ed.GetKeywords(pko);
                HighlightObjects(tr, ids, false);

                if (pr.Status != PromptStatus.OK) return null;
                if (pr.StringResult == "Y") return psrAll;
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

            double buffer = 0.05 * sideLength;

            // Collect all vertices for optimal shift calculation
            List<double> modX = new List<double>();
            List<double> modY = new List<double>();

            if (boundaryCurve is Polyline)
            {
                Polyline currPoly = (Polyline)boundaryCurve;
                for (int i = 0; i < currPoly.NumberOfVertices; i++)
                {
                    Point3d pt = currPoly.GetPoint3dAt(i).TransformBy(ucsInv);
                    modX.Add((pt.X % sideLength + sideLength) % sideLength);
                    modY.Add((pt.Y % sideLength + sideLength) % sideLength);
                }
            }
            else if (boundaryCurve is Polyline2d)
            {
                Polyline2d p2d = (Polyline2d)boundaryCurve;
                foreach (ObjectId vid in p2d)
                {
                    Vertex2d v = tr.GetObject(vid, OpenMode.ForRead) as Vertex2d;
                    Point3d pt = p2d.VertexPosition(v).TransformBy(ucsInv);
                    modX.Add((pt.X % sideLength + sideLength) % sideLength);
                    modY.Add((pt.Y % sideLength + sideLength) % sideLength);
                }
            }
            else if (boundaryCurve is Polyline3d)
            {
                Polyline3d p3d = (Polyline3d)boundaryCurve;
                foreach (ObjectId vid in p3d)
                {
                    PolylineVertex3d v = tr.GetObject(vid, OpenMode.ForRead) as PolylineVertex3d;
                    Point3d pt = v.Position.TransformBy(ucsInv);
                    modX.Add((pt.X % sideLength + sideLength) % sideLength);
                    modY.Add((pt.Y % sideLength + sideLength) % sideLength);
                }
            }

            if (modX.Count == 0)
            {
                modX.Add((ucsMinX % sideLength + sideLength) % sideLength);
                modX.Add((ucsMaxX % sideLength + sideLength) % sideLength);
                modY.Add((ucsMinY % sideLength + sideLength) % sideLength);
                modY.Add((ucsMaxY % sideLength + sideLength) % sideLength);
            }

            double shiftX = FindOptimalGridShift(modX, sideLength);
            double shiftY = FindOptimalGridShift(modY, sideLength);

            double diffX = (ucsMinX % sideLength + sideLength) % sideLength - shiftX;
            if (diffX < 0) diffX += sideLength;
            double startX = ucsMinX - diffX;
            if (ucsMinX - startX < buffer) startX -= sideLength;

            double diffY = (ucsMinY % sideLength + sideLength) % sideLength - shiftY;
            if (diffY < 0) diffY += sideLength;
            double startY = ucsMinY - diffY;
            if (ucsMinY - startY < buffer) startY -= sideLength;

            // Create boundary region for overlap testing
            Region boundaryRegion = CreateFlatRegionFromCurve(boundaryCurve, ed);
            if (boundaryRegion == null) return 0;

            ObjectId gridLayerId = GetOrCreateLayer(db, tr, gridLayerName, 180);
            int validCount = 0;

            try
            {
                int cols = (int)Math.Ceiling((ucsMaxX - startX + buffer) / sideLength);
                int rows = (int)Math.Ceiling((ucsMaxY - startY + buffer) / sideLength);
                ed.WriteMessage(string.Format("\nBuilding {0} x {1} optimally shifted grid pattern...", cols, rows));

                for (int row = 0; row < rows; row++)
                {
                    for (int col = 0; col < cols; col++)
                    {
                        double x = startX + col * sideLength;
                        double y = startY + row * sideLength;

                        // Grid corners in UCS ??transformed to WCS
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

                        // Check overlap with boundary ??erase if no intersection
                        // Strict limit to project boundary (move algorithm ensures no sticking)
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
                        // 蝳?.1 ??Polyline3d has no Elevation property; vertices are 3D.
                        // Region.CreateFromCurves may fail if non-planar ??catch handles with clear message.
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
            Region gridRegion = null;
            try
            {
                // Use CreateFlatRegionFromCurve to ensure the grid is fully planar at Z=0,
                // fixing Boolean intersect failures for pre-existing grids drawn at elevations.
                gridRegion = CreateFlatRegionFromCurve(pline, ed);
                if (gridRegion != null)
                {
                    using (Region boundClone = boundaryRegion.Clone() as Region)
                    {
                        gridRegion.BooleanOperation(BooleanOperationType.BoolIntersect, boundClone);
                        return gridRegion.Area;
                    }
                }
            }
            catch (System.Exception ex)
            {
                // log instead of silently swallowing
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
        /// Finds the mathematically optimal shift offset [0, L) that maximizes the
        /// minimum distance between grid lines and all vertices in the project boundary.
        /// This intelligently prevents grid lines from perfectly overlapping bounds and sticking.
        /// </summary>
        private double FindOptimalGridShift(List<double> mods, double L)
        {
            if (mods.Count == 0) return 0;
            
            List<double> uniqueMods = mods.Distinct().ToList();
            uniqueMods.Sort();

            if (uniqueMods.Count == 1) return (uniqueMods[0] + L / 2) % L;

            double maxGap = L - uniqueMods.Last() + uniqueMods.First();
            double bestShift = (uniqueMods.Last() + maxGap / 2) % L;

            for (int i = 0; i < uniqueMods.Count - 1; i++)
            {
                double gap = uniqueMods[i + 1] - uniqueMods[i];
                if (gap > maxGap)
                {
                    maxGap = gap;
                    bestShift = uniqueMods[i] + gap / 2;
                }
            }
            return bestShift;
        }

        /// <summary>
        /// Determines downhill direction via least-squares plane fit: z = a(x???) + b(y??? + z?.
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
        /// 蝳?.2 ??extracted from duplicated logic in GenerateSummaryTable and ExportToCSV.
        /// 蝳?.1 ??mode direction now weighted by overlap area, not by count.
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

            // 蝳?.1 ??area-weighted mode direction instead of simple count
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
        /// Draws a 3?3 compass legend grid with directional arrows in each cardinal/intercardinal cell.
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

            // Rotation: template points +Y (90蝪?, rotate to match dir
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

        #region Data Table & XLSX Export

        /// <summary>
        /// Returns the slope classification category for a given slope percentage.
        /// </summary>
        private string GetSlopeClassification(double slope)
        {
            if (slope <= 5.0) return "\u4e00\u7d1a\u5761";
            if (slope <= 15.0) return "\u4e8c\u7d1a\u5761";
            if (slope <= 30.0) return "\u4e09\u7d1a\u5761";
            if (slope <= 40.0) return "\u56db\u7d1a\u5761";
            if (slope <= 55.0) return "\u4e94\u7d1a\u5761";
            if (slope <= 100.0) return "\u516d\u7d1a\u5761";
            return "\u4e03\u7d1a\u5761";
        }

        /// <summary>
        /// Returns 0-based slope class index (0=\u4e00\u7d1a\u5761 .. 6=\u4e03\u7d1a\u5761).
        /// </summary>
        private int GetSlopeClassIndex(double slope)
        {
            if (slope <= 5.0) return 0;
            if (slope <= 15.0) return 1;
            if (slope <= 30.0) return 2;
            if (slope <= 40.0) return 3;
            if (slope <= 55.0) return 4;
            if (slope <= 100.0) return 5;
            return 6;
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
        private ObjectId GenerateSummaryTable(Transaction tr, BlockTableRecord btr, Database db,
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

            // 蝳?.2 ??uses extracted helper
            WeightedSummary summary = ComputeWeightedSummary(validGrids);

            int summaryRow = totalRows - 1;
            tb.MergeCells(CellRange.Create(tb, summaryRow, 0, summaryRow, totalCols - 1));
            tb.Cells[summaryRow, 0].TextString = string.Format(
                "Area-Weighted Average Slope: {0:F2}%   |   Primary Direction: {1}",
                summary.MeanSlope, summary.ModeDirection);

            tb.GenerateLayout();
            btr.AppendEntity(tb);
            tr.AddNewlyCreatedDBObject(tb, true);
            return tb.ObjectId;
        }


           /// <summary>
        /// Saves grid parameters (L, ?h) and Table ObjectId to the drawing's Named Object Dictionary.
        /// </summary>
        private void SaveParamsToNOD(Database db, Transaction tr, double sideLength, double contourInterval, ObjectId tableId)
        {
            DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);
            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.Real, sideLength),
                new TypedValue((int)DxfCode.Real, contourInterval),
                new TypedValue((int)DxfCode.SoftPointerId, tableId)
            );

            if (nod.Contains(NOD_KEY))
            {
                Xrecord old = (Xrecord)tr.GetObject(nod.GetAt(NOD_KEY), OpenMode.ForWrite);
                old.Data = rb;
            }
            else
            {
                Xrecord xrec = new Xrecord();
                xrec.Data = rb;
                nod.SetAt(NOD_KEY, xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }
        }

        /// <summary>
        /// Loads grid parameters from the drawing's Named Object Dictionary.
        /// Returns true if found.
        /// </summary>
        private bool LoadParamsFromNOD(Database db, Transaction tr, out double sideLength, out double contourInterval, out ObjectId tableId)
        {
            sideLength = 0;
            contourInterval = 0;
            tableId = ObjectId.Null;
            tableId = ObjectId.Null;
            DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
            if (nod.Contains(NOD_KEY))
            {
                Xrecord xrec = (Xrecord)tr.GetObject(nod.GetAt(NOD_KEY), OpenMode.ForRead);
                TypedValue[] values = xrec.Data.AsArray();
                if (values.Length >= 2)
                {
                    sideLength = (double)values[0].Value;
                    contourInterval = (double)values[1].Value;
                    if (values.Length >= 3) {
                        tableId = (ObjectId)values[2].Value;
                    }
                    return true;
                }
            }
            return false;
        }

        #endregion

        #region Update Command

        /// <summary>
        /// Spatial update command: re-syncs existing MText + Hatch inside each grid cell
        /// after contour edits, using Editor.SelectCrossingPolygon.
        /// </summary>
        [CommandMethod("UpdateGridSlopeCSV5")]
        public void UpdateGridSlope()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    double sideLength, contourInterval;
                    ObjectId tableId = ObjectId.Null;
                    if (!LoadParamsFromNOD(db, tr, out sideLength, out contourInterval, out tableId))
                    {
                        ed.WriteMessage("\nNo saved parameters found. Run CalcGridSlopeCSV5 first.");
                        return;
                    }
                    ed.WriteMessage(string.Format("\nLoaded from NOD: L={0}, ?h={1}", sideLength, contourInterval));

                    // Select grids
                    PromptSelectionResult psrAllGrids = GetSelectionByLayer(ed, tr, "GRID(s)", "LWPOLYLINE,POLYLINE");
                    if (psrAllGrids == null) return;

                    // Select boundary
                    PromptEntityOptions peo = new PromptEntityOptions(
                        "\nSelect the Project Boundary (Polyline): ");
                    peo.SetRejectMessage("\nBoundary must be a Polyline!");
                    peo.AddAllowedClass(typeof(Polyline), false);
                    peo.AddAllowedClass(typeof(Polyline2d), false);
                    peo.AddAllowedClass(typeof(Polyline3d), false);
                    PromptEntityResult perBoundary = ed.GetEntity(peo);
                    if (perBoundary.Status != PromptStatus.OK) return;

                    Matrix3d ucs = ed.CurrentUserCoordinateSystem;
                    CoordinateSystem3d cs = ucs.CoordinateSystem3d;

                    Curve baseBoundaryCurve = tr.GetObject(perBoundary.ObjectId, OpenMode.ForRead) as Curve;
                    if (baseBoundaryCurve == null) { ed.WriteMessage("\nFailed to load boundary."); return; }

                    Region boundaryRegion = CreateFlatRegionFromCurve(baseBoundaryCurve, ed);
                    if (boundaryRegion == null) { ed.WriteMessage("\nERROR: Failed to create boundary region."); return; }

                    try
                    {
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(
                            bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                        // Create/get heatmap layers
                        ObjectId[] slopeClassLayerIds = new ObjectId[7];
                        for (int sci = 0; sci < 7; sci++)
                            slopeClassLayerIds[sci] = GetOrCreateLayer(db, tr, SlopeClassLayerNames[sci], SlopeClassACI[sci]);

                        List<GridData> preValidGrids = GetValidGrids(tr, psrAllGrids, sideLength, ucs, ed);
                        if (preValidGrids.Count == 0) return;

                        double snapTol = sideLength * 0.5;
                        preValidGrids = preValidGrids
                            .OrderByDescending(g => Math.Round(g.SortY / snapTol) * snapTol)
                            .ThenBy(g => g.SortX)
                            .ToList();

                        List<GridData> validGrids = new List<GridData>();
                        Regex rxN = new Regex(@"n=\d+");
                        HashSet<ObjectId> consumedEdgeTexts = new HashSet<ObjectId>();

                        // Gag associative warnings during massive DOM manipulations
                        try { Application.SetSystemVariable("NOMUTT", 1); } catch { }
                        
                        try 
                        {
                            for (int i = 0; i < preValidGrids.Count; i++)
                        {
                            GridData currentGrid = preValidGrids[i];
                            Polyline pline = tr.GetObject(currentGrid.Id, OpenMode.ForRead) as Polyline;
                            if (pline == null || pline.NumberOfVertices < 4) continue;

                            currentGrid.LappingArea = CalculateOverlapArea(pline, boundaryRegion, ed);
                            if (currentGrid.LappingArea <= MinOverlapArea) continue;

                            // Build crossing polygon vertices for SelectCrossingPolygon
                            Point3dCollection cellVerts = new Point3dCollection();
                            for (int v = 0; v < pline.NumberOfVertices; v++)
                                cellVerts.Add(pline.GetPoint3dAt(v));

                            // ----- SPATIAL DOM SCRAPING -----
                            // Instead of contour intersects, scrape user edits from the CAD layout directly
                            
                            int scrapedTotalIntersects = 0;
                            double sideTol = sideLength * 0.15; // Small buffer window (15% of side length)

                            // 1. Scrape 'n counts' from the 4 grid edges
                            for (int j = 0; j < 4; j++)
                            {
                                Point3d pt1 = pline.GetPoint3dAt(j);
                                Point3d pt2 = pline.GetPoint3dAt((j + 1) % 4);
                                Point3d midPt = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);

                                // Small crossing polygon around midpoint
                                Point3dCollection edgeCrossPts = new Point3dCollection();
                                edgeCrossPts.Add(new Point3d(midPt.X - sideTol, midPt.Y - sideTol, 0));
                                edgeCrossPts.Add(new Point3d(midPt.X + sideTol, midPt.Y - sideTol, 0));
                                edgeCrossPts.Add(new Point3d(midPt.X + sideTol, midPt.Y + sideTol, 0));
                                edgeCrossPts.Add(new Point3d(midPt.X - sideTol, midPt.Y + sideTol, 0));

                                TypedValue[] edgeFilterArray = new TypedValue[] {
                                    new TypedValue((int)DxfCode.Start, "MTEXT")
                                };
                                SelectionFilter edgeSf = new SelectionFilter(edgeFilterArray);
                                PromptSelectionResult psrEdge = ed.SelectCrossingPolygon(edgeCrossPts, edgeSf);

                                if (psrEdge.Status == PromptStatus.OK)
                                {
                                    foreach (SelectedObject so in psrEdge.Value)
                                    {
                                        MText mt = tr.GetObject(so.ObjectId, OpenMode.ForRead) as MText;
                                        if (mt != null)
                                        {
                                            // Extract raw integer from any MText formatting
                                            string rawText = Regex.Replace(mt.Contents, @"{\\.*?\\|.*?;|}", "");
                                            int parsedVal = 0;
                                            if (int.TryParse(rawText.Trim(), out parsedVal))
                                            {
                                                scrapedTotalIntersects += parsedVal;
                                                break; // Only capture one MText count per edge
                                            }
                                        }
                                    }
                                }
                            }
                            
                            currentGrid.TotalIntersects = scrapedTotalIntersects;

                            // 2. Scrape Direction Override
                            string scrapedDir = "";
                            TypedValue[] dirFilterArray = new TypedValue[] {
                                new TypedValue((int)DxfCode.Start, "MTEXT")
                            };
                            SelectionFilter dirSf = new SelectionFilter(dirFilterArray);
                            PromptSelectionResult psrDir = ed.SelectCrossingPolygon(cellVerts, dirSf);

                            if (psrDir.Status == PromptStatus.OK)
                            {
                                string[] validDirs = { "N", "NE", "E", "SE", "S", "SW", "W", "NW" };
                                foreach (SelectedObject so in psrDir.Value)
                                {
                                    MText mt = tr.GetObject(so.ObjectId, OpenMode.ForRead) as MText;
                                    if (mt != null)
                                    {
                                        string rawText = Regex.Replace(mt.Contents, @"{\\.*?\\|.*?;|}", "").Trim().ToUpper();
                                        if (validDirs.Contains(rawText))
                                        {
                                            scrapedDir = rawText;
                                            break; // Found the user's direction text
                                        }
                                    }
                                }
                            }
                            currentGrid.Direction = scrapedDir;
                            currentGrid.SlopePercent = ((currentGrid.TotalIntersects * Math.PI * contourInterval)
                                / (8 * sideLength)) * 100;
                            currentGrid.Classification = GetSlopeClassification(currentGrid.SlopePercent);

                            // Update MText inside cell via crossing polygon
                            try
                            {
                                TypedValue[] filterArray = new TypedValue[] {
                                    new TypedValue((int)DxfCode.Start, "MTEXT")
                                };
                                SelectionFilter sf = new SelectionFilter(filterArray);
                                PromptSelectionResult psrInCell = ed.SelectCrossingPolygon(cellVerts, sf);

                                if (psrInCell.Status == PromptStatus.OK)
                                {
                                    foreach (SelectedObject so in psrInCell.Value)
                                    {
                                        MText mt = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as MText;
                                        if (mt == null) continue;
                                        string contents = mt.Contents;

                                        // Update n= text
                                        if (rxN.IsMatch(contents))
                                        {
                                            mt.Contents = "{\\fArial|b0|i0|c0|p0;n=" + currentGrid.TotalIntersects + "}";
                                        }
                                        // Update slope% text
                                        else if (contents.Contains("%") && !contents.Contains("("))
                                        {
                                            mt.Contents = "{\\fArial|b0|i0|c0|p0;" + string.Format("{0:F2}%", currentGrid.SlopePercent) + "}";
                                        }
                                        // Update classification text
                                        else if (contents.Contains("(") && contents.Contains("\u7D1A\u5761"))
                                        {
                                            mt.Contents = "{\\fArial|b0|i0|c0|p0;(" + currentGrid.Classification + ")}";
                                        }
                                        // Normalize overwritten direction string formatting natively 
                                        else
                                        {
                                            string rawMtDir = Regex.Replace(contents, @"{\\.*?\\|.*?;|}", "").Trim().ToUpper();
                                            string[] vDirs = { "N", "NE", "E", "SE", "S", "SW", "W", "NW" };
                                            if (vDirs.Contains(rawMtDir))
                                            {
                                                mt.Contents = "{\\fArial|b0|i0|c0|p0;" + currentGrid.Direction + "}";
                                            }
                                        }
                                    }
                                }
                            }
                            catch { /* crossing polygon may fail on edge grids */ }

                            // Erase existing hatches inside cell and regenerate
                            try
                            {
                                TypedValue[] hatchFilter = new TypedValue[] {
                                    new TypedValue((int)DxfCode.Start, "HATCH")
                                };
                                PromptSelectionResult psrHatch = ed.SelectCrossingPolygon(cellVerts,
                                    new SelectionFilter(hatchFilter));
                                if (psrHatch.Status == PromptStatus.OK)
                                {
                                    foreach (SelectedObject so in psrHatch.Value)
                                    {
                                        Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForWrite) as Entity;
                                        if (ent != null) ent.Erase();
                                    }
                                }
                            }
                            catch { }

                            // Create new hatch
                            if (currentGrid.TotalIntersects >= 0)
                            {
                                int classIdx = GetSlopeClassIndex(currentGrid.SlopePercent);
                                try
                                {
                                    Hatch hatch = new Hatch();
                                    hatch.SetDatabaseDefaults();
                                    hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                                    hatch.LayerId = slopeClassLayerIds[classIdx];
                                    hatch.Transparency = new Transparency(127);
                                    hatch.Normal = Vector3d.ZAxis;
                                    hatch.Elevation = 0.0;
                                    btr.AppendEntity(hatch);
                                    tr.AddNewlyCreatedDBObject(hatch, true);
                                    hatch.Associative = true;
                                    hatch.AppendLoop(HatchLoopTypes.Default,
                                        new ObjectIdCollection(new ObjectId[] { currentGrid.Id }));
                                    hatch.EvaluateHatch(true);

                                    DrawOrderTable dot = (DrawOrderTable)tr.GetObject(
                                        btr.DrawOrderTableId, OpenMode.ForWrite);
                                    dot.MoveToBottom(new ObjectIdCollection(new ObjectId[] { hatch.ObjectId }));
                                }
                                catch { }
                            }

                            validGrids.Add(currentGrid);
                        }

                        } // End of NOMUTT try-block
                        finally { 
                            try { Application.SetSystemVariable("NOMUTT", 0); } catch { } 
                        }

                        // Export XLSX
                        ObjectId newTableId = tableId;
                        PromptKeywordOptions pkoTable = new PromptKeywordOptions(
                            "\nRegenerate Summary Table? [Y/N] ", "Y N");
                        pkoTable.Keywords.Default = "Y";
                        PromptResult prTable = ed.GetKeywords(pkoTable);
                        if (prTable.Status == PromptStatus.OK && prTable.StringResult == "Y")
                        {
                            Point3d tableInsertPtWCS = Point3d.Origin;
                            bool pointFound = false;

                            if (tableId != ObjectId.Null)
                            {
                                try {
                                    Table oldTable = tr.GetObject(tableId, OpenMode.ForWrite) as Table;
                                    if (oldTable != null && !oldTable.IsErased) {
                                        tableInsertPtWCS = oldTable.Position;
                                        oldTable.Erase();
                                        pointFound = true;
                                    }
                                } catch { }
                            }

                            if (!pointFound)
                            {
                                PromptPointOptions ppo = new PromptPointOptions("\nSelect table insertion point: ");
                                PromptPointResult ppr = ed.GetPoint(ppo);
                                if (ppr.Status == PromptStatus.OK) {
                                    tableInsertPtWCS = ppr.Value.TransformBy(ucs);
                                    pointFound = true;
                                }
                            }

                            if (pointFound) {
                                ObjectId layerTable = GetOrCreateLayer(db, tr, "Grid_Outputs_Table", 7);
                                newTableId = GenerateSummaryTable(tr, btr, db, cs, validGrids, tableInsertPtWCS, layerTable);
                            }
                        }

                        PromptKeywordOptions pkoXlsx = new PromptKeywordOptions(
                            "\nExport data to XLSX? [Y/N] ", "Y N");
                        pkoXlsx.AllowNone = true;
                        pkoXlsx.Keywords.Default = "N";
                        PromptResult prExport = ed.GetKeywords(pkoXlsx);
                        if (prExport.Status == PromptStatus.OK && prExport.StringResult == "Y")
                            ExportToXLSX(doc, validGrids, sideLength, contourInterval);

                        SaveParamsToNOD(db, tr, sideLength, contourInterval, newTableId);
                        tr.Commit();
                        ed.WriteMessage(string.Format("\nUpdate complete! Processed {0} grids.", validGrids.Count));
                        ed.WriteMessage("\ngrid data has correctly updated\n");
                        
                    }
                    finally
                    {
                        if (boundaryRegion != null) boundaryRegion.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(string.Format("\nError: {0}\n{1}", ex.Message, ex.StackTrace));
            }
        }

        #endregion

        #region XLSX Export

        /// <summary>
        /// Exports grid results to an XLSX file with two sheets:
        ///   Sheet 1 "GridData": raw per-cell data
        ///   Sheet 2 "Summary": slope classification + direction tables with live Excel formulas
        /// Uses System.IO.Packaging (WindowsBase.dll) for zero-dependency XLSX generation.
        /// </summary>
        private void ExportToXLSX(Document doc, List<GridData> grids, double sideLength, double contourInterval)
        {
            try
            {
                string dir = Path.GetDirectoryName(doc.Name) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string name = Path.GetFileNameWithoutExtension(doc.Name) ?? "slopedata";
                string baseName = name + "_slopedata";
                string filePath = Path.Combine(dir, baseName + ".xlsx");

                int idx = 1;
                while (File.Exists(filePath))
                {
                    try { 
                        File.Delete(filePath); 
                        break; 
                    }
                    catch {
                        filePath = Path.Combine(dir, string.Format("{0}_{1}.xlsx", baseName, idx));
                        idx++;
                    }
                }




                int lastDataRow = grids.Count + 1; // Row 1 = header, data starts at row 2
                string L = lastDataRow.ToString();

                // Collect shared strings
                List<string> sharedStrings = new List<string>();
                Dictionary<string, int> ssIndex = new Dictionary<string, int>();

                // Pre-register all strings we'll use
                string[] sheet1Headers = { "\u65B9\u683C\u7DE8\u865F", "\u4EA4\u9EDE\u6578 n", "\u5761\u5EA6%", "\u5761\u5EA6\u7D1A\u5225", "\u5761\u5411", "\u9762\u7A4D(m\u00B2)" };
                string[] sheet2T1Headers = { "\u5761\u5EA6\u7D1A\u5225", "\u5761\u5EA6\u7BC4\u570DS(%)", "\u65B9\u683C\u6578", "\u9762\u7A4D(m2)", "\u767E\u5206\u6BD4(%)" };
                string[] sheet2T2Headers = { "\u5761\u5411\u7D1A\u5E8F", "\u5761\u5411\u5225", "\u65B9\u683C\u6578", "\u9762\u7A4D(m2)", "\u767E\u5206\u6BD4(%)" };

                foreach (string h in sheet1Headers) XlsxAddSS(h, sharedStrings, ssIndex);
                foreach (string h in sheet2T1Headers) XlsxAddSS(h, sharedStrings, ssIndex);
                foreach (string h in sheet2T2Headers) XlsxAddSS(h, sharedStrings, ssIndex);
                foreach (string h in SlopeClassNames) XlsxAddSS(h, sharedStrings, ssIndex);
                foreach (string h in SlopeClassRanges) XlsxAddSS(h, sharedStrings, ssIndex);
                foreach (string h in DirLabels) XlsxAddSS(h, sharedStrings, ssIndex);
                XlsxAddSS("\u5408\u3000\u8A08", sharedStrings, ssIndex);
                XlsxAddSS("\u5E73\u5747\u5761\u5EA6(%)", sharedStrings, ssIndex);
                XlsxAddSS("\u5E73\u5747\u5761\u5411", sharedStrings, ssIndex);

                // Grid data strings
                for (int i = 0; i < grids.Count; i++)
                {
                    XlsxAddSS(grids[i].Classification, sharedStrings, ssIndex);
                    XlsxAddSS(grids[i].Direction, sharedStrings, ssIndex);
                }

                // Build XLSX package
                string nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                string nsS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                using (Package pkg = Package.Open(filePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    // Create workbook part
                    Uri wbUri = new Uri("/xl/workbook.xml", UriKind.Relative);
                    PackagePart wbPart = pkg.CreatePart(wbUri,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                        CompressionOption.Normal);
                    pkg.CreateRelationship(wbUri, TargetMode.Internal,
                        nsR + "/officeDocument", "rId1");

                    // Create worksheet parts
                    Uri ws1Uri = new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative);
                    PackagePart ws1Part = pkg.CreatePart(ws1Uri,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                        CompressionOption.Normal);
                    wbPart.CreateRelationship(ws1Uri, TargetMode.Internal,
                        nsR + "/worksheet", "rId1");

                    Uri ws2Uri = new Uri("/xl/worksheets/sheet2.xml", UriKind.Relative);
                    PackagePart ws2Part = pkg.CreatePart(ws2Uri,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                        CompressionOption.Normal);
                    wbPart.CreateRelationship(ws2Uri, TargetMode.Internal,
                        nsR + "/worksheet", "rId2");

                    // Create styles part
                    Uri styUri = new Uri("/xl/styles.xml", UriKind.Relative);
                    PackagePart styPart = pkg.CreatePart(styUri,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                        CompressionOption.Normal);
                    wbPart.CreateRelationship(styUri, TargetMode.Internal,
                        nsR + "/styles", "rId3");

                    // Create shared strings part
                    Uri ssUri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
                    PackagePart ssPart = pkg.CreatePart(ssUri,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                        CompressionOption.Normal);
                    wbPart.CreateRelationship(ssUri, TargetMode.Internal,
                        nsR + "/sharedStrings", "rId4");

                    // ========== Write Workbook ==========
                    using (Stream st = wbPart.GetStream())
                    using (XmlWriter xw = XmlWriter.Create(st, XlsxXws()))
                    {
                        xw.WriteStartDocument();
                        xw.WriteStartElement("workbook", nsS);
                        xw.WriteAttributeString("xmlns", "r", null, nsR);
                        xw.WriteStartElement("sheets", nsS);
                        XlsxWriteSheet(xw, nsS, "GridData", 1, "rId1");
                        XlsxWriteSheet(xw, nsS, "Summary", 2, "rId2");
                        xw.WriteEndElement(); // sheets
                        xw.WriteStartElement("calcPr", nsS);
                        xw.WriteAttributeString("fullCalcOnLoad", "1");
                        xw.WriteEndElement();
                        xw.WriteEndElement(); // workbook
                    }

                    // ========== Write Styles ==========
                    XlsxWriteStyles(styPart, nsS);

                    // ========== Write Shared Strings ==========
                    XlsxWriteSST(ssPart, nsS, sharedStrings);

                    // ========== Write Sheet 1: GridData ==========
                    XlsxWriteGridData(ws1Part, nsS, grids, sheet1Headers, ssIndex);

                    // ========== Write Sheet 2: Summary ==========
                    XlsxWriteSummary(ws2Part, nsS, grids, lastDataRow,
                        sheet2T1Headers, sheet2T2Headers, ssIndex);
                }

                doc.Editor.WriteMessage(string.Format("\n[+] XLSX exported to: {0}", filePath));
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage(string.Format("\n[-] XLSX export failed: {0}", ex.Message));
            }
        }

        // --- XLSX Helper Methods ---

        private static int XlsxAddSS(string s, List<string> list, Dictionary<string, int> idx)
        {
            if (s == null) s = "";
            int i;
            if (idx.TryGetValue(s, out i)) return i;
            i = list.Count;
            list.Add(s);
            idx[s] = i;
            return i;
        }

        private static XmlWriterSettings XlsxXws()
        {
            XmlWriterSettings xs = new XmlWriterSettings();
            xs.Encoding = new UTF8Encoding(false);
            xs.Indent = false;
            return xs;
        }

        private static void XlsxWriteSheet(XmlWriter xw, string ns, string name, int sheetId, string rId)
        {
            xw.WriteStartElement("sheet", ns);
            xw.WriteAttributeString("name", name);
            xw.WriteAttributeString("sheetId", sheetId.ToString());
            xw.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", rId);
            xw.WriteEndElement();
        }

        private static string XlsxColRef(int col)
        {
            string result = "";
            int c = col + 1;
            while (c > 0) { c--; result = (char)('A' + c % 26) + result; c /= 26; }
            return result;
        }

        private static string XlsxCellRef(int row, int col)
        {
            return XlsxColRef(col) + (row + 1).ToString();
        }

        private static void XlsxCellStr(XmlWriter xw, string ns, int row, int col, int ssIdx, int style)
        {
            xw.WriteStartElement("c", ns);
            xw.WriteAttributeString("r", XlsxCellRef(row, col));
            xw.WriteAttributeString("t", "s");
            if (style > 0) xw.WriteAttributeString("s", style.ToString());
            xw.WriteElementString("v", ns, ssIdx.ToString());
            xw.WriteEndElement();
        }

        private static void XlsxCellNum(XmlWriter xw, string ns, int row, int col, double val, int style)
        {
            xw.WriteStartElement("c", ns);
            xw.WriteAttributeString("r", XlsxCellRef(row, col));
            if (style > 0) xw.WriteAttributeString("s", style.ToString());
            xw.WriteElementString("v", ns, val.ToString("G"));
            xw.WriteEndElement();
        }

        private static void XlsxCellFml(XmlWriter xw, string ns, int row, int col, string formula, int style)
        {
            xw.WriteStartElement("c", ns);
            xw.WriteAttributeString("r", XlsxCellRef(row, col));
            if (style > 0) xw.WriteAttributeString("s", style.ToString());
            xw.WriteElementString("f", ns, formula);
            xw.WriteEndElement();
        }

        /// <summary>Writes styles.xml: 2 fonts (default + bold), custom numFmt, 4 cellXfs.</summary>
        private static void XlsxWriteStyles(PackagePart part, string ns)
        {
            using (Stream st = part.GetStream())
            using (XmlWriter xw = XmlWriter.Create(st, XlsxXws()))
            {
                xw.WriteStartDocument();
                xw.WriteStartElement("styleSheet", ns);

                // numFmts
                xw.WriteStartElement("numFmts", ns);
                xw.WriteAttributeString("count", "1");
                xw.WriteStartElement("numFmt", ns);
                xw.WriteAttributeString("numFmtId", "164");
                xw.WriteAttributeString("formatCode", "0.00");
                xw.WriteEndElement();
                xw.WriteEndElement();

                // fonts: 0=default, 1=bold
                xw.WriteStartElement("fonts", ns);
                xw.WriteAttributeString("count", "2");
                xw.WriteStartElement("font", ns);
                xw.WriteStartElement("sz", ns); xw.WriteAttributeString("val", "11"); xw.WriteEndElement();
                xw.WriteStartElement("name", ns); xw.WriteAttributeString("val", "Calibri"); xw.WriteEndElement();
                xw.WriteEndElement();
                xw.WriteStartElement("font", ns);
                xw.WriteStartElement("b", ns); xw.WriteEndElement();
                xw.WriteStartElement("sz", ns); xw.WriteAttributeString("val", "11"); xw.WriteEndElement();
                xw.WriteStartElement("name", ns); xw.WriteAttributeString("val", "Calibri"); xw.WriteEndElement();
                xw.WriteEndElement();
                xw.WriteEndElement();

                // fills (min 2)
                xw.WriteStartElement("fills", ns);
                xw.WriteAttributeString("count", "2");
                xw.WriteStartElement("fill", ns);
                xw.WriteStartElement("patternFill", ns); xw.WriteAttributeString("patternType", "none"); xw.WriteEndElement();
                xw.WriteEndElement();
                xw.WriteStartElement("fill", ns);
                xw.WriteStartElement("patternFill", ns); xw.WriteAttributeString("patternType", "gray125"); xw.WriteEndElement();
                xw.WriteEndElement();
                xw.WriteEndElement();

                // borders (min 1)
                xw.WriteStartElement("borders", ns);
                xw.WriteAttributeString("count", "1");
                xw.WriteStartElement("border", ns);
                xw.WriteStartElement("left", ns); xw.WriteEndElement();
                xw.WriteStartElement("right", ns); xw.WriteEndElement();
                xw.WriteStartElement("top", ns); xw.WriteEndElement();
                xw.WriteStartElement("bottom", ns); xw.WriteEndElement();
                xw.WriteStartElement("diagonal", ns); xw.WriteEndElement();
                xw.WriteEndElement();
                xw.WriteEndElement();

                // cellStyleXfs
                xw.WriteStartElement("cellStyleXfs", ns);
                xw.WriteAttributeString("count", "1");
                xw.WriteStartElement("xf", ns);
                xw.WriteAttributeString("numFmtId", "0"); xw.WriteAttributeString("fontId", "0");
                xw.WriteAttributeString("fillId", "0"); xw.WriteAttributeString("borderId", "0");
                xw.WriteEndElement();
                xw.WriteEndElement();

                // cellXfs: 0=default, 1=bold, 2=number(0.00), 3=bold+number
                xw.WriteStartElement("cellXfs", ns);
                xw.WriteAttributeString("count", "4");
                // 0: default
                xw.WriteStartElement("xf", ns);
                xw.WriteAttributeString("numFmtId", "0"); xw.WriteAttributeString("fontId", "0");
                xw.WriteAttributeString("fillId", "0"); xw.WriteAttributeString("borderId", "0");
                xw.WriteAttributeString("xfId", "0");
                xw.WriteEndElement();
                // 1: bold
                xw.WriteStartElement("xf", ns);
                xw.WriteAttributeString("numFmtId", "0"); xw.WriteAttributeString("fontId", "1");
                xw.WriteAttributeString("fillId", "0"); xw.WriteAttributeString("borderId", "0");
                xw.WriteAttributeString("xfId", "0"); xw.WriteAttributeString("applyFont", "1");
                xw.WriteEndElement();
                // 2: number 0.00
                xw.WriteStartElement("xf", ns);
                xw.WriteAttributeString("numFmtId", "164"); xw.WriteAttributeString("fontId", "0");
                xw.WriteAttributeString("fillId", "0"); xw.WriteAttributeString("borderId", "0");
                xw.WriteAttributeString("xfId", "0"); xw.WriteAttributeString("applyNumberFormat", "1");
                xw.WriteEndElement();
                // 3: bold + number
                xw.WriteStartElement("xf", ns);
                xw.WriteAttributeString("numFmtId", "164"); xw.WriteAttributeString("fontId", "1");
                xw.WriteAttributeString("fillId", "0"); xw.WriteAttributeString("borderId", "0");
                xw.WriteAttributeString("xfId", "0"); xw.WriteAttributeString("applyFont", "1");
                xw.WriteAttributeString("applyNumberFormat", "1");
                xw.WriteEndElement();
                xw.WriteEndElement();

                xw.WriteEndElement(); // styleSheet
            }
        }

        /// <summary>Writes sharedStrings.xml.</summary>
        private static void XlsxWriteSST(PackagePart part, string ns, List<string> strings)
        {
            using (Stream st = part.GetStream())
            using (XmlWriter xw = XmlWriter.Create(st, XlsxXws()))
            {
                xw.WriteStartDocument();
                xw.WriteStartElement("sst", ns);
                xw.WriteAttributeString("count", strings.Count.ToString());
                xw.WriteAttributeString("uniqueCount", strings.Count.ToString());
                foreach (string str in strings)
                {
                    xw.WriteStartElement("si", ns);
                    xw.WriteElementString("t", ns, str ?? "");
                    xw.WriteEndElement();
                }
                xw.WriteEndElement();
            }
        }

        /// <summary>
        /// Writes Sheet 1 (GridData): per-cell data.
        /// Columns: A=GridID, B=Intersects, C=Slope%, D=Class, E=Direction, F=Area
        /// </summary>
        private void XlsxWriteGridData(PackagePart part, string ns, List<GridData> grids,
            string[] headers, Dictionary<string, int> ssIndex)
        {
            using (Stream st = part.GetStream())
            using (XmlWriter xw = XmlWriter.Create(st, XlsxXws()))
            {
                xw.WriteStartDocument();
                xw.WriteStartElement("worksheet", ns);

                xw.WriteStartElement("cols", ns);
                double[] widths = { 12, 12, 12, 14, 10, 14 };
                for (int i = 0; i < 6; i++)
                {
                    xw.WriteStartElement("col", ns);
                    xw.WriteAttributeString("min", (i + 1).ToString());
                    xw.WriteAttributeString("max", (i + 1).ToString());
                    xw.WriteAttributeString("width", widths[i].ToString());
                    xw.WriteAttributeString("customWidth", "1");
                    xw.WriteEndElement();
                }
                xw.WriteEndElement();

                xw.WriteStartElement("sheetData", ns);

                // Row 1 (index 0): Headers (bold style=1)
                xw.WriteStartElement("row", ns);
                xw.WriteAttributeString("r", "1");
                for (int c = 0; c < headers.Length; c++)
                    XlsxCellStr(xw, ns, 0, c, ssIndex[headers[c]], 1);
                xw.WriteEndElement();

                // Data rows
                for (int i = 0; i < grids.Count; i++)
                {
                    int row = i + 1;
                    xw.WriteStartElement("row", ns);
                    xw.WriteAttributeString("r", (row + 1).ToString());

                    XlsxCellNum(xw, ns, row, 0, i + 1, 0);                                   // A: Grid ID
                    XlsxCellNum(xw, ns, row, 1, grids[i].TotalIntersects, 0);                 // B: n
                    XlsxCellNum(xw, ns, row, 2, Math.Round(grids[i].SlopePercent, 2), 2);     // C: Slope%
                    XlsxCellStr(xw, ns, row, 3, ssIndex[grids[i].Classification], 0);         // D: Class
                    XlsxCellStr(xw, ns, row, 4, ssIndex[grids[i].Direction], 0);               // E: Direction
                    XlsxCellNum(xw, ns, row, 5, Math.Round(grids[i].LappingArea, 2), 2);       // F: Area

                    xw.WriteEndElement();
                }

                xw.WriteEndElement(); // sheetData
                xw.WriteEndElement(); // worksheet
            }
        }

        /// <summary>
        /// Writes Sheet 2 (Summary): slope classification + direction distribution with live formulas.
        /// Table 1: Rows 1-10 (slope), Table 2: Rows 13-23 (direction).
        /// </summary>
        private void XlsxWriteSummary(PackagePart part, string ns, List<GridData> grids,
            int lastDataRow, string[] t1Headers, string[] t2Headers,
            Dictionary<string, int> ssIndex)
        {
            string L = lastDataRow.ToString();

            using (Stream st = part.GetStream())
            using (XmlWriter xw = XmlWriter.Create(st, XlsxXws()))
            {
                xw.WriteStartDocument();
                xw.WriteStartElement("worksheet", ns);

                xw.WriteStartElement("cols", ns);
                double[] widths = { 14, 18, 10, 14, 12 };
                for (int i = 0; i < 5; i++)
                {
                    xw.WriteStartElement("col", ns);
                    xw.WriteAttributeString("min", (i + 1).ToString());
                    xw.WriteAttributeString("max", (i + 1).ToString());
                    xw.WriteAttributeString("width", widths[i].ToString());
                    xw.WriteAttributeString("customWidth", "1");
                    xw.WriteEndElement();
                }
                xw.WriteEndElement();

                xw.WriteStartElement("sheetData", ns);

                // ===== TABLE 1: Slope Classification (Rows 1-10) =====
                // Row 1: Headers
                xw.WriteStartElement("row", ns);
                xw.WriteAttributeString("r", "1");
                for (int c = 0; c < t1Headers.Length; c++)
                    XlsxCellStr(xw, ns, 0, c, ssIndex[t1Headers[c]], 1);
                xw.WriteEndElement();

                // Rows 2-8: Slope class data with formulas
                for (int i = 0; i < 7; i++)
                {
                    int row = i + 1;
                    int excelRow = row + 1;
                    xw.WriteStartElement("row", ns);
                    xw.WriteAttributeString("r", excelRow.ToString());

                    XlsxCellStr(xw, ns, row, 0, ssIndex[SlopeClassNames[i]], 0);       // A: Class name
                    XlsxCellStr(xw, ns, row, 1, ssIndex[SlopeClassRanges[i]], 0);       // B: Range

                    // C: =COUNTIF(GridData!D$2:D$L,"className")
                    XlsxCellFml(xw, ns, row, 2,
                        string.Format("COUNTIF(GridData!D$2:D${0},\"{1}\")", L, SlopeClassNames[i]), 0);

                    // D: =ROUND(SUMIF(GridData!D$2:D$L,"className",GridData!F$2:F$L),2)
                    XlsxCellFml(xw, ns, row, 3,
                        string.Format("ROUND(SUMIF(GridData!D$2:D${0},\"{1}\",GridData!F$2:F${0}),2)", L, SlopeClassNames[i]), 2);

                    // E: =IF($D$9>0,ROUND(D{r}/$D$9*100,2),0)
                    XlsxCellFml(xw, ns, row, 4,
                        string.Format("IF($D$9>0,ROUND(D{0}/$D$9*100,2),0)", excelRow), 2);

                    xw.WriteEndElement();
                }

                // Row 9: Totals
                xw.WriteStartElement("row", ns);
                xw.WriteAttributeString("r", "9");
                XlsxCellStr(xw, ns, 8, 0, ssIndex["\u5408\u3000\u8A08"], 1);
                XlsxCellFml(xw, ns, 8, 2, "SUM(C2:C8)", 1);
                XlsxCellFml(xw, ns, 8, 3, "ROUND(SUM(D2:D8),2)", 3);
                XlsxCellNum(xw, ns, 8, 4, 100, 3);
                xw.WriteEndElement();

                // Row 10: Average slope
                xw.WriteStartElement("row", ns);
                xw.WriteAttributeString("r", "10");
                XlsxCellStr(xw, ns, 9, 0, ssIndex["\u5E73\u5747\u5761\u5EA6(%)"], 1);
                XlsxCellFml(xw, ns, 9, 2,
                    string.Format("IF(D9>0,ROUND(SUMPRODUCT(GridData!C$2:C${0},GridData!F$2:F${0})/D9,2),0)", L), 3);
                xw.WriteEndElement();

                // ===== TABLE 2: Direction Distribution (Rows 13-23) =====
                // Row 13: Headers
                xw.WriteStartElement("row", ns);
                xw.WriteAttributeString("r", "13");
                for (int c = 0; c < t2Headers.Length; c++)
                    XlsxCellStr(xw, ns, 12, c, ssIndex[t2Headers[c]], 1);
                xw.WriteEndElement();

                // Rows 14-21: Direction data with formulas
                for (int i = 0; i < 8; i++)
                {
                    int row = 13 + i;
                    int excelRow = row + 1;
                    xw.WriteStartElement("row", ns);
                    xw.WriteAttributeString("r", excelRow.ToString());

                    XlsxCellNum(xw, ns, row, 0, i + 1, 0);                             // A: Seq
                    XlsxCellStr(xw, ns, row, 1, ssIndex[DirLabels[i]], 0);               // B: Direction

                    // C: =COUNTIF(GridData!E$2:E$L,"key")
                    XlsxCellFml(xw, ns, row, 2,
                        string.Format("COUNTIF(GridData!E$2:E${0},\"{1}\")", L, DirKeys[i]), 0);

                    // D: =ROUND(SUMIF(GridData!E$2:E$L,"key",GridData!F$2:F$L),2)
                    XlsxCellFml(xw, ns, row, 3,
                        string.Format("ROUND(SUMIF(GridData!E$2:E${0},\"{1}\",GridData!F$2:F${0}),2)", L, DirKeys[i]), 2);

                    // E: =IF($D$22>0,ROUND(D{r}/$D$22*100,2),0)
                    XlsxCellFml(xw, ns, row, 4,
                        string.Format("IF($D$22>0,ROUND(D{0}/$D$22*100,2),0)", excelRow), 2);

                    xw.WriteEndElement();
                }

                // Row 22: Totals
                xw.WriteStartElement("row", ns);
                xw.WriteAttributeString("r", "22");
                XlsxCellStr(xw, ns, 21, 0, ssIndex["\u5408\u3000\u8A08"], 1);
                XlsxCellFml(xw, ns, 21, 2, "SUM(C14:C21)", 1);
                XlsxCellFml(xw, ns, 21, 3, "ROUND(SUM(D14:D21),2)", 3);
                XlsxCellNum(xw, ns, 21, 4, 100, 3);
                xw.WriteEndElement();

                // Row 23: Dominant direction
                xw.WriteStartElement("row", ns);
                xw.WriteAttributeString("r", "23");
                XlsxCellStr(xw, ns, 22, 0, ssIndex["\u5E73\u5747\u5761\u5411"], 1);
                XlsxCellFml(xw, ns, 22, 3,
                    "IF(D22>0,INDEX(B14:B21,MATCH(MAX(D14:D21),D14:D21,0)),\"\")", 1);
                xw.WriteEndElement();

                xw.WriteEndElement(); // sheetData
                xw.WriteEndElement(); // worksheet
            }
        }

        #endregion
    }
}

















