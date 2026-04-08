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
    /// V3 Grid Slope Calculator Utility
    /// Automates slope percentage calculations based on contour intersections natively across predefined grid cells.
    /// Incorporates Area-Weighted Averaging strictly via Region Geometric Booleans against project boundaries.
    /// </summary>
    public class GridSlopeCalculatorV3
    {
        #region Data Structures

        /// <summary>
        /// Structurally tracks all data evaluated intrinsically for a single localized grid cell physically mapped in CAD.
        /// </summary>
        public class GridData
        {
            /// <summary>The AutoCAD Database Object ID referencing the physical Polyline native boundaries.</summary>
            public ObjectId Id;
            
            /// <summary>The explicitly anchored 3D World Space coordinates.</summary>
            public Point3d CentroidWCS;
            
            /// <summary>The explicitly mapped relative User Vector parameters.</summary>
            public Point3d CentroidUCS;
            
            /// <summary>Top-to-Bottom sorting alignment values.</summary>
            public double SortX;
            public double SortY;
            
            /// <summary>Total valid intersection fragments natively mapping the drawn edges against physical contour gradients.</summary>
            public int TotalIntersects = 0;
            
            /// <summary>The mathematically resolved gradient mapping theoretically calculated exclusively.</summary>
            public double SlopePercent = 0.0;
            
            /// <summary>Standardized category strings parsed explicitly from Slope threshold outputs limits.</summary>
            public string Classification = "";
            
            /// <summary>Target mapped down-hill gradient resolving physical Least-Square Plane mathematical solid vectors.</summary>
            public string Direction = "";
            
            /// <summary>The square parameters mathematically contained validly existing exactly inherently inside the Project Footprint layout.</summary>
            public double LappingArea = 0.0;
        }

        #endregion

        #region Main AutoCAD Command

        /// <summary>
        /// Entry execution pipeline driving the Slope Extractor mapping Engine. 
        /// Trigger natively from AutoCAD CLI: CalcGridSlopeCSV3
        /// </summary>
        [CommandMethod("CalcGridSlopeCSV3")]
        public void CalculateGridSlope()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // ------------- STEP 1: INITIALIZE INPUTS ------------- 
                double sideLength = PromptForDouble(ed, "\nEnter grid side length (L) in meters: ", 25.0);
                if (sideLength <= 0) return;

                double contourInterval = PromptForDouble(ed, "\nEnter contour interval (\u0394h) in meters: ", 1.0);
                if (contourInterval <= 0) return;

                PromptKeywordOptions pko = new PromptKeywordOptions("\nExport data to CSV? [Yes/No] ", "Yes No");
                pko.AllowNone = true;
                pko.Keywords.Default = "No";
                PromptResult prExport = ed.GetKeywords(pko);
                if (prExport.Status != PromptStatus.OK && prExport.Status != PromptStatus.None) return;
                bool shouldExportCsv = (prExport.StringResult == "Yes");

                double intersectTolerance = 0.02;

                // ------------- STEP 2: CAD SELECTION PROCESS ------------- 
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 2a. Prompt user to select grids to operate on
                    PromptSelectionResult psrAllGrids = GetSelectionByLayer(ed, tr, "GRID(s)", "LWPOLYLINE,POLYLINE");
                    if (psrAllGrids == null) return;

                    // 2b. Prompt user to selectively filter exactly which contour boundaries to evaluate
                    PromptSelectionResult psrAllContours = GetSelectionByLayer(ed, tr, "CONTOUR(s)", "LWPOLYLINE,POLYLINE,3DPOLYLINE");
                    if (psrAllContours == null) return;

                    // 2c. Target Boundary Verification securely parsing valid footprint solids mathematically
                    PromptEntityOptions peo = new PromptEntityOptions("\nSelect the Project Boundary (Polyline): ");
                    peo.SetRejectMessage("\nBoundary must be a Polyline!");
                    peo.AddAllowedClass(typeof(Polyline), false);
                    peo.AddAllowedClass(typeof(Polyline2d), false);
                    peo.AddAllowedClass(typeof(Polyline3d), false);
                    PromptEntityResult perBoundary = ed.GetEntity(peo);
                    if (perBoundary.Status != PromptStatus.OK) return;

                    ObjectId[] boundIds = new ObjectId[] { perBoundary.ObjectId };
                    HighlightObjects(tr, boundIds, true);
                    PromptKeywordOptions pkoBound = new PromptKeywordOptions("\nConfirm selected project boundary? [Yes/No] ", "Yes No");
                    pkoBound.Keywords.Default = "Yes";
                    PromptResult prBound = ed.GetKeywords(pkoBound);
                    HighlightObjects(tr, boundIds, false);
                    if (prBound.Status != PromptStatus.OK || prBound.StringResult != "Yes") return;

                    // 2d. Explicit user interface tracking coordinate layouts
                    Matrix3d ucs = ed.CurrentUserCoordinateSystem;
                    CoordinateSystem3d cs = ucs.CoordinateSystem3d;
                    double textRotAngle = Vector3d.XAxis.GetAngleTo(cs.Xaxis, cs.Zaxis);

                    PromptPointOptions ppo = new PromptPointOptions("\nSelect insertion point for the Summary Table: ");
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
                    ObjectId layerArea = GetOrCreateLayer(db, tr, "Grid_Outputs_Area", 3);
                    ObjectId layerTable = GetOrCreateLayer(db, tr, "Grid_Outputs_Table", 7);

                    List<Curve> allContourCurves = GetCurvesFromSelection(tr, psrAllContours);
                    if (allContourCurves.Count == 0)
                    {
                        ed.WriteMessage("\nNo valid contour curves found.");
                        return;
                    }

                    // Translate Boundary to flat Boolean region exclusively blocking bad geometry exceptions
                    Curve baseBoundaryCurve = tr.GetObject(perBoundary.ObjectId, OpenMode.ForRead) as Curve;
                    if (baseBoundaryCurve == null)
                    {
                        ed.WriteMessage("\nFailed to load boundary polyline.");
                        return;
                    }

                    Region boundaryRegion = null;
                    try
                    {
                        using (Curve closedBoundary = baseBoundaryCurve.Clone() as Curve)
                        {
                            if (closedBoundary is Polyline) 
                            {
                                ((Polyline)closedBoundary).Closed = true;
                                ((Polyline)closedBoundary).Elevation = 0.0;
                            }
                            else if (closedBoundary is Polyline2d) 
                            {
                                ((Polyline2d)closedBoundary).Closed = true;
                                ((Polyline2d)closedBoundary).Elevation = 0.0;
                            }
                            else if (closedBoundary is Polyline3d) ((Polyline3d)closedBoundary).Closed = true;

                            DBObjectCollection curveCollection = new DBObjectCollection();
                            curveCollection.Add(closedBoundary);
                            DBObjectCollection regionCollection = Region.CreateFromCurves(curveCollection);
                            if (regionCollection.Count > 0)
                            {
                                boundaryRegion = regionCollection[0] as Region;
                            }
                        }
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception aCADEx)
                    {
                        ed.WriteMessage(string.Format("\nERROR mapping boundary to Region. Self-intersections detected! Please clean boundaries. [{0}]", aCADEx.Message));
                        return;
                    }

                    if (boundaryRegion == null)
                    {
                        ed.WriteMessage("\nERROR: Failed to establish strict contiguous boundary mathematical modeling.");
                        return;
                    }

                    // ------------- STEP 4: MATHEMATICAL ENGINE LOOP ------------- 
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    Plane flatPlane = new Plane(Point3d.Origin, Vector3d.ZAxis);
                    HashSet<string> processedEdges = new HashSet<string>();
                    double textHeight = sideLength * 0.15;

                    List<GridData> preValidGrids = GetValidGrids(tr, psrAllGrids, sideLength, ucs);
                    if (preValidGrids.Count == 0) return;

                    double tolerance = 1.0;
                    preValidGrids = preValidGrids.OrderByDescending(g => Math.Round(g.SortY / tolerance) * tolerance).ThenBy(g => g.SortX).ToList();
                    List<GridData> validGrids = new List<GridData>();

                    for (int i = 0; i < preValidGrids.Count; i++)
                    {
                        GridData currentGrid = preValidGrids[i];
                        Polyline pline = tr.GetObject(currentGrid.Id, OpenMode.ForRead) as Polyline;
                        
                        // Boolean overlapping area explicitly calculated dynamically natively
                        currentGrid.LappingArea = 0.0;
                        DBObjectCollection gridCol = new DBObjectCollection();
                        using (Curve gridClone = pline.Clone() as Curve)
                        {
                            gridCol.Add(gridClone);
                            DBObjectCollection regionCol = Region.CreateFromCurves(gridCol);
                            if (regionCol.Count > 0)
                            {
                                using (Region gridRegion = regionCol[0] as Region)
                                {
                                    try
                                    {
                                        using (Region boundClone = boundaryRegion.Clone() as Region)
                                        {
                                            gridRegion.BooleanOperation(BooleanOperationType.BoolIntersect, boundClone);
                                            currentGrid.LappingArea = gridRegion.Area;
                                        }
                                    }
                                    catch { currentGrid.LappingArea = pline.Area; }
                                }
                                foreach (DBObject obj in regionCol) obj.Dispose();
                            }
                        }

                        // Rejects rendering vectors functionally operating entirely securely outside bounding structures
                        if (currentGrid.LappingArea <= 0.001) continue;

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
                                        segment.IntersectWith(contourCurve, Intersect.OnBothOperands, flatPlane, rawPts, IntPtr.Zero, IntPtr.Zero);
                                        foreach (Point3d pt in rawPts)
                                        {
                                            if (!ContainsPointWithTolerance(uniqueSegmentPts, pt, intersectTolerance))
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
                                Point3d midPt = new Point3d((segment.StartPoint.X + segment.EndPoint.X) / 2, (segment.StartPoint.Y + segment.EndPoint.Y) / 2, 0);
                                
                                string edgeKey = string.Format("{0}_{1}", Math.Round(midPt.X, 3), Math.Round(midPt.Y, 3));

                                if (processedEdges.Add(edgeKey))
                                {
                                    AddTextToBTR(tr, btr, segmentIntersects.ToString(), midPt, textHeight * 0.8, layerEdges, cs, AttachmentPoint.MiddleCenter);
                                }
                            }
                        }

                        List<Point3d> finalGridPts = new List<Point3d>();
                        foreach (Point3d pt in allGridIntersects)
                        {
                            if (!ContainsPointWithTolerance(finalGridPts, pt, intersectTolerance)) finalGridPts.Add(pt);
                        }

                        Vector3d slopeDir;
                        currentGrid.Direction = CalculateAspectAndDirection(finalGridPts, out slopeDir);

                        currentGrid.TotalIntersects = finalGridPts.Count;
                        currentGrid.SlopePercent = ((currentGrid.TotalIntersects * Math.PI * contourInterval) / (8 * sideLength)) * 100;
                        currentGrid.Classification = GetSlopeClassification(currentGrid.SlopePercent);

                        Point3d cUCS = currentGrid.CentroidUCS;
                        Point3d cWCS = currentGrid.CentroidWCS;

                        if (!string.IsNullOrEmpty(currentGrid.Direction))
                        {
                            DrawDirectionArrow(tr, btr, cWCS, slopeDir, sideLength, layerDirection);
                        }

                        // ID mapped using accurate counting purely against *overlapping* grids
                        int displayId = validGrids.Count + 1;
                        Point3d idPt = new Point3d(cUCS.X - (sideLength * 0.45), cUCS.Y + (sideLength * 0.45), 0).TransformBy(ucs);
                        AddTextToBTR(tr, btr, displayId.ToString(), idPt, textHeight, layerId, cs, AttachmentPoint.TopLeft);

                        Point3d areaPt = new Point3d(cUCS.X, cUCS.Y + (sideLength * 0.45), 0).TransformBy(ucs);
                        string areaText = string.Format("A={0:F0} m\u00B2", currentGrid.LappingArea);
                        AddTextToBTR(tr, btr, areaText, areaPt, textHeight * 0.8, layerArea, cs, AttachmentPoint.TopCenter);

                        Point3d nPt = new Point3d(cUCS.X, cUCS.Y - (sideLength * 0.40), 0).TransformBy(ucs);
                        AddTextToBTR(tr, btr, string.Format("n={0}", currentGrid.TotalIntersects), nPt, textHeight, layerTotal, cs, AttachmentPoint.BottomCenter);

                        Point3d slopePt = new Point3d(cUCS.X, cUCS.Y + (textHeight * 0.2), 0).TransformBy(ucs);
                        AddTextToBTR(tr, btr, string.Format("{0:F2}%", currentGrid.SlopePercent), slopePt, textHeight, layerSlope, cs, AttachmentPoint.BottomCenter);

                        Point3d classPt = new Point3d(cUCS.X, cUCS.Y - (textHeight * 0.2), 0).TransformBy(ucs);
                        string classText = string.Format("({0})", currentGrid.Classification);
                        AddTextToBTR(tr, btr, classText, classPt, textHeight * 0.8, layerSlope, cs, AttachmentPoint.TopCenter);

                        if (!string.IsNullOrEmpty(currentGrid.Direction))
                        {
                            Point3d dirTextPt = new Point3d(cUCS.X, cUCS.Y - (sideLength * 0.40), 0).TransformBy(ucs);
                            AddTextToBTR(tr, btr, currentGrid.Direction, dirTextPt, textHeight * 0.8, layerDirText, cs, AttachmentPoint.BottomCenter);
                        }

                        validGrids.Add(currentGrid);
                    }

                    // ------------- STEP 6: OUTPUT GENERATION ------------- 
                    GenerateSummaryTable(tr, btr, db, cs, validGrids, tableInsertPtWCS, layerTable);
                    
                    int mainChunks = (validGrids.Count == 0) ? 1 : ((validGrids.Count - 1) / 20) + 1;
                    double tableWidth = mainChunks * 88.0; 
                    Point3d legendInsertPt = tableInsertPtWCS + (cs.Xaxis * (tableWidth + 30.0));
                    GenerateLegendTable(tr, btr, db, cs, legendInsertPt, layerTable, contourInterval, sideLength);

                    // Compass Legend rendering identically mapping geometric vectors globally offset dynamically 
                    // 2 chunks (width 98) + spacing offset roughly 50 structural limits purely separating legends mathematically.
                    Point3d compassCenterPt = legendInsertPt + (cs.Xaxis * ((2 * 49.0) + 50.0)) - (cs.Yaxis * (sideLength * 1.5));
                    DrawCompassLegend(tr, btr, cs, ucs, compassCenterPt, sideLength, layerEdges, layerDirection, layerDirText);

                    if (shouldExportCsv) ExportToCSV(doc, validGrids);

                    tr.Commit();
                    ed.WriteMessage(string.Format("\nSuccess! Processed {0} grids natively onto screen matrix.", validGrids.Count));
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(string.Format("\nError executing process inherently: {0}", ex.Message));
            }
        }

        #endregion
        
        #region User Interactive Core (CLI)

        /// <summary>
        /// Prompts the user exclusively for Double numeric parameters.
        /// Native blocking checks natively prevent 0 or strictly negative constraints mathematically natively tracking default parameters securely.
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
        /// Explicit interactive method natively targeting Sample items to automatically resolve target layers inherently natively blocking manual typing workflows.
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

                PromptKeywordOptions pko = new PromptKeywordOptions(string.Format("\nFound {0} items on layer(s) [{1}]. Is this correct? [Yes/No] ", ids.Length, layers), "Yes No");
                pko.Keywords.Default = "Yes"; 
                
                PromptResult pr = ed.GetKeywords(pko);
                HighlightObjects(tr, ids, false);
                
                if (pr.Status != PromptStatus.OK) return null;
                if (pr.StringResult == "Yes") return psrAll;
            }
        }

        /// <summary>
        /// Translates visual Selection Result vectors actively back into explicitly handled Curve Solid parameters mapped explicitly out of CAD DB Arrays.
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
        /// Rigid Area calculation engine iterating physically bound grid slices and validating dimensional thresholds ensuring 
        /// isolated segments perfectly align accurately against theoretical target bounding parameters natively mathematically evaluating layout arrays.
        /// </summary>
        private List<GridData> GetValidGrids(Transaction tr, PromptSelectionResult psr, double sideLength, Matrix3d ucs)
        {
            List<GridData> valid = new List<GridData>();
            double targetArea = sideLength * sideLength;

            foreach (SelectedObject so in psr.Value)
            {
                Curve curve = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Curve;
                if (curve != null && curve.Closed)
                {
                    if (Math.Abs(curve.Area - targetArea) < (targetArea * 0.05))
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
        /// Safely evaluates mapping coordinates resolving tight geometry intersections preventing artificial doubling vectors.
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
        /// Extreme 3D Mathematics solving planar vectors algorithmically identifying purely theoretically calculated downhill orientation strings.
        /// Translates point elevations into a planar Least Square Model calculating the steepest slope bearing.
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
            if (Math.Abs(D) < 1e-8) return ""; 

            double a = (sxz * syy - syz * sxy) / D;
            double b = (syz * sxx - sxz * sxy) / D;

            if (Math.Abs(a) < 1e-8 && Math.Abs(b) < 1e-8) return ""; 

            Vector3d rawDir = new Vector3d(-a, -b, 0).GetNormal();
            double angle = Math.Atan2(rawDir.Y, rawDir.X);
            double deg = angle * 180.0 / Math.PI;

            if (deg < 0) deg += 360.0;

            if (deg >= 337.5 || deg < 22.5) { direction = new Vector3d(1, 0, 0); return "E"; }
            if (deg >= 22.5 && deg < 67.5) { direction = new Vector3d(1, 1, 0).GetNormal(); return "NE"; }
            if (deg >= 67.5 && deg < 112.5) { direction = new Vector3d(0, 1, 0); return "N"; }
            if (deg >= 112.5 && deg < 157.5) { direction = new Vector3d(-1, 1, 0).GetNormal(); return "NW"; }
            if (deg >= 157.5 && deg < 202.5) { direction = new Vector3d(-1, 0, 0); return "W"; }
            if (deg >= 202.5 && deg < 247.5) { direction = new Vector3d(-1, -1, 0).GetNormal(); return "SW"; }
            if (deg >= 247.5 && deg < 292.5) { direction = new Vector3d(0, -1, 0); return "S"; }
            if (deg >= 292.5 && deg < 337.5) { direction = new Vector3d(1, -1, 0).GetNormal(); return "SE"; }

            return "";
        }

        #endregion

        #region AutoCAD Drawing Handlers

        /// <summary>
        /// Natively constructs an absolute visual 3x3 array identically echoing physical direction shapes ensuring output variables identically conform structurally across varying grid lengths!
        /// </summary>
        private void DrawCompassLegend(Transaction tr, BlockTableRecord btr, CoordinateSystem3d cs, Matrix3d ucs, Point3d centerPtWCS, double sideLength, ObjectId layerEdges, ObjectId layerDirection, ObjectId layerDirText)
        {
            double textRotAngle = Vector3d.XAxis.GetAngleTo(cs.Xaxis, cs.Zaxis);
            double textHeight = sideLength * 0.15;

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
                        AddTextToBTR(tr, btr, cardinalNames[arrY, arrX], textPt, textHeight * 0.8, layerDirText, cs, AttachmentPoint.BottomCenter);
                    }
                    else
                    {
                        Point3d centerLabelPt = new Point3d(cUCS.X, cUCS.Y, 0).TransformBy(ucs);
                        AddTextToBTR(tr, btr, cardinalNames[arrY, arrX], centerLabelPt, textHeight * 0.8, layerDirText, cs, AttachmentPoint.MiddleCenter);
                    }
                }
            }
        }

        /// <summary>
        /// Explicit geometric drawing algorithm plotting Directional aspects as physically shaped variable-width Polyline arrows mapped over CAD Modelspace.
        /// </summary>
        private void DrawDirectionArrow(Transaction tr, BlockTableRecord btr, Point3d centroid, Vector3d dir, double cellSide, ObjectId layerId)
        {
            if (Math.Abs(dir.X) < 1e-8 && Math.Abs(dir.Y) < 1e-8) return;

            double length = cellSide * 0.4;
            Point3d start = centroid - dir * (length * 0.5);
            Point3d mid = centroid + dir * (length * 0.2);
            Point3d end = centroid + dir * (length * 0.5);

            double shaftWidth = cellSide * 0.015;
            double headWidth = cellSide * 0.1;

            Polyline pl = new Polyline();
            pl.LayerId = layerId;
            
            pl.AddVertexAt(0, new Point2d(start.X, start.Y), 0, shaftWidth, shaftWidth);
            pl.AddVertexAt(1, new Point2d(mid.X, mid.Y), 0, headWidth, 0);
            pl.AddVertexAt(2, new Point2d(end.X, end.Y), 0, 0, 0);
            
            pl.Elevation = 0;

            btr.AppendEntity(pl);
            tr.AddNewlyCreatedDBObject(pl, true);
        }

        /// <summary>
        /// Highlight verification mapping engine explicitly identifying object selections for UI Prompts avoiding false positives.
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
        /// Translates Selection parameters to explicitly trace their mapped Layer names grouping similar drawing metrics.
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
        /// Automatically maps requested output layers natively into the drawing Database explicitly defining visual standard Color formatting constraints.
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
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, colorIndex);
                ObjectId newLayerId = lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
                return newLayerId;
            }
        }

        /// <summary>
        /// Core rendering function injecting literal metric parameters natively parsed physically scaling Text strings directly into drawing matrices.
        /// Upgraded to MText ensuring TrueType Unicode handling preventing legacy SHX formatting breaking CJK Chinese translations globally across native UI states.
        /// Exact CoordinateSystem mappings permanently lock angles inherently preventing text rotation bounding box shifts ensuring flawless visual matching!
        /// </summary>
        private void AddTextToBTR(Transaction tr, BlockTableRecord btr, string text, Point3d pos, double height, ObjectId layerId, CoordinateSystem3d cs, AttachmentPoint align)
        {
            MText mText = new MText();
            // Force pure TrueType Arial formatting natively mapping Unicode & Chinese strings identically avoiding basic font errors!
            mText.Contents = "{\\fArial|b0|i0|c0|p0;" + text + "}"; 
            mText.TextHeight = height;
            mText.LayerId = layerId;
            try { mText.Attachment = align; } 
            catch { mText.Attachment = AttachmentPoint.BottomLeft; } // Set Attachment strictly prior to position locks matrix geometries securely
            
            mText.Normal = cs.Zaxis;
            mText.Direction = cs.Xaxis;
            mText.Location = pos;
            
            btr.AppendEntity(mText);
            tr.AddNewlyCreatedDBObject(mText, true);
        }

        #endregion

        #region Data Table & CSV Export

        /// <summary>
        /// Core categorizer actively sorting Slope Percentages matching explicit regional bounds arrays exclusively translating physical gradients into legal bounds limits.
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
        /// Assembles an explicit Legend layout structurally scaling mathematically defined iterations explicitly from explicitly 0 to explicitly 100.
        /// Modulates natively parsing exactly Even coordinates actively cleanly clustering layouts natively.
        /// </summary>
        private void GenerateLegendTable(Transaction tr, BlockTableRecord btr, Database db, CoordinateSystem3d cs, Point3d insertPt, ObjectId layerTable, double contourInterval, double sideLength)
        {
            int maxDataRows = 26; 
            int totalItems = 51; // Scale bounds for explicitly even ranges natively n=0, 2, 4... 100
            int chunks = ((totalItems - 1) / maxDataRows) + 1; 
            int totalCols = chunks * 3;
            int totalRows = maxDataRows + 2; 

            Table tb = new Table();
            tb.TableStyle = db.Tablestyle;
            tb.LayerId = layerTable;
            
            tb.Position = insertPt;
            tb.Normal = cs.Zaxis;
            tb.Direction = cs.Xaxis;

            tb.SetSize(totalRows, totalCols);
            
            // Loop strictly across explicit structure cells avoiding globally deprecated legacy flags
            for(int r = 0; r < totalRows; r++)
            {
                tb.Rows[r].Height = 4.0;
                for(int c = 0; c < totalCols; c++)
                {
                    tb.Cells[r, c].TextHeight = 1.8;
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
                int chunkIndex = index / maxDataRows;
                int localRow = (index % maxDataRows) + 2;
                int baseCol = chunkIndex * 3;

                // Native calculation simulating formula engine parameters cleanly identically natively outside of physical bounds checks
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
        /// Explicit summary report matrix compiler looping inherently structurally compiling Area parameters formatting raw values natively across CAD spaces.
        /// </summary>
        private void GenerateSummaryTable(Transaction tr, BlockTableRecord btr, Database db, CoordinateSystem3d cs, List<GridData> validGrids, Point3d insertPt, ObjectId layerTable)
        {
            int maxDataRows = 20; 
            int chunks = ((validGrids.Count - 1) / maxDataRows) + 1;
            int totalCols = chunks * 6; // Expanded to 6 columns for Area mapping
            int totalRows = Math.Min(validGrids.Count, maxDataRows) + 3; // +2 for default headers +1 for master summary row

            Table tb = new Table();
            tb.TableStyle = db.Tablestyle;
            tb.LayerId = layerTable;
            
            tb.Position = insertPt;
            tb.Normal = cs.Zaxis;
            tb.Direction = cs.Xaxis;

            tb.SetSize(totalRows, totalCols);
            
            // Loop through the specific raw cells to force apply format overrides instead of using deprecated global flags
            for(int r = 0; r < totalRows; r++)
            {
                tb.Rows[r].Height = 4.0;
                for(int c = 0; c < totalCols; c++)
                {
                    tb.Cells[r, c].TextHeight = 1.8;
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
                tb.Cells[1, baseCol + 5].TextString = "Area (m\u00B2)"; // properly decoded square symbol
            }

            for (int i = 0; i < validGrids.Count; i++)
            {
                int chunkIndex = i / maxDataRows;
                int localRow = (i % maxDataRows) + 2;
                int baseCol = chunkIndex * 6;

                tb.Cells[localRow, baseCol].TextString = (i + 1).ToString();
                tb.Cells[localRow, baseCol + 1].TextString = validGrids[i].TotalIntersects.ToString();
                tb.Cells[localRow, baseCol + 2].TextString = string.Format("{0:F2}%", validGrids[i].SlopePercent);
                tb.Cells[localRow, baseCol + 3].TextString = validGrids[i].Classification;
                tb.Cells[localRow, baseCol + 4].TextString = validGrids[i].Direction;
                tb.Cells[localRow, baseCol + 5].TextString = string.Format("{0:F2}", validGrids[i].LappingArea);
            }

            double meanSlope = 0.0;
            string modeDirection = "";

            if (validGrids.Count > 0)
            {
                double totalWeights = validGrids.Sum(g => g.LappingArea);
                if (totalWeights > 0.0001)
                {
                    meanSlope = validGrids.Sum(g => g.SlopePercent * g.LappingArea) / totalWeights;
                }
                else
                {
                    meanSlope = validGrids.Average(g => g.SlopePercent); // Fallback
                }
                
                var dirGroups = validGrids.Where(g => !string.IsNullOrEmpty(g.Direction))
                                          .GroupBy(g => g.Direction)
                                          .OrderByDescending(g => g.Count());
                
                if (dirGroups.Any()) modeDirection = dirGroups.First().Key;
            }

            int summaryRow = totalRows - 1;
            tb.MergeCells(CellRange.Create(tb, summaryRow, 0, summaryRow, totalCols - 1));
            tb.Cells[summaryRow, 0].TextString = string.Format("Area-Weighted Average Slope: {0:F2}%   |   Primary Direction: {1}", meanSlope, modeDirection);
            
            tb.GenerateLayout();
            btr.AppendEntity(tb);
            tr.AddNewlyCreatedDBObject(tb, true);
        }

        /// <summary>
        /// Writes native variables streaming explicitly formatted structurally into pure UTF-8 BOM encoding ensuring exact mapping parsing native Chinese Classifications string formatting into external Excel arrays explicitly safely.
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
                    filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "slopedata.csv");
                }

                double meanSlope = 0.0;
                string modeDir = "";

                if (grids.Count > 0)
                {
                    double totalWeights = grids.Sum(g => g.LappingArea);
                    if (totalWeights > 0.0001)
                    {
                        meanSlope = grids.Sum(g => g.SlopePercent * g.LappingArea) / totalWeights;
                    }
                    else
                    {
                        meanSlope = grids.Average(g => g.SlopePercent); // Fallback
                    }
                    
                    var dirGroups = grids.Where(g => !string.IsNullOrEmpty(g.Direction))
                                         .GroupBy(g => g.Direction)
                                         .OrderByDescending(g => g.Count());
                    if (dirGroups.Any()) modeDir = dirGroups.First().Key;
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Grid ID,Intersects (n),Slope (%),Classification,Direction,Area (sq m)");
                for (int i = 0; i < grids.Count; i++)
                {
                    GridData g = grids[i];
                    sb.AppendLine(string.Format("{0},{1},{2:F2},{3},{4},{5:F2}", i + 1, g.TotalIntersects, g.SlopePercent, g.Classification, g.Direction, g.LappingArea));
                }
                
                sb.AppendLine();
                sb.AppendLine(string.Format("Area-Weighted Average Slope (%),{0:F2}", meanSlope));
                sb.AppendLine(string.Format("Primary Direction,{0}", modeDir));

                File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(true));
                doc.Editor.WriteMessage(string.Format("\n[+] Successfully exported data matrix CSV to: {0}", filePath));
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage(string.Format("\n[-] Failed to generate mapping CSV file: {0}", ex.Message));
            }
        }

        #endregion

    }
}
