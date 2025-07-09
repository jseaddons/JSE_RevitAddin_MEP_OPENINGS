using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class PipeSleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DebugLogger.Log("Starting PipeSleeveCommand");
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Select the pipe sleeve symbol (PS# in OpeningOnWall family)
            var pipeSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family != null
                                       && sym.Family.Name.IndexOf("OpeningOnWall", System.StringComparison.OrdinalIgnoreCase) >= 0
                                       && sym.Name.Replace(" ", "").StartsWith("PS#", System.StringComparison.OrdinalIgnoreCase));

            // Find cluster symbols for suppression (PipeOpeningOnWall and PipeOpeningOnSlab)
            var clusterSymbolWall = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family != null && sym.Family.Name.Equals("PipeOpeningOnWall", StringComparison.OrdinalIgnoreCase));
            var clusterSymbolSlab = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family != null && sym.Family.Name.Equals("PipeOpeningOnSlab", StringComparison.OrdinalIgnoreCase));
            if (pipeSymbol == null)
            {
                TaskDialog.Show("Error", "No pipe opening family symbol (PS#) found.");
                return Result.Failed;
            }

            using (var txActivate = new Transaction(doc, "Activate Pipe Symbol"))
            {
                txActivate.Start();
                if (!pipeSymbol.IsActive)
                    pipeSymbol.Activate();
                txActivate.Commit();
            }

            // Find a non-template 3D view
            var view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate);
            if (view3D == null)
            {
                TaskDialog.Show("Error", "No non-template 3D view found.");
                return Result.Failed;
            }

            // Wall filter and intersector (face-based)
            ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            // Use face-based intersection to get actual wall face reference, including linked documents
            var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Face, view3D)
            {
                FindReferencesInRevitLinks = true
            };

            // Sleeve placer
            var placer = new PipeSleevePlacer(doc);

            // Configurable clearance (10mm in feet)
            double clearance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters);

            // Collect all pipes
            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .WhereElementIsNotElementType()
                .Cast<Pipe>()
                // Filter out inclined and vertical pipes
                .Where(p =>
                {
                    var locCurve = p.Location as LocationCurve;
                    if (locCurve?.Curve is Line line)
                    {
                        var dir = line.Direction.Normalize();
                        // Only allow pipes that are (almost) horizontal (Z direction close to 0)
                        return Math.Abs(dir.Z) < 1e-3;
                    }
                    return false;
                })
                .ToList();

            // Special debug logging for target pipe 1891461
            var targetPipe = pipes.FirstOrDefault(p => p.Id.IntegerValue == 1891461);
            if (targetPipe != null)
            {
                DebugLogger.Log($"*** TARGET PIPE 1891461 FOUND IN COLLECTION! ***");
                var targetLocation = targetPipe.Location as LocationCurve;
                if (targetLocation?.Curve is Line targetLine)
                {
                    DebugLogger.Log($"*** TARGET PIPE 1891461: Location = {targetLine.GetEndPoint(0)} to {targetLine.GetEndPoint(1)}, Direction = {targetLine.Direction}, Length = {targetLine.Length} ***");
                }
                else
                {
                    DebugLogger.Log($"*** TARGET PIPE 1891461: FOUND but NO VALID LOCATION CURVE ***");
                }
            }
            else
            {
                DebugLogger.Log($"*** TARGET PIPE 1891461: NOT FOUND in initial collection - searching all elements ***");
                // Search for the specific element by ID to see if it exists at all
                try
                {
                    var elementById = doc.GetElement(new ElementId(1891461));
                    if (elementById != null)
                    {
                        DebugLogger.Log($"*** TARGET PIPE 1891461: Element found by ID as {elementById.GetType().Name}, Category = {elementById.Category?.Name ?? "None"} ***");
                    }
                    else
                    {
                        DebugLogger.Log($"*** TARGET PIPE 1891461: Element ID not found in document ***");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"*** TARGET PIPE 1891461: Error getting element by ID: {ex.Message} ***");
                }
            }

            // SUPPRESSION: Collect existing pipe sleeves for reference (duplicate detection is handled by PipeSleevePlacer)
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("PS#"))
                .ToList();
            
            DebugLogger.Log($"Found {existingSleeves.Count} existing pipe sleeves in the model");

            // Initialize counters for detailed logging
            int totalPipes = pipes.Count();
            int intersectionCount = 0;
            int placedCount = 0;
            int missingCount = 0;
            int skippedExistingCount = 0; // Counter for pipes with existing sleeves
            HashSet<ElementId> processedPipes = new HashSet<ElementId>(); // Track processed pipes to prevent duplicates
            
            DebugLogger.Log($"Found {totalPipes} pipes to process");
            DebugLogger.Log($"Target pipe 1891461 in collection: {pipes.Any(p => p.Id.IntegerValue == 1891461)}");

            // Start transaction for placement
            using (var tx = new Transaction(doc, "Place Pipe Sleeves"))
            {
                tx.Start();
                foreach (var pipe in pipes)
                {
                    // --- CLUSTERING CHECK FOR SOIL/WASTE SYSTEMS ---
                    var system = pipe.MEPSystem;
                    if (system != null && !string.IsNullOrEmpty(system.Name))
                    {
                        string sysName = system.Name.ToLowerInvariant();
                        if (sysName.Contains("soil") || sysName.Contains("waste"))
                        {
                            // Check for clustering of fittings near this pipe
                            // Define a small radius for clustering (e.g., 300mm)
                            double clusterRadius = UnitUtils.ConvertToInternalUnits(300.0, UnitTypeId.Millimeters);
                            var pipeCurve = (pipe.Location as LocationCurve)?.Curve as Line;
                            if (pipeCurve != null)
                            {
                                // Collect all fittings in the system
                                var fittings = new FilteredElementCollector(doc)
                                    .OfClass(typeof(FamilyInstance))
                                    .WhereElementIsNotElementType()
                                    .Cast<FamilyInstance>()
                                    .Where(fi =>
                                        fi.MEPModel != null &&
                                        fi.MEPModel.ConnectorManager != null &&
                                        fi.MEPModel.ConnectorManager.Connectors != null &&
                                        fi.MEPModel.ConnectorManager.Connectors.Cast<Connector>().Any(c => c.MEPSystem != null && c.MEPSystem.Id == system.Id)
                                    )
                                    .ToList();

                                // Check if there are 2 or more fittings within the cluster radius of the pipe's midpoint
                                XYZ midPt = pipeCurve.Evaluate(0.5, true);
                                int clusterCount = fittings.Count(fi =>
                                {
                                    var fittingLoc = fi.Location as LocationPoint;
                                    if (fittingLoc != null)
                                        return fittingLoc.Point.DistanceTo(midPt) < clusterRadius;
                                    // Fallback: check connectors
                                    var connectors = fi.MEPModel?.ConnectorManager?.Connectors;
                                    if (connectors != null)
                                    {
                                        foreach (Connector c in connectors)
                                        {
                                            if (c.Origin.DistanceTo(midPt) < clusterRadius)
                                                return true;
                                        }
                                    }
                                    return false;
                                });

                                if (clusterCount >= 2)
                                {
                                    DebugLogger.Log($"Pipe ID={pipe.Id.IntegerValue}: Skipped sleeve placement due to fitting clustering for SOIL/WASTE system.");
                                    continue; // Skip sleeve placement for this pipe
                                }
                            }
                        }
                    }
                    // --- END CLUSTERING CHECK ---

                    bool isTargetPipe = pipe.Id.IntegerValue == 1891461;
                    if (isTargetPipe)
                    {
                        DebugLogger.Log($"*** DETAILED DEBUG FOR PIPE 1891461 ***");
                    }
                    DebugLogger.Log($"Processing Pipe ID={pipe.Id.IntegerValue}");
                    // Prevent processing the same pipe twice (avoid duplicate sleeves)
                    if (processedPipes.Contains(pipe.Id))
                    {
                        DebugLogger.Log($"Pipe ID={pipe.Id.IntegerValue}: already processed, skipping");
                        continue;
                    }
                    var locCurve = (pipe.Location as LocationCurve)?.Curve as Line;
                    if (locCurve == null)
                    {
                        DebugLogger.Log($"Pipe ID={pipe.Id.IntegerValue}: no curve, skipping");
                        if (isTargetPipe)
                        {
                            DebugLogger.Log($"*** TARGET PIPE 1891461: EXCLUDED - no location curve ***");
                        }
                        missingCount++;
                        continue;
                    }
                    // Determine clearance: 25mm per side for insulated, 50mm per side for non-insulated
                    double clearancePerSideMM = 50.0; // default: 50mm per side
                    var insulationThicknessParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS);
                    if (insulationThicknessParam != null && insulationThicknessParam.HasValue && insulationThicknessParam.AsDouble() > 0)
                    {
                        clearancePerSideMM = 25.0; // 25mm per side for insulated
                        DebugLogger.Log($"Pipe ID={pipe.Id.IntegerValue}: Insulated pipe detected, using 25mm per side clearance");
                    }
                    else
                    {
                        DebugLogger.Log($"Pipe ID={pipe.Id.IntegerValue}: Non-insulated pipe, using 50mm per side clearance");
                    }
                    double totalClearance = UnitUtils.ConvertToInternalUnits(clearancePerSideMM * 2.0, UnitTypeId.Millimeters); // total clearance to add to diameter
                    // NOTE: Suppression check moved to actual placement points to prevent duplicates
                    var line = locCurve;

                    // Enhanced wall intersection: use robust intersection logic
                    XYZ rayDirection = line.Direction;
                    XYZ perpDirection = new XYZ(-rayDirection.Y, rayDirection.X, 0).Normalize();
                    IList<ReferenceWithContext> refWithContext;
                    bool isXOrientation = Math.Abs(rayDirection.X) > Math.Abs(rayDirection.Y);

                    if (isXOrientation)
                    {
                        DebugLogger.Log($"[PipeSleeveCommand] Entering X-direction logic for pipe {pipe.Id.IntegerValue}");
                        DebugLogger.Log($"[PipeSleeveCommand] Pipe direction: {rayDirection}, Start: {line.GetEndPoint(0)}, End: {line.GetEndPoint(1)}");
                        
                        var sampleFractions = new[] { 0.0, 0.25, 0.5, 0.75, 1.0 };
                        var allHits = new List<ReferenceWithContext>();
                        
                        // Cast rays in both positive and negative directions from each sample point
                        // This ensures we catch walls regardless of pipe direction
                        foreach (double t in sampleFractions)
                        {
                            var samplePt = line.Evaluate(t, true);
                            DebugLogger.Log($"[PipeSleeveCommand] Sampling at t={t}, point={samplePt}");
                            
                            if (isTargetPipe)
                            {
                                DebugLogger.Log($"*** TARGET PIPE 1891461: Casting rays from sample point {samplePt} ***");
                            }
                            
                            // Cast in pipe direction and opposite direction
                            var hitsFwd = refIntersector.Find(samplePt, rayDirection)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            var hitsBack = refIntersector.Find(samplePt, rayDirection.Negate())?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            
                            // Also cast in perpendicular directions to catch walls parallel to pipe
                            var perpDir1 = new XYZ(-rayDirection.Y, rayDirection.X, 0).Normalize();
                            var perpDir2 = perpDir1.Negate();
                            var hitsPerp1 = refIntersector.Find(samplePt, perpDir1)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            var hitsPerp2 = refIntersector.Find(samplePt, perpDir2)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            
                            DebugLogger.Log($"[PipeSleeveCommand] Sample {t}: hitsFwd={hitsFwd.Count}, hitsBack={hitsBack.Count}, hitsPerp1={hitsPerp1.Count}, hitsPerp2={hitsPerp2.Count}");
                            
                            if (isTargetPipe)
                            {
                                DebugLogger.Log($"*** TARGET PIPE 1891461: Ray results - hitsFwd={hitsFwd.Count}, hitsBack={hitsBack.Count}, hitsPerp1={hitsPerp1.Count}, hitsPerp2={hitsPerp2.Count} ***");
                            }
                            
                            allHits.AddRange(hitsFwd);
                            allHits.AddRange(hitsBack);
                            allHits.AddRange(hitsPerp1);
                            allHits.AddRange(hitsPerp2);
                        }
                        
                        var grouped = allHits.GroupBy(h => {
                            var r = h.GetReference();
                            var linkInst = doc.GetElement(r.ElementId) as RevitLinkInstance;
                            return linkInst != null ? r.LinkedElementId : r.ElementId;
                        });
                        
                        bool placed = false;
                        double pipeLength = line.Length;
                        
                        if (isTargetPipe)
                        {
                            DebugLogger.Log($"*** TARGET PIPE 1891461: Total hits found = {allHits.Count}, Grouped into {grouped.Count()} wall groups ***");
                        }
                        
                        // --- BEGIN ROBUST SOLID INTERSECTION LOGIC FOR X-PIPES ---
                        DebugLogger.Log($"[PipeSleeveCommand] USING DUCT LOGIC FOR PIPE INTERSECTION (X-direction)");
                        var bestWall = (Wall)null;
                        XYZ bestEntry = null, bestExit = null;
                        XYZ bestFaceNormal = null;
                        double bestWallThickness = 0.0;
                        double maxSegment = 0.0;
                        var pipeLine = locCurve;
                        // Validate pipe line length to prevent short curve errors
                        if (pipeLength < 0.01) // Less than 0.01 feet (about 3mm)
                        {
                            DebugLogger.Log($"[PipeSleeveCommand] Skipping very short X-pipe {pipe.Id.IntegerValue} with length {pipeLength:F6}");
                            continue;
                        }
                        foreach (var group in grouped)
                        {
                            var hits = group.OrderBy(h => h.Proximity).ToList();
                            var rEntry = hits.First().GetReference();
                            var linkInstEntryExit = doc.GetElement(rEntry.ElementId) as RevitLinkInstance;
                            var targetDocEntryExit = linkInstEntryExit != null ? linkInstEntryExit.GetLinkDocument() : doc;
                            var eidEntryExit = linkInstEntryExit != null ? rEntry.LinkedElementId : rEntry.ElementId;
                            var hostWallEntryExit = targetDocEntryExit?.GetElement(eidEntryExit) as Wall;
                            if (hostWallEntryExit == null) {
                                DebugLogger.Log($"[PipeSleeveCommand] X-pipe: Wall not found for hit group, skipping");
                                continue;
                            }
                            double wallThickness = hostWallEntryExit.Width;
                            XYZ wallOrient = hostWallEntryExit.Orientation.Normalize();
                            Solid wallSolid = null;
                            try {
                                Options geomOptions = new Options { ComputeReferences = true, IncludeNonVisibleObjects = false };
                                GeometryElement geomElem = hostWallEntryExit.get_Geometry(geomOptions);
                                foreach (GeometryObject obj in geomElem) {
                                    if (obj is Solid solid && solid.Volume > 0) {
                                        wallSolid = solid;
                                        break;
                                    }
                                }
                            } catch { wallSolid = null; }
                            if (wallSolid != null) {
                                // --- Robust intersection logic (like Y-pipe) ---
                                List<XYZ> intersectionPoints = new List<XYZ>();
                                foreach (Face face in wallSolid.Faces) {
                                    IntersectionResultArray ira = null;
                                    SetComparisonResult res = face.Intersect(line, out ira);
                                    if (res == SetComparisonResult.Overlap && ira != null) {
                                        foreach (IntersectionResult ir in ira) {
                                            var intersectionPoint = GetIntersectionPoint(ir);
                                            if (intersectionPoint != null)
                                            {
                                                intersectionPoints.Add(intersectionPoint);
                                            }
                                        }
                                    }
                                }
                                bool startInside = IsPointInsideSolid(wallSolid, line.GetEndPoint(0), hostWallEntryExit.Orientation);
                                bool endInside = IsPointInsideSolid(wallSolid, line.GetEndPoint(1), hostWallEntryExit.Orientation);
                                // Fallback: segment sampling if no intersections
                                if (intersectionPoints.Count == 0 && !startInside && !endInside) {
                                    List<XYZ> altIntersections = new List<XYZ>();
                                    int segments = 10;
                                    for (int i = 0; i < segments; i++) {
                                        double t1 = (double)i / segments;
                                        double t2 = (double)(i + 1) / segments;
                                        XYZ pt1 = line.Evaluate(t1, true);
                                        XYZ pt2 = line.Evaluate(t2, true);
                                        double segmentLength = pt1.DistanceTo(pt2);
                                        if (segmentLength < 0.01) continue;
                                        Line segment;
                                        try { segment = Line.CreateBound(pt1, pt2); } catch { continue; }
                                        foreach (Face face in wallSolid.Faces) {
                                            IntersectionResultArray ira2 = null;
                                            SetComparisonResult res2 = face.Intersect(segment, out ira2);
                                            if (res2 == SetComparisonResult.Overlap && ira2 != null) {
                                                foreach (IntersectionResult ir in ira2) {
                                                    var intersectionPoint = GetIntersectionPoint(ir);
                                                    if (intersectionPoint != null && !altIntersections.Any(pt => pt.DistanceTo(intersectionPoint) < UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters))) {
                                                        altIntersections.Add(intersectionPoint);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (altIntersections.Count > 0) {
                                        intersectionPoints = altIntersections;
                                    }
                                }
                                if (intersectionPoints.Count == 0 && !startInside && !endInside) {
                                    DebugLogger.Log($"[PipeSleeveCommand] X-pipe {pipe.Id.IntegerValue}: No intersection found with wall {hostWallEntryExit.Id.IntegerValue}, skipping.");
                                    continue;
                                }
                                XYZ ptEntry = null, ptExit = null;
                                if (intersectionPoints.Count >= 2) {
                                    intersectionPoints = intersectionPoints.OrderBy(pt => (pt - line.GetEndPoint(0)).GetLength()).ToList();
                                    ptEntry = intersectionPoints.First();
                                    ptExit = intersectionPoints.Last();
                                } else if (intersectionPoints.Count == 1) {
                                    ptEntry = intersectionPoints[0];
                                    ptExit = startInside ? line.GetEndPoint(0) : line.GetEndPoint(1);
                                } else if (startInside || endInside) {
                                    ptEntry = startInside ? line.GetEndPoint(0) : line.GetEndPoint(1);
                                    ptExit = ptEntry;
                                }
                                double segmentLen = ptEntry.DistanceTo(ptExit);
                                if (segmentLen < UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters)) {
                                    DebugLogger.Log($"[PipeSleeveCommand] X-pipe {pipe.Id.IntegerValue}: Intersection segment too short, skipping.");
                                    continue;
                                }
                                var wallLocCurve = hostWallEntryExit.Location as LocationCurve;
                                XYZ wallCenter = null;
                                if (wallLocCurve != null && wallLocCurve.Curve != null)
                                    wallCenter = wallLocCurve.Curve.Evaluate(0.5, true);
                                XYZ sleevePoint;
                                string placementMethod = "";
                                if (ptEntry != null && ptExit != null && ptEntry.DistanceTo(ptExit) > UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters)) {
                                    // --- UPDATED LOGIC: Always place at wall centerline, not midpoint ---
                                    XYZ wallNormal = hostWallEntryExit.Orientation.Normalize();
                                    double wallWidth = hostWallEntryExit.Width;
                                    // Use existing wallLocCurve and wallCenter
                                    if (wallLocCurve != null && wallLocCurve.Curve != null)
                                        wallCenter = wallLocCurve.Curve.Evaluate(0.5, true);
                                    if (wallCenter != null) {
                                        // Project ptEntry onto wall centerline
                                        double distToCenter = (ptEntry - wallCenter).DotProduct(wallNormal);
                                        sleevePoint = ptEntry - wallNormal.Multiply(distToCenter);
                                        placementMethod = $"projected entry {ptEntry} to wall centerline {wallCenter}";
                                        DebugLogger.Log($"[PipeSleeveCommand] X-pipe: Projected entry to wall centerline: {sleevePoint}");
                                    } else {
                                        sleevePoint = (ptEntry + ptExit) * 0.5;
                                        placementMethod = $"midpoint between entry {ptEntry} and exit {ptExit} (fallback)";
                                        DebugLogger.Log($"[PipeSleeveCommand] X-pipe: Wall centerline not found, using midpoint: {sleevePoint}");
                                    }
                                } else if (ptEntry != null) {
                                    XYZ wallNormal = hostWallEntryExit.Orientation.Normalize();
                                    // Use existing wallLocCurve and wallCenter
                                    if (wallLocCurve != null && wallLocCurve.Curve != null)
                                        wallCenter = wallLocCurve.Curve.Evaluate(0.5, true);
                                    if (wallCenter != null) {
                                        double distToCenterInner = (ptEntry - wallCenter).DotProduct(wallNormal);
                                        sleevePoint = ptEntry - wallNormal.Multiply(distToCenterInner);
                                        placementMethod = $"projected entry {ptEntry} to wall centerline {wallCenter}";
                                        DebugLogger.Log($"[PipeSleeveCommand] X-pipe: Projected entry to wall center: {sleevePoint}");
                                    } else {
                                        sleevePoint = ptEntry;
                                        placementMethod = $"entry point {ptEntry}";
                                        DebugLogger.Log($"[PipeSleeveCommand] X-pipe: Using entry point for sleeve placement: {sleevePoint}");
                                    }
                                } else {
                                    DebugLogger.Log($"[PipeSleeveCommand] X-pipe {pipe.Id.IntegerValue}: No valid intersection for sleeve placement");
                                    continue;
                                }
                                // Log wall face normal at placement
                                XYZ faceNormal = null;
                                if (wallSolid != null && sleevePoint != null) {
                                    faceNormal = GetWallFaceNormal(wallSolid, sleevePoint);
                                    DebugLogger.Log($"[PipeSleeveCommand] X-pipe: Wall face normal at placement: {faceNormal}");
                                }
                                // Log wall type and exterior/interior if available
                                string wallType = hostWallEntryExit.WallType?.Name ?? "Unknown";
                                string wallFunction = "Unknown";
                                var functionParam = hostWallEntryExit.get_Parameter(BuiltInParameter.FUNCTION_PARAM);
                                if (functionParam != null && functionParam.HasValue) {
                                    int funcVal = functionParam.AsInteger();
                                    wallFunction = funcVal == 0 ? "Interior" : funcVal == 1 ? "Exterior" : $"Other({funcVal})";
                                }
                                DebugLogger.Log($"[PipeSleeveCommand] X-pipe: Placing sleeve in wall {hostWallEntryExit.Id.IntegerValue} (Type: {wallType}, Function: {wallFunction}), placement method: {placementMethod}");
                                double pipeDiameter = GetPipeDiameter(pipe) + totalClearance;
                                bool shouldRotate = false;
                                // --- CLUSTER SUPPRESSION: Use helper to check for any cluster at this location ---
                                double clusterTolerance = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                                bool clusterFound = false;
                                if (clusterSymbolWall != null && OpeningDuplicationChecker.IsClusterAtLocation(doc, sleevePoint, clusterTolerance, clusterSymbolWall))
                                    clusterFound = true;
                                if (!clusterFound && clusterSymbolSlab != null && OpeningDuplicationChecker.IsClusterAtLocation(doc, sleevePoint, clusterTolerance, clusterSymbolSlab))
                                    clusterFound = true;
                                if (clusterFound)
                                {
                                    DebugLogger.Log($"Pipe ID={pipe.Id.IntegerValue}: Skipping sleeve placement because cluster sleeve found at location (helper check)");
                                    continue;
                                }
                                placer.PlaceSleeve(
                                    sleevePoint,
                                    pipeDiameter,
                                    rayDirection,
                                    pipeSymbol,
                                    hostWallEntryExit,
                                    shouldRotate,
                                    pipe.Id.IntegerValue
                                );
                                placed = true;
                                placedCount++;
                                intersectionCount++;
                                DebugLogger.Log($"X-direction pipe sleeve placed for pipe {pipe.Id.IntegerValue} at {sleevePoint}");
                                break;
                            } else {
                                DebugLogger.Log($"[PipeSleeveCommand] X-pipe {pipe.Id.IntegerValue}: No wall solid found, skipping.");
                            }
                        } // <-- close foreach (var group in grouped)
                        if (!placed) {
                            DebugLogger.Log($"[PipeSleeveCommand] No valid X-direction intersection found for pipe {pipe.Id.IntegerValue} after all fallback methods. Sleeve NOT placed.");
                            // Log all wall and pipe geometry for debugging
                            foreach (var group in grouped) {
                                var hits = group.OrderBy(h => h.Proximity).ToList();
                                var rEntry = hits.First().GetReference();
                                var linkInstEntryExit = doc.GetElement(rEntry.ElementId) as RevitLinkInstance;
                                var targetDocEntryExit = linkInstEntryExit != null ? linkInstEntryExit.GetLinkDocument() : doc;
                                var eidEntryExit = linkInstEntryExit != null ? rEntry.LinkedElementId : rEntry.ElementId;
                                var hostWallEntryExit = targetDocEntryExit?.GetElement(eidEntryExit) as Wall;
                                if (hostWallEntryExit == null) continue;
                                var wallSolid = hostWallEntryExit.get_Geometry(new Options { ComputeReferences = true, IncludeNonVisibleObjects = false })
                                    .OfType<Solid>().FirstOrDefault(s => s.Volume > 0);
                                if (wallSolid != null) {
                                    var wallBounds = wallSolid.GetBoundingBox();
                                    DebugLogger.Log($"[PipeSleeveCommand] Wall {hostWallEntryExit.Id.IntegerValue} bounding box: Min={wallBounds.Min}, Max={wallBounds.Max}, Orientation={hostWallEntryExit.Orientation}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // Y-DIRECTION LOGIC (similar structure to X-direction but adapted for Y)
                        DebugLogger.Log($"[PipeSleeveCommand] Entering Y-direction logic for pipe {pipe.Id.IntegerValue}");
                        DebugLogger.Log($"[PipeSleeveCommand] Pipe direction: {rayDirection}, Start: {line.GetEndPoint(0)}, End: {line.GetEndPoint(1)}");
                        
                        var sampleFractions = new[] { 0.0, 0.25, 0.5, 0.75, 1.0 };
                        var allHits = new List<ReferenceWithContext>();
                        
                        // Cast rays in both positive and negative directions from each sample point
                        foreach (double t in sampleFractions)
                        {
                            var samplePt = line.Evaluate(t, true);
                            DebugLogger.Log($"[PipeSleeveCommand] Sampling at t={t}, point={samplePt}");
                            
                            // Cast in pipe direction and opposite direction
                            var hitsFwd = refIntersector.Find(samplePt, rayDirection)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            var hitsBack = refIntersector.Find(samplePt, rayDirection.Negate())?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            
                            // Also cast in perpendicular directions
                            var perpDir1 = new XYZ(-rayDirection.Y, rayDirection.X, 0).Normalize();
                            var perpDir2 = perpDir1.Negate();
                            var hitsPerp1 = refIntersector.Find(samplePt, perpDir1)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            var hitsPerp2 = refIntersector.Find(samplePt, perpDir2)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            
                            DebugLogger.Log($"[PipeSleeveCommand] Sample {t}: hitsFwd={hitsFwd.Count}, hitsBack={hitsBack.Count}, hitsPerp1={hitsPerp1.Count}, hitsPerp2={hitsPerp2.Count}");
                            
                            allHits.AddRange(hitsFwd);
                            allHits.AddRange(hitsBack);
                            allHits.AddRange(hitsPerp1);
                            allHits.AddRange(hitsPerp2);
                        }
                        
                        var grouped = allHits.GroupBy(h => {
                            var r = h.GetReference();
                            var linkInst = doc.GetElement(r.ElementId) as RevitLinkInstance;
                            return linkInst != null ? r.LinkedElementId : r.ElementId;
                        });
                        
                        bool placedY = false;
                        double pipeLength = line.Length;
                        
                        // --- BEGIN DUCT-STYLE LOGIC FOR Y-PIPES ---
                        DebugLogger.Log($"[PipeSleeveCommand] USING DUCT LOGIC FOR PIPE INTERSECTION (Y-direction)");
                        refWithContext = null;
                        foreach (double t in new[] { 0.0, 0.25, 0.5, 0.75, 1.0 })
                        {
                            var samplePt = line.Evaluate(t, true);
                            var testDirections = new[] { rayDirection, rayDirection.Negate() };
                            foreach (var testDir in testDirections)
                            {
                                var hits = refIntersector.Find(samplePt, testDir);
                                if (hits?.Count > 0)
                                {
                                    refWithContext = hits;
                                    DebugLogger.Log($"[PipeSleeveCommand] Y-pipe found wall hits at t={t} with direction {testDir}: {hits.Count} hits");
                                    break;
                                }
                            }
                            if (refWithContext?.Count > 0) break;
                        }
                        if (refWithContext == null || refWithContext.Count == 0)
                        {
                            DebugLogger.Log($"[PipeSleeveCommand] No wall hits found for Y-pipe {pipe.Id.IntegerValue}, skipping");
                            continue;
                        }
                        foreach (var wallHit in refWithContext)
                        {
                            var r = wallHit.GetReference();
                            var linkInst = doc.GetElement(r.ElementId) as RevitLinkInstance;
                            var targetDoc = linkInst != null ? linkInst.GetLinkDocument() : doc;
                            var eid = linkInst != null ? r.LinkedElementId : r.ElementId;
                            var hostWall = targetDoc?.GetElement(eid) as Wall;
                            if (hostWall == null) continue;
                            var pipeLine = locCurve;
                            double pipeLineLength = pipeLine.Length;
                            if (pipeLineLength < 0.01) {
                                DebugLogger.Log($"[PipeSleeveCommand] Skipping very short Y-pipe {pipe.Id.IntegerValue} with length {pipeLineLength:F6}");
                                continue;
                            }
                            Solid wallSolid = null;
                            try {
                                Options geomOptions = new Options { ComputeReferences = true, IncludeNonVisibleObjects = false };
                                GeometryElement geomElem = hostWall.get_Geometry(geomOptions);
                                foreach (GeometryObject obj in geomElem) {
                                    if (obj is Solid solid && solid.Volume > 0) {
                                        wallSolid = solid;
                                        break;
                                    }
                                }
                            } catch { wallSolid = null; }
                            if (wallSolid == null) continue;
                            List<XYZ> intersectionPoints = new List<XYZ>();
                            foreach (Face face in wallSolid.Faces)
                            {
                                IntersectionResultArray ira = null;
                                SetComparisonResult res = face.Intersect(pipeLine, out ira);
                                if (res == SetComparisonResult.Overlap && ira != null)
                                {
                                    foreach (IntersectionResult ir in ira)
                                    {
                                        var intersectionPoint = GetIntersectionPoint(ir);
                                        if (intersectionPoint != null)
                                        {
                                            intersectionPoints.Add(intersectionPoint);
                                        }
                                    }
                                }
                            }
                            bool startInside = IsPointInsideSolid(wallSolid, pipeLine.GetEndPoint(0), hostWall.Orientation);
                            bool endInside = IsPointInsideSolid(wallSolid, pipeLine.GetEndPoint(1), hostWall.Orientation);
                            // Segmented fallback
                            if (intersectionPoints.Count == 0 && !startInside && !endInside)
                            {
                                List<XYZ> altIntersections = new List<XYZ>();
                                int segments = 10;
                                for (int i = 0; i < segments; i++)
                                {
                                    double t1 = (double)i / segments;
                                    double t2 = (double)(i + 1) / segments;
                                    XYZ pt1 = pipeLine.Evaluate(t1, true);
                                    XYZ pt2 = pipeLine.Evaluate(t2, true);
                                    double segmentLength = pt1.DistanceTo(pt2);
                                    if (segmentLength < 0.01) continue;
                                    Line segment;
                                    try { segment = Line.CreateBound(pt1, pt2); }
                                    catch { continue; }
                                    foreach (Face face in wallSolid.Faces)
                                    {
                                        IntersectionResultArray ira2 = null;
                                        SetComparisonResult res2 = face.Intersect(segment, out ira2);
                                        if (res2 == SetComparisonResult.Overlap && ira2 != null)
                                        {
                                            foreach (IntersectionResult ir in ira2)
                                            {
                                                var intersectionPoint = GetIntersectionPoint(ir);
                                                if (intersectionPoint != null && !altIntersections.Any(pt => pt.DistanceTo(intersectionPoint) < UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters)))
                                                {
                                                    altIntersections.Add(intersectionPoint);
                                                }
                                            }
                                        }
                                    }
                                }
                                // Bounds fallback
                                if (altIntersections.Count == 0)
                                {
                                    BoundingBoxXYZ wallBounds = wallSolid.GetBoundingBox();
                                    if (wallBounds != null)
                                    {
                                        double tolerance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters);
                                        XYZ min = wallBounds.Min - new XYZ(tolerance, tolerance, tolerance);
                                        XYZ max = wallBounds.Max + new XYZ(tolerance, tolerance, tolerance);
                                        XYZ pipeStart = pipeLine.GetEndPoint(0);
                                        XYZ pipeEnd = pipeLine.GetEndPoint(1);
                                        if (IsPointInBounds(pipeStart, min, max) != IsPointInBounds(pipeEnd, min, max))
                                        {
                                            var wallLocCurveAlt = hostWall.Location as LocationCurve;
                                            if (wallLocCurveAlt != null && wallLocCurveAlt.Curve != null)
                                            {
                                                XYZ wallCenterAlt2 = wallLocCurveAlt.Curve.Evaluate(0.5, true);
                                                XYZ wallNormal = hostWall.Orientation.Normalize();
                                                XYZ pipeDirection = (pipeEnd - pipeStart).Normalize();
                                                double denominator = pipeDirection.DotProduct(wallNormal);
                                                if (Math.Abs(denominator) > 1e-6)
                                                {
                                                    double t = (wallCenterAlt2 - pipeStart).DotProduct(wallNormal) / denominator;
                                                    if (t >= 0 && t <= 1)
                                                    {
                                                        XYZ artificialIntersection = pipeStart + pipeDirection.Multiply(t * pipeStart.DistanceTo(pipeEnd));
                                                        altIntersections.Add(artificialIntersection);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (altIntersections.Count > 0)
                                {
                                    intersectionPoints = altIntersections;
                                }
                            }
                            if (intersectionPoints.Count == 0 && !startInside && !endInside) continue;
                            XYZ ptEntry = null, ptExit = null;
                            if (intersectionPoints.Count >= 2)
                            {
                                intersectionPoints = intersectionPoints.OrderBy(pt => (pt - pipeLine.GetEndPoint(0)).GetLength()).ToList();
                                ptEntry = intersectionPoints.First();
                                ptExit = intersectionPoints.Last();
                            }
                            else if (intersectionPoints.Count == 1)
                            {
                                ptEntry = intersectionPoints[0];
                                ptExit = startInside ? pipeLine.GetEndPoint(0) : pipeLine.GetEndPoint(1);
                            }
                            else if (startInside || endInside)
                            {
                                ptEntry = startInside ? pipeLine.GetEndPoint(0) : pipeLine.GetEndPoint(1);
                                ptExit = ptEntry;
                            }
                            double segmentLen = ptEntry.DistanceTo(ptExit);
                            if (segmentLen < UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters)) continue;
                            var wallLocCurve = hostWall.Location as LocationCurve;
                            XYZ wallCenter = null;
                            if (wallLocCurve != null && wallLocCurve.Curve != null)
                                wallCenter = wallLocCurve.Curve.Evaluate(0.5, true);
                            XYZ sleevePoint;
                            string placementMethod = "";
                            if (ptEntry != null && ptExit != null && ptEntry.DistanceTo(ptExit) > UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters)) {
                                sleevePoint = (ptEntry + ptExit) * 0.5;
                                placementMethod = $"midpoint between entry {ptEntry} and exit {ptExit}";
                                DebugLogger.Log($"[PipeSleeveCommand] {(isXOrientation ? "X" : "Y")}-pipe: Using midpoint for sleeve placement: {sleevePoint}");
                            } else if (ptEntry != null) {
                                XYZ wallCenterInner = null;
                                XYZ wallNormal = hostWall.Orientation.Normalize();
                                if (wallLocCurve != null && wallLocCurve.Curve != null)
                                    wallCenterInner = wallLocCurve.Curve.Evaluate(0.5, true);
                                if (wallCenterInner != null) {
                                    double distToCenterInner = (ptEntry - wallCenterInner).DotProduct(wallNormal);
                                    sleevePoint = ptEntry - wallNormal.Multiply(distToCenterInner);
                                    placementMethod = $"projected entry {ptEntry} to wall centerline {wallCenterInner}";
                                    DebugLogger.Log($"[PipeSleeveCommand] {(isXOrientation ? "X" : "Y")}-pipe: Projected entry to wall center: {sleevePoint}");
                                } else {
                                    sleevePoint = ptEntry;
                                    placementMethod = $"entry point {ptEntry}";
                                    DebugLogger.Log($"[PipeSleeveCommand] {(isXOrientation ? "X" : "Y")}-pipe: Using entry point for sleeve placement: {sleevePoint}");
                                }
                            } else {
                                DebugLogger.Log($"[PipeSleeveCommand] {(isXOrientation ? "X" : "Y")}-pipe {pipe.Id.IntegerValue}: No valid intersection for sleeve placement");
                                continue;
                            }
                            // Log wall face normal at placement
                            XYZ faceNormal = null;
                            if (wallSolid != null && sleevePoint != null) {
                                faceNormal = GetWallFaceNormal(wallSolid, sleevePoint);
                                DebugLogger.Log($"[PipeSleeveCommand] Y-pipe: Wall face normal at placement: {faceNormal}");
                            }
                            // Log wall type and exterior/interior if available
                            string wallType = hostWall.WallType?.Name ?? "Unknown";
                            string wallFunction = "Unknown";
                            var functionParam = hostWall.get_Parameter(BuiltInParameter.FUNCTION_PARAM);
                            if (functionParam != null && functionParam.HasValue) {
                                int funcVal = functionParam.AsInteger();
                                wallFunction = funcVal == 0 ? "Interior" : funcVal == 1 ? "Exterior" : $"Other({funcVal})";
                            }
                            DebugLogger.Log($"[PipeSleeveCommand] Y-pipe: Placing sleeve in wall {hostWall.Id.IntegerValue} (Type: {wallType}, Function: {wallFunction}), placement method: {placementMethod}");
                            double pipeDiameter = GetPipeDiameter(pipe) + totalClearance;
                            bool shouldRotate = true;
                            // --- CLUSTER SUPPRESSION: Use helper to check for any cluster at this location ---
                            double clusterTolerance = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                            // --- DUPLICATE SUPPRESSION: Check for existing sleeve at this location ---
                            bool hasExistingSleeveAtPlacement = existingSleeves.Any(sleeveInst =>
                                ((sleeveInst.Location as LocationPoint)?.Point ?? sleeveInst.GetTransform().Origin).DistanceTo(sleevePoint) <= clusterTolerance);
                            bool clusterFound = false;
                            if (hasExistingSleeveAtPlacement)
                            {
                                DebugLogger.Log($"Pipe ID={pipe.Id.IntegerValue}: Skipping sleeve placement because existing sleeve found within {UnitUtils.ConvertFromInternalUnits(clusterTolerance, UnitTypeId.Millimeters):F0}mm at placement point");
                                skippedExistingCount++;
                                continue;
                            }
                            if (clusterSymbolWall != null && OpeningDuplicationChecker.IsClusterAtLocation(doc, sleevePoint, clusterTolerance, clusterSymbolWall))
                                clusterFound = true;
                            if (!clusterFound && clusterSymbolSlab != null && OpeningDuplicationChecker.IsClusterAtLocation(doc, sleevePoint, clusterTolerance, clusterSymbolSlab))
                                clusterFound = true;
                            if (clusterFound)
                            {
                                DebugLogger.Log($"Pipe ID={pipe.Id.IntegerValue}: Skipping sleeve placement because cluster sleeve found at location (helper check)");
                                continue;
                            }
                            placer.PlaceSleeve(
                                sleevePoint,
                                pipeDiameter,
                                rayDirection,
                                pipeSymbol,
                                hostWall,
                                shouldRotate,
                                pipe.Id.IntegerValue
                            );
                            placedY = true;
                            placedCount++;
                            intersectionCount++;
                            DebugLogger.Log($"Y-direction pipe sleeve placed for pipe {pipe.Id.IntegerValue} at {sleevePoint}");
                            break;
                        }
                        if (!placedY) {
                            DebugLogger.Log($"[PipeSleeveCommand] No valid Y-direction intersection found for pipe {pipe.Id.IntegerValue} after all fallback methods. Sleeve NOT placed.");
                            DebugLogger.Log($"[PipeSleeveCommand] Y-pipe {pipe.Id.IntegerValue} endpoints: {line.GetEndPoint(0)} to {line.GetEndPoint(1)}");
                            // Do NOT place sleeve at midpoint or wall centerline for Y-pipes if no intersection found
                            // Only log and skip placement
                            continue;
                        }
                    }
                    // Mark pipe as processed
                    processedPipes.Add(pipe.Id);
                }
                tx.Commit();
            }

            // Summary report
            string summary = $"Pipe Sleeve Placement Complete:\n" +
                             $"Total Pipes Processed: {totalPipes}\n" +
                             $"Sleeves Placed: {placedCount}\n" +
                             $"Intersections Found: {intersectionCount}\n" +
                             $"Pipes Missing Location Curve: {missingCount}\n" +
                             $"Skipped Existing Sleeves: {skippedExistingCount}";
            TaskDialog.Show("Summary", summary);

            DebugLogger.Log("PipeSleeveCommand completed");
            return Result.Succeeded;
        }

        private double GetPipeDiameter(Pipe pipe)
        {
            // Get the pipe diameter, considering insulation if present
            var diameterParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
            if (diameterParam != null && diameterParam.HasValue)
            {
                return diameterParam.AsDouble();
            }
            return 0.0;
        }

        private bool IsPointInsideSolid(Solid solid, XYZ point, XYZ wallNormal)
        {
            // Check if a point is inside a solid using a ray-casting method
            // Cast a ray in the direction of the wall normal
            XYZ rayDirection = wallNormal.Normalize();
            double tolerance = 1e-6;
            double maxDistance = 100.0; // Arbitrary large distance
            var origin = point + rayDirection.Multiply(tolerance);
            var end = point - rayDirection.Multiply(maxDistance);
            Line ray = Line.CreateBound(origin, end);

            // Count intersections with solid faces
            int intersectionCount = 0;
            foreach (Face face in solid.Faces)
            {
                IntersectionResultArray ira = null;
                SetComparisonResult res = face.Intersect(ray, out ira);
                if (res == SetComparisonResult.Overlap && ira != null)
                {
                    intersectionCount += ira.Size;
                }
            }

            // If odd number of intersections, point is inside
            return (intersectionCount % 2) == 1;
        }

        private XYZ GetWallFaceNormal(Solid wallSolid, XYZ point)
        {
            // Get the normal of the wall face at the given point
            foreach (Face face in wallSolid.Faces)
            {
                if (face != null)
                {
                    IntersectionResultArray ira = null;
                    SetComparisonResult res = face.Intersect(Line.CreateBound(point, point + XYZ.BasisZ), out ira);
                    if (res == SetComparisonResult.Overlap && ira != null && ira.Size > 0)
                    {
                        return face.ComputeNormal(new UV(0.5, 0.5));
                    }
                }
            }
            return null;
        }

        // Helper: Check if a point is within a bounding box
        private static bool IsPointInBounds(XYZ pt, XYZ min, XYZ max)
        {
            return pt.X >= min.X && pt.X <= max.X &&
                   pt.Y >= min.Y && pt.Y <= max.Y &&
                   pt.Z >= min.Z && pt.Z <= max.Z;
        }

        // Helper method to get XYZ point from IntersectionResult (compatibility across Revit versions)
        private static XYZ GetIntersectionPoint(IntersectionResult ir)
        {
            // Try different property names for different Revit versions
            try
            {
                // For newer versions, try Point property first (Revit 2024+)
                var pointProperty = ir.GetType().GetProperty("Point");
                if (pointProperty != null)
                    return (XYZ)pointProperty.GetValue(ir);
            }
            catch { }
            
            try
            {
                // For Revit 2020-2023, try XYZPoint property
                var xyzPointProperty = ir.GetType().GetProperty("XYZPoint");
                if (xyzPointProperty != null)
                    return (XYZ)xyzPointProperty.GetValue(ir);
            }
            catch { }
            
            // Fallback - this should not happen if API is consistent
            throw new InvalidOperationException("Unable to get intersection point from IntersectionResult");
        }
    }
}
