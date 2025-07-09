using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DuctSleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DebugLogger.Log("Starting DuctSleeveCommand");
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;            // Select the duct sleeve symbol (DS# in OpeningOnWall family)
            var ductSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family != null
                                       && sym.Family.Name.IndexOf("OpeningOnWall", System.StringComparison.OrdinalIgnoreCase) >= 0
                                       && sym.Name.Replace(" ", "").StartsWith("DS#", System.StringComparison.OrdinalIgnoreCase));
            if (ductSymbol == null)
            {
                TaskDialog.Show("Error", "No duct opening family symbol (DS#) found.");
                return Result.Failed;
            }
            using (var txActivate = new Transaction(doc, "Activate Duct Symbol"))
            {
                txActivate.Start();
                if (!ductSymbol.IsActive)
                    ductSymbol.Activate();
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
            };            // Sleeve placer
            var placer = new DuctSleevePlacer(doc);

            // Collect all ducts
            var ducts = new FilteredElementCollector(doc)
                .OfClass(typeof(Duct))
                .WhereElementIsNotElementType()
                .Cast<Duct>();

            // SUPPRESSION: Collect existing duct sleeves AND cluster sleeves to avoid duplicates
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("DS#"))
                .ToList();
            
            // Also collect cluster sleeves to prevent individual sleeves from being placed inside them
            // Check for family names containing "ONWALL" (for wall/structural framing) or "ONSLAB" (for floor)
            
            // First, let's debug what family instances exist to understand the naming pattern
            var allFamilyInstances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Name != null)
                .ToList();
            
            DebugLogger.Log($"[CLUSTER_DEBUG] Total family instances in model: {allFamilyInstances.Count}");
            
            // Log all family names that might be cluster sleeves
            var potentialClusterFamilies = allFamilyInstances
                .Where(fi => fi.Symbol.Family.Name.IndexOf("ONWALL", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fi.Symbol.Family.Name.IndexOf("ONSLAB", StringComparison.OrdinalIgnoreCase) >= 0 ||
                           fi.Symbol.Family.Name.IndexOf("CLUSTER", StringComparison.OrdinalIgnoreCase) >= 0)
                .GroupBy(fi => fi.Symbol.Family.Name)
                .ToList();
            
            DebugLogger.Log($"[CLUSTER_DEBUG] Found {potentialClusterFamilies.Count} potential cluster family types:");
            foreach (var familyGroup in potentialClusterFamilies)
            {
                DebugLogger.Log($"[CLUSTER_DEBUG] Family: '{familyGroup.Key}' - Count: {familyGroup.Count()}");
                var firstInstance = familyGroup.First();
                DebugLogger.Log($"[CLUSTER_DEBUG]   Symbol Name: '{firstInstance.Symbol.Name}'");
                DebugLogger.Log($"[CLUSTER_DEBUG]   Category: '{firstInstance.Category?.Name ?? "Unknown"}'");
            }
            
            var existingClusterSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Name != null && 
                           (string.Equals(fi.Symbol.Family.Name, "DuctOpeningOnWall", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(fi.Symbol.Family.Name, "DuctOpeningOnSlab", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(fi.Symbol.Family.Name, "PipeOpeningOnWallRect", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(fi.Symbol.Family.Name, "PipeOpeningOnSlabRect", StringComparison.OrdinalIgnoreCase)))
                .ToList();
            
            DebugLogger.Log($"[CLUSTER_DEBUG] Found {existingSleeves.Count} existing duct sleeves and {existingClusterSleeves.Count} cluster sleeves in the model");
            
            // Log details of each cluster sleeve found
            foreach (var clusterSleeve in existingClusterSleeves)
            {
                var location = (clusterSleeve.Location as LocationPoint)?.Point ?? clusterSleeve.GetTransform().Origin;
                var boundingBox = clusterSleeve.get_BoundingBox(null);
                DebugLogger.Log($"[CLUSTER_DEBUG] Cluster Sleeve ID={clusterSleeve.Id.IntegerValue}:");
                DebugLogger.Log($"[CLUSTER_DEBUG]   Family: '{clusterSleeve.Symbol.Family.Name}'");
                DebugLogger.Log($"[CLUSTER_DEBUG]   Symbol: '{clusterSleeve.Symbol.Name}'");
                DebugLogger.Log($"[CLUSTER_DEBUG]   Location: {location}");
                if (boundingBox != null)
                {
                    DebugLogger.Log($"[CLUSTER_DEBUG]   BoundingBox: Min={boundingBox.Min}, Max={boundingBox.Max}");
                }
                else
                {
                    DebugLogger.Log($"[CLUSTER_DEBUG]   BoundingBox: NULL");
                }
            }
            
            // Create a map of existing sleeve locations for quick lookup
            var existingSleeveLocations = existingSleeves.ToDictionary(
                sleeve => sleeve,
                sleeve => (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin
            );
            
            // Create a map of cluster sleeve locations with larger tolerance zones
            var existingClusterLocationsWithBounds = existingClusterSleeves.ToDictionary(
                clusterSleeve => clusterSleeve,
                clusterSleeve => new {
                    Location = (clusterSleeve.Location as LocationPoint)?.Point ?? clusterSleeve.GetTransform().Origin,
                    BoundingBox = clusterSleeve.get_BoundingBox(null)
                }
            );

            // Initialize counters for detailed logging
            int totalDucts = ducts.Count();
            int intersectionCount = 0;
            int placedCount = 0;
            int missingCount = 0;
            int damperSkippedCount = 0;
            int skippedExistingCount = 0; // Counter for ducts with existing sleeves
            HashSet<ElementId> processedDucts = new HashSet<ElementId>(); // Track processed ducts to prevent duplicates
            DebugLogger.Log($"Found {totalDucts} ducts to process");            // Start transaction for placement
            using (var tx = new Transaction(doc, "Place Duct Sleeves"))
            {
                tx.Start();
                foreach (var duct in ducts)
                {
                    DebugLogger.Log($"Processing Duct ID={duct.Id.IntegerValue}");
                    
                    // Prevent processing the same duct twice (avoid duplicate sleeves)
                    if (processedDucts.Contains(duct.Id))
                    {
                        DebugLogger.Log($"Duct ID={duct.Id.IntegerValue}: already processed, skipping");
                        continue;
                    }
                    
                    var locCurve = (duct.Location as LocationCurve)?.Curve as Line;
                    if (locCurve == null)
                    {
                        DebugLogger.Log($"Duct ID={duct.Id.IntegerValue}: no curve, skipping");
                        missingCount++;
                        continue;
                    }

                    // NOTE: Suppression check moved to actual placement points to prevent duplicates
                    var line = locCurve;

                    // Enhanced wall intersection: use robust intersection logic
                    XYZ rayDirection = line.Direction;
                    XYZ perpDirection = new XYZ(-rayDirection.Y, rayDirection.X, 0).Normalize();
                    IList<ReferenceWithContext> refWithContext;
                    bool isXOrientation = Math.Abs(rayDirection.X) > Math.Abs(rayDirection.Y);

                    if (isXOrientation)
                    {
                        DebugLogger.Log($"[DuctSleeveCommand] Entering X-direction logic for duct {duct.Id.IntegerValue}");
                        DebugLogger.Log($"[DuctSleeveCommand] Duct direction: {rayDirection}, Start: {line.GetEndPoint(0)}, End: {line.GetEndPoint(1)}");
                        
                        var sampleFractions = new[] { 0.0, 0.25, 0.5, 0.75, 1.0 };
                        var allHits = new List<ReferenceWithContext>();
                        
                        // Cast rays in both positive and negative directions from each sample point
                        // This ensures we catch walls regardless of duct direction
                        foreach (double t in sampleFractions)
                        {
                            var samplePt = line.Evaluate(t, true);
                            DebugLogger.Log($"[DuctSleeveCommand] Sampling at t={t}, point={samplePt}");
                            
                            // Cast in duct direction and opposite direction
                            var hitsFwd = refIntersector.Find(samplePt, rayDirection)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            var hitsBack = refIntersector.Find(samplePt, rayDirection.Negate())?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            
                            // Also cast in perpendicular directions to catch walls parallel to duct
                            var perpDir1 = new XYZ(-rayDirection.Y, rayDirection.X, 0).Normalize();
                            var perpDir2 = perpDir1.Negate();
                            var hitsPerp1 = refIntersector.Find(samplePt, perpDir1)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            var hitsPerp2 = refIntersector.Find(samplePt, perpDir2)?.Where(h => h != null).ToList() ?? new List<ReferenceWithContext>();
                            
                            DebugLogger.Log($"[DuctSleeveCommand] Sample {t}: hitsFwd={hitsFwd.Count}, hitsBack={hitsBack.Count}, hitsPerp1={hitsPerp1.Count}, hitsPerp2={hitsPerp2.Count}");
                            
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
                        double ductLength = line.Length;
                        
                        // --- BEGIN ROBUST SOLID INTERSECTION LOGIC FOR X-DUCTS ---
                        var bestWall = (Wall)null;
                        XYZ bestEntry = null, bestExit = null;                        XYZ bestFaceNormal = null;
                        double bestWallThickness = 0.0;
                        double maxSegment = 0.0;                        var ductLine = locCurve;
                        
                        // Validate duct line length to prevent short curve errors (use existing ductLength variable)
                        if (ductLength < 0.01) // Less than 0.01 feet (about 3mm)
                        {
                            DebugLogger.Log($"[DuctSleeveCommand] Skipping very short duct {duct.Id.IntegerValue} with length {ductLength:F6}");
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
                            if (hostWallEntryExit == null) continue;
                            
                            double wallThickness = hostWallEntryExit.Width;
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
                            if (wallSolid == null) continue;
                            
                            // DEBUG: Wall geometry information
                            bool isLinkedWall = linkInstEntryExit != null;
                            string wallTypeInfo = hostWallEntryExit.WallType?.Name ?? "Unknown";
                            DebugLogger.Log($"[DUCT_WALL_DEBUG] Wall ID={hostWallEntryExit.Id.IntegerValue}, IsLinked={isLinkedWall}, Type={wallTypeInfo}");
                            DebugLogger.Log($"[DUCT_WALL_DEBUG] Wall Width={wallThickness}, Solid Volume={wallSolid.Volume}, Face Count={wallSolid.Faces.Size}");
                            
                            // Intersect duct centerline with wall solid faces
                            List<XYZ> intersectionPoints = new List<XYZ>();
                            foreach (Face face in wallSolid.Faces)
                            {
                                IntersectionResultArray ira = null;
                                SetComparisonResult res = face.Intersect(ductLine, out ira);
                                if (res == SetComparisonResult.Overlap && ira != null)
                                {
                                    foreach (IntersectionResult ir in ira)
                                    {
                                        var intersectionPoint = GetIntersectionPoint(ir);
                                        if (intersectionPoint != null)
                                        {
                                            intersectionPoints.Add(intersectionPoint);
                                            DebugLogger.Log($"[DEBUG] Intersection found: Duct {duct.Id.IntegerValue} with Wall {hostWallEntryExit.Id.IntegerValue} at {intersectionPoint}");
                                        }
                                    }
                                }
                            }
                            
                            // Also check if duct endpoint is inside wall (for stub ducts)
                            bool startInside = IsPointInsideSolid(wallSolid, ductLine.GetEndPoint(0), hostWallEntryExit.Orientation);
                            bool endInside = IsPointInsideSolid(wallSolid, ductLine.GetEndPoint(1), hostWallEntryExit.Orientation);
                            DebugLogger.Log($"[DEBUG] Duct {duct.Id.IntegerValue} - Wall {hostWallEntryExit.Id.IntegerValue}: intersectionPoints.Count={intersectionPoints.Count}, startInside={startInside}, endInside={endInside}");
                            
                            // ENHANCED INTERSECTION DETECTION: Try alternative methods if basic approach fails
                            if (intersectionPoints.Count == 0 && !startInside && !endInside)
                            {                                // Method 2: Try smaller line segments
                                List<XYZ> altIntersections = new List<XYZ>();
                                int segments = 10;
                                for (int i = 0; i < segments; i++)
                                {
                                    double t1 = (double)i / segments;
                                    double t2 = (double)(i + 1) / segments;
                                    XYZ pt1 = ductLine.Evaluate(t1, true);
                                    XYZ pt2 = ductLine.Evaluate(t2, true);
                                    
                                    // Check if segment is long enough to avoid Revit tolerance error
                                    double segmentLength = pt1.DistanceTo(pt2);
                                    if (segmentLength < 0.01) // Increased to 0.01 feet (about 3mm) - more conservative
                                    {
                                        DebugLogger.Log($"[DEBUG] Skipping X-segment {i} with length {segmentLength:F9} (too short)");
                                        continue;
                                    }
                                    
                                    Line segment;
                                    try
                                    {
                                        segment = Line.CreateBound(pt1, pt2);
                                    }
                                    catch (Exception segmentEx)
                                    {
                                        DebugLogger.Log($"[DEBUG] Failed to create X-segment {i}: {segmentEx.Message}, pt1={pt1}, pt2={pt2}, distance={segmentLength:F9}");
                                        continue;
                                    }
                                    
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
                                                    DebugLogger.Log($"[DEBUG] Alt intersection found: Duct {duct.Id.IntegerValue} with Wall {hostWallEntryExit.Id.IntegerValue} at segment {i} - {intersectionPoint}");
                                                }
                                            }
                                        }
                                    }
                                }
                                
                                // Method 3: Check if duct line passes through wall bounds with tolerance
                                if (altIntersections.Count == 0)
                                {
                                    BoundingBoxXYZ wallBounds = wallSolid.GetBoundingBox();
                                    if (wallBounds != null)
                                    {
                                        // Expand bounds slightly for tolerance
                                        double tolerance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters);
                                        XYZ min = wallBounds.Min - new XYZ(tolerance, tolerance, tolerance);
                                        XYZ max = wallBounds.Max + new XYZ(tolerance, tolerance, tolerance);
                                        
                                        // Check if duct line passes through expanded bounds
                                        XYZ ductStart = ductLine.GetEndPoint(0);
                                        XYZ ductEnd = ductLine.GetEndPoint(1);
                                        
                                        bool startInBounds = IsPointInBounds(ductStart, min, max);
                                        bool endInBounds = IsPointInBounds(ductEnd, min, max);
                                        
                                        DebugLogger.Log($"[DEBUG] Bounds check: Duct {duct.Id.IntegerValue} - Wall {hostWallEntryExit.Id.IntegerValue}: startInBounds={startInBounds}, endInBounds={endInBounds}");
                                        
                                        if (startInBounds != endInBounds) // Line crosses bounds
                                        {
                                            // Create artificial intersection at wall center plane
                                            var wallLocCurveAlt = hostWallEntryExit.Location as LocationCurve;
                                            if (wallLocCurveAlt != null && wallLocCurveAlt.Curve != null)
                                            {
                                                XYZ wallCenterAlt = wallLocCurveAlt.Curve.Evaluate(0.5, true);
                                                XYZ wallNormal = hostWallEntryExit.Orientation.Normalize();
                                                
                                                // Find intersection with wall center plane
                                                XYZ ductDirection = (ductEnd - ductStart).Normalize();
                                                double denominator = ductDirection.DotProduct(wallNormal);
                                                
                                                if (Math.Abs(denominator) > 1e-6) // Not parallel
                                                {
                                                    double t = (wallCenterAlt - ductStart).DotProduct(wallNormal) / denominator;
                                                    if (t >= 0 && t <= 1) // Intersection within duct segment
                                                    {
                                                        XYZ artificialIntersection = ductStart + ductDirection.Multiply(t * ductStart.DistanceTo(ductEnd));
                                                        altIntersections.Add(artificialIntersection);
                                                        DebugLogger.Log($"[DEBUG] Artificial intersection created: Duct {duct.Id.IntegerValue} with Wall {hostWallEntryExit.Id.IntegerValue} at {artificialIntersection}");
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                
                                if (altIntersections.Count > 0)
                                {
                                    intersectionPoints = altIntersections;
                                    DebugLogger.Log($"[DEBUG] Using alternative intersections: {intersectionPoints.Count} points found");
                                }
                                else
                                {
                                    DebugLogger.Log($"[DEBUG] No intersections found with any method - skipping duct {duct.Id.IntegerValue} with wall {hostWallEntryExit.Id.IntegerValue}");
                                    continue;
                                }
                            }
                            
                            if (intersectionPoints.Count == 0 && !startInside && !endInside) continue;
                            
                            // Use entry/exit or endpoint for centering
                            XYZ ptEntry = null, ptExit = null;
                            if (intersectionPoints.Count >= 2)
                            {
                                intersectionPoints = intersectionPoints.OrderBy(pt => (pt - ductLine.GetEndPoint(0)).GetLength()).ToList();
                                ptEntry = intersectionPoints.First();
                                ptExit = intersectionPoints.Last();
                                DebugLogger.Log($"[DEBUG] Duct {duct.Id.IntegerValue} - Wall {hostWallEntryExit.Id.IntegerValue}: ptEntry={ptEntry}, ptExit={ptExit}");
                            }
                            else if (intersectionPoints.Count == 1)
                            {
                                ptEntry = intersectionPoints[0];
                                ptExit = startInside ? ductLine.GetEndPoint(0) : ductLine.GetEndPoint(1);
                                DebugLogger.Log($"[DEBUG] Duct {duct.Id.IntegerValue} - Wall {hostWallEntryExit.Id.IntegerValue}: Single intersection, ptEntry={ptEntry}, ptExit={ptExit}");
                            }
                            else if (startInside || endInside)
                            {
                                ptEntry = startInside ? ductLine.GetEndPoint(0) : ductLine.GetEndPoint(1);
                                ptExit = ptEntry;
                                DebugLogger.Log($"[DEBUG] Duct {duct.Id.IntegerValue} - Wall {hostWallEntryExit.Id.IntegerValue}: Endpoint inside wall, ptEntry/ptExit={ptEntry}");
                            }
                            else
                            {
                                DebugLogger.Log($"[DEBUG] Duct {duct.Id.IntegerValue} - Wall {hostWallEntryExit.Id.IntegerValue}: No intersection or endpoint inside wall, skipping sleeve placement.");
                                continue;
                            }
                            
                            double segmentLen = ptEntry.DistanceTo(ptExit);
                            if (segmentLen > maxSegment)
                            {
                                maxSegment = segmentLen;
                                bestWall = hostWallEntryExit;
                                bestEntry = ptEntry;
                                bestExit = ptExit;
                                bestFaceNormal = hostWallEntryExit.Orientation.Normalize();
                                bestWallThickness = wallThickness;
                            }
                        }
                        
                        if (bestWall != null && maxSegment > UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters))
                        {
                            // FIXED CENTERING: Calculate direction toward wall center for proper sleeve placement
                            var wallLocCurveX = bestWall.Location as LocationCurve;
                            XYZ wallCenterX = null;
                            if (wallLocCurveX != null && wallLocCurveX.Curve != null)
                                wallCenterX = wallLocCurveX.Curve.Evaluate(0.5, true);
                            
                            XYZ sleevePointEntryExit;
                            if (wallCenterX != null) {
                                // Calculate direction from entry point to wall center
                                XYZ toCenter = wallCenterX - bestEntry;
                                // Project this direction onto the wall normal to get signed distance
                                double distanceToCenter = toCenter.DotProduct(bestFaceNormal);
                                // Move toward center by the projected distance
                                sleevePointEntryExit = bestEntry + bestFaceNormal.Multiply(distanceToCenter);
                                DebugLogger.Log($"[DuctSleeveCommand] X-duct: Centering sleeve - Entry={bestEntry}, WallCenter={wallCenterX}, DistanceToCenter={distanceToCenter}, FinalPosition={sleevePointEntryExit}");
                            } else {
                                // Fallback: use entry point if wall center cannot be determined
                                sleevePointEntryExit = bestEntry;
                                DebugLogger.Log($"[DuctSleeveCommand] X-duct: Wall center unavailable, using entry point");
                            }
                            
                            // Check for dampers near the intersection point
                            double damperTol = UnitUtils.ConvertToInternalUnits(75.0, UnitTypeId.Millimeters);
                            // Check for dampers truly inside the wall (not just near)
                            BoundingBoxXYZ wallBB = bestWall.get_BoundingBox(null);
                            bool hasDamper = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                                .OfClass(typeof(FamilyInstance))
                                .Cast<FamilyInstance>()
                                .Any(fi => fi.Symbol.Family.Name.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0
                                    && IsPointInsideBoundingBox(fi.GetTransform().Origin, wallBB));
                            if (hasDamper)
                            {
                                DebugLogger.Log($"Duct ID={duct.Id.IntegerValue}: damper present inside wall bounding box, skipping sleeve");
                                damperSkippedCount++;
                                continue;
                            }
                            
                            // Compute duct dimensions
                            double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                            double h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                            double clearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                            double sw = w + clearance * 2;
                            double sh = h + clearance * 2;

                            // ENHANCED SUPPRESSION CHECK: Skip if sleeve already exists OR if placement point is inside a cluster sleeve
                            double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters); // 1mm precision tolerance
                            bool hasExistingSleeveAtPlacement = existingSleeveLocations.Values.Any(sleeveLocation => 
                                sleevePointEntryExit.DistanceTo(sleeveLocation) <= sleeveCheckRadius);
                            
                            DebugLogger.Log($"[SUPPRESSION_DEBUG] Duct ID={duct.Id.IntegerValue} - X-orientation suppression check:");
                            DebugLogger.Log($"[SUPPRESSION_DEBUG]   Placement point: {sleevePointEntryExit}");
                            DebugLogger.Log($"[SUPPRESSION_DEBUG]   Existing sleeves: {existingSleeveLocations.Count}");
                            DebugLogger.Log($"[SUPPRESSION_DEBUG]   Cluster sleeves: {existingClusterLocationsWithBounds.Count}");
                            DebugLogger.Log($"[SUPPRESSION_DEBUG]   Has existing sleeve at placement: {hasExistingSleeveAtPlacement}");
                            
                            // Check if placement point is inside any cluster sleeve's bounding box
                            bool isInsideClusterSleeve = false;
                            int clusterIndex = 0;
                            foreach (var clusterInfo in existingClusterLocationsWithBounds.Values)
                            {
                                clusterIndex++;
                                DebugLogger.Log($"[SUPPRESSION_DEBUG]   Checking cluster sleeve #{clusterIndex}:");
                                DebugLogger.Log($"[SUPPRESSION_DEBUG]     Cluster location: {clusterInfo.Location}");
                                
                                if (clusterInfo.BoundingBox != null)
                                {
                                    // Expand cluster bounding box slightly for tolerance
                                    double clusterTolerance = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm tolerance - MATCHES cluster command tolerance
                                    XYZ minExpanded = clusterInfo.BoundingBox.Min - new XYZ(clusterTolerance, clusterTolerance, clusterTolerance);
                                    XYZ maxExpanded = clusterInfo.BoundingBox.Max + new XYZ(clusterTolerance, clusterTolerance, clusterTolerance);
                                    
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Original BoundingBox: Min={clusterInfo.BoundingBox.Min}, Max={clusterInfo.BoundingBox.Max}");
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Expanded BoundingBox: Min={minExpanded}, Max={maxExpanded}");
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Tolerance: {UnitUtils.ConvertFromInternalUnits(clusterTolerance, UnitTypeId.Millimeters):F0}mm");
                                    
                                    bool pointInBounds = IsPointInBounds(sleevePointEntryExit, minExpanded, maxExpanded);
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Point in expanded bounds: {pointInBounds}");
                                    
                                    if (pointInBounds)
                                    {
                                        isInsideClusterSleeve = true;
                                        DebugLogger.Log($"[SUPPRESSION_DEBUG] *** SUPPRESSED *** Duct ID={duct.Id.IntegerValue}: placement point inside cluster sleeve bounding box, skipping individual sleeve placement");
                                        break;
                                    }
                                }
                                else
                                {
                                    // Fallback: use distance check if bounding box is not available
                                    double distanceToCluster = sleevePointEntryExit.DistanceTo(clusterInfo.Location);
                                    double clusterRadius = UnitUtils.ConvertToInternalUnits(200.0, UnitTypeId.Millimeters); // 200mm radius
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     No bounding box - using distance check");
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Distance to cluster: {UnitUtils.ConvertFromInternalUnits(distanceToCluster, UnitTypeId.Millimeters):F1}mm");
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Cluster radius: {UnitUtils.ConvertFromInternalUnits(clusterRadius, UnitTypeId.Millimeters):F0}mm");
                                    
                                    if (distanceToCluster <= clusterRadius)
                                    {
                                        isInsideClusterSleeve = true;
                                        DebugLogger.Log($"[SUPPRESSION_DEBUG] *** SUPPRESSED *** Duct ID={duct.Id.IntegerValue}: placement point within {UnitUtils.ConvertFromInternalUnits(clusterRadius, UnitTypeId.Millimeters):F0}mm of cluster sleeve, skipping individual sleeve placement");
                                        break;
                                    }
                                }
                            }
                            
                            DebugLogger.Log($"[SUPPRESSION_DEBUG]   Final result - isInsideClusterSleeve: {isInsideClusterSleeve}");
                            
                            if (hasExistingSleeveAtPlacement || isInsideClusterSleeve)
                            {
                                if (hasExistingSleeveAtPlacement)
                                {
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG] *** SUPPRESSED *** Duct ID={duct.Id.IntegerValue}: existing sleeve found within {UnitUtils.ConvertFromInternalUnits(sleeveCheckRadius, UnitTypeId.Millimeters):F0}mm at placement point, skipping X-orientation placement");
                                }
                                skippedExistingCount++;
                                placed = true; // Mark as placed to avoid counting as missing
                                continue;
                            }
                            
                            DebugLogger.Log($"[SUPPRESSION_DEBUG] *** PROCEEDING *** Duct ID={duct.Id.IntegerValue}: No suppression conditions met, proceeding with sleeve placement");

                              DebugLogger.Log($"[DuctSleeveCommand] X-duct: Placing sleeve at {sleevePointEntryExit} in wall {bestWall.Id.IntegerValue}");
                            DebugLogger.Log($"[DuctSleeveCommand] About to call placer.PlaceDuctSleeve() with sw={sw}, sh={sh}");
                            string uniqueKey = $"{duct.Id.IntegerValue}_{bestWall.Id.IntegerValue}";                            

                            // Using simple duplicate prevention - the FINAL placer handles this internally
                            // if (!DuctSleevePlacer.HasPlacedSleeve(uniqueKey))
                            {
                                try
                                {
                                    placer.PlaceDuctSleeve(duct, sleevePointEntryExit, sw, sh, rayDirection, ductSymbol, bestWall, bestFaceNormal);
                                    DebugLogger.Log($"[DuctSleeveCommand] Successfully called placer.PlaceDuctSleeve()");
                                    processedDucts.Add(duct.Id); // Mark as processed to prevent duplicate placement
                                    placedCount++; // Only increment when actual placement occurs
                                }
                                catch (Exception placerEx)
                                {
                                    DebugLogger.Log($"[DuctSleeveCommand] Exception in placer.PlaceDuctSleeve(): {placerEx.Message}");
                                    DebugLogger.Log($"[DuctSleeveCommand] Placer exception stack: {placerEx.StackTrace}");
                                }
                                // DuctSleevePlacer.MarkSleevePlaced(uniqueKey);
                            }
                            placed = true;
                        }
                        if (!placed) { missingCount++; }
                        continue;
                    }
                    else
                    {
                        // --- BEGIN COMBINED LOGIC FOR Y-DUCTS: ReferenceIntersector + robust intersection/endpoint check ---
                        refWithContext = null;
                        DebugLogger.Log($"[DuctSleeveCommand] Y-duct {duct.Id.IntegerValue}: direction={rayDirection}, Start={line.GetEndPoint(0)}, End={line.GetEndPoint(1)}");
                        
                        // Try multiple sample points and directions for Y-ducts too
                        foreach (double t in new[] { 0.0, 0.25, 0.5, 0.75, 1.0 })
                        {
                            var samplePt = line.Evaluate(t, true);
                            // Try both forward and backward directions
                            var testDirections = new[] { rayDirection, rayDirection.Negate() };
                            foreach (var testDir in testDirections)
                            {
                                var hits = refIntersector.Find(samplePt, testDir);
                                if (hits?.Count > 0)
                                {
                                    refWithContext = hits;
                                    DebugLogger.Log($"[DuctSleeveCommand] Y-duct found wall hits at t={t} with direction {testDir}: {hits.Count} hits");
                                    break;
                                }
                            }
                            if (refWithContext?.Count > 0) break;
                        }
                        if (refWithContext == null || refWithContext.Count == 0)
                        {
                            missingCount++;
                            continue;
                        }
                        
                        // For each candidate wall, apply robust intersection/endpoint logic
                        bool placedY = false;
                        foreach (var wallHit in refWithContext)
                        {
                            var r = wallHit.GetReference();
                            var linkInst = doc.GetElement(r.ElementId) as RevitLinkInstance;
                            var targetDoc = linkInst != null ? linkInst.GetLinkDocument() : doc;
                            var eid = linkInst != null ? r.LinkedElementId : r.ElementId;
                            var hostWall = targetDoc?.GetElement(eid) as Wall;
                            if (hostWall == null) continue;
                              var ductLine = locCurve;
                            
                            // Validate duct line length to prevent short curve errors
                            double ductLength = ductLine.Length;
                            if (ductLength < 0.01) // Less than 0.01 feet (about 3mm)
                            {
                                DebugLogger.Log($"[DuctSleeveCommand] Skipping very short Y-duct {duct.Id.IntegerValue} with length {ductLength:F6}");
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
                                SetComparisonResult res = face.Intersect(ductLine, out ira);
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
                            bool startInside = IsPointInsideSolid(wallSolid, ductLine.GetEndPoint(0), hostWall.Orientation);
                            bool endInside = IsPointInsideSolid(wallSolid, ductLine.GetEndPoint(1), hostWall.Orientation);
                            
                            // Apply enhanced intersection detection for Y-ducts too
                            if (intersectionPoints.Count == 0 && !startInside && !endInside)
                            {                                // Try segmented approach for Y-ducts
                                List<XYZ> altIntersections = new List<XYZ>();
                                int segments = 10;
                                for (int i = 0; i < segments; i++)
                                {
                                    double t1 = (double)i / segments;
                                    double t2 = (double)(i + 1) / segments;
                                    XYZ pt1 = ductLine.Evaluate(t1, true);
                                    XYZ pt2 = ductLine.Evaluate(t2, true);
                                    
                                    // Check if segment is long enough to avoid Revit tolerance error
                                    double segmentLength = pt1.DistanceTo(pt2);
                                    if (segmentLength < 0.01) // Increased to 0.01 feet (about 3mm) - more conservative
                                    {
                                        DebugLogger.Log($"[DEBUG] Skipping Y-segment {i} with length {segmentLength:F9} (too short)");
                                        continue;
                                    }
                                    
                                    Line segment;
                                    try
                                    {
                                        segment = Line.CreateBound(pt1, pt2);
                                    }
                                    catch (Exception segmentEx)
                                    {
                                        DebugLogger.Log($"[DEBUG] Failed to create Y-segment {i}: {segmentEx.Message}, pt1={pt1}, pt2={pt2}, distance={segmentLength:F9}");
                                        continue;
                                    }
                                    
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
                                
                                if (altIntersections.Count == 0)
                                {
                                    // Try bounds-based approach
                                    BoundingBoxXYZ wallBounds = wallSolid.GetBoundingBox();
                                    if (wallBounds != null)
                                    {
                                        double tolerance = UnitUtils.ConvertToInternalUnits(10.0, UnitTypeId.Millimeters);
                                        XYZ min = wallBounds.Min - new XYZ(tolerance, tolerance, tolerance);
                                        XYZ max = wallBounds.Max + new XYZ(tolerance, tolerance, tolerance);
                                        
                                        XYZ ductStart = ductLine.GetEndPoint(0);
                                        XYZ ductEnd = ductLine.GetEndPoint(1);
                                        
                                        if (IsPointInBounds(ductStart, min, max) != IsPointInBounds(ductEnd, min, max))
                                        {
                                            var wallLocCurveAlt = hostWall.Location as LocationCurve;
                                            if (wallLocCurveAlt != null && wallLocCurveAlt.Curve != null)
                                            {
                                                XYZ wallCenterAlt2 = wallLocCurveAlt.Curve.Evaluate(0.5, true);
                                                XYZ wallNormal = hostWall.Orientation.Normalize();
                                                XYZ ductDirection = (ductEnd - ductStart).Normalize();
                                                double denominator = ductDirection.DotProduct(wallNormal);
                                                
                                                if (Math.Abs(denominator) > 1e-6)
                                                {
                                                    double t = (wallCenterAlt2 - ductStart).DotProduct(wallNormal) / denominator;
                                                    if (t >= 0 && t <= 1)
                                                    {
                                                        XYZ artificialIntersection = ductStart + ductDirection.Multiply(t * ductStart.DistanceTo(ductEnd));
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
                                intersectionPoints = intersectionPoints.OrderBy(pt => (pt - ductLine.GetEndPoint(0)).GetLength()).ToList();
                                ptEntry = intersectionPoints.First();
                                ptExit = intersectionPoints.Last();
                            }
                            else if (intersectionPoints.Count == 1)
                            {
                                ptEntry = intersectionPoints[0];
                                ptExit = startInside ? ductLine.GetEndPoint(0) : ductLine.GetEndPoint(1);
                            }
                            else if (startInside || endInside)
                            {
                                ptEntry = startInside ? ductLine.GetEndPoint(0) : ductLine.GetEndPoint(1);
                                ptExit = ptEntry;
                            }
                            
                            double segmentLen = ptEntry.DistanceTo(ptExit);
                            if (segmentLen < UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters)) continue;
                            
                            // Center using entry/exit or offset to wall center
                            var wallLocCurve = hostWall.Location as LocationCurve;
                            XYZ wallCenter = null;
                            if (wallLocCurve != null && wallLocCurve.Curve != null)
                                wallCenter = wallLocCurve.Curve.Evaluate(0.5, true);
                            
                            XYZ sleevePoint;
                            // --- SLEEVE PLACEMENT POINT LOGIC (applies to both X- and Y-ducts, per wall candidate) ---
                            if (ptEntry != null && ptExit != null && ptEntry.DistanceTo(ptExit) > UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters)) {
                                sleevePoint = (ptEntry + ptExit) * 0.5;
                                DebugLogger.Log($"[DuctSleeveCommand] Using midpoint for sleeve placement: {sleevePoint}");
                            } else if (ptEntry != null) {
                                // Project entry to wall center along wall orientation
                                var wallLocCurveInner = hostWall.Location as LocationCurve;
                                XYZ wallCenterInner = null;
                                XYZ wallNormal = hostWall.Orientation.Normalize();
                                double wallThickness = hostWall.Width;
                                if (wallLocCurveInner != null && wallLocCurveInner.Curve != null)
                                    wallCenterInner = wallLocCurveInner.Curve.Evaluate(0.5, true);
                                if (wallCenterInner != null) {
                                    double distToCenterInner = (ptEntry - wallCenterInner).DotProduct(wallNormal);
                                    sleevePoint = ptEntry - wallNormal.Multiply(distToCenterInner);
                                    DebugLogger.Log($"[DuctSleeveCommand] Projected entry to wall center: {sleevePoint}");
                                } else {
                                    sleevePoint = ptEntry;
                                    DebugLogger.Log($"[DuctSleeveCommand] Using entry point for sleeve placement: {sleevePoint}");
                                }
                            } else {
                                DebugLogger.Log($"[DuctSleeveCommand] No valid intersection for sleeve placement");
                                continue;
                            }
                            
                            // Check for dampers near the intersection point
                            double damperTol = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters); // 1mm precision tolerance
                            bool hasDamper = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_DuctAccessory)
                                .OfClass(typeof(FamilyInstance))
                                .Cast<FamilyInstance>()
                                .Any(fi => fi.Symbol.Family.Name.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0
                                           && fi.GetTransform().Origin.DistanceTo(sleevePoint) < damperTol);
                            
                            if (hasDamper)
                            {
                                DebugLogger.Log($"Duct ID={duct.Id.IntegerValue}: damper present near intersection, skipping sleeve");
                                damperSkippedCount++;
                                continue;
                            }
                            
                            // Compute duct dimensions
                            double w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0;
                            double h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)?.AsDouble() ?? 0;
                            double clearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                            double sw = w + clearance * 2;
                            double sh = h + clearance * 2;

                            // ENHANCED SUPPRESSION CHECK: Skip if sleeve already exists OR if placement point is inside a cluster sleeve
                            double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters); // 1mm precision tolerance
                            bool hasExistingSleeveAtPlacement = existingSleeveLocations.Values.Any(sleeveLocation => 
                                sleevePoint.DistanceTo(sleeveLocation) <= sleeveCheckRadius);
                            
                            DebugLogger.Log($"[SUPPRESSION_DEBUG] Duct ID={duct.Id.IntegerValue} - Y-orientation suppression check:");
                            DebugLogger.Log($"[SUPPRESSION_DEBUG]   Placement point: {sleevePoint}");
                            DebugLogger.Log($"[SUPPRESSION_DEBUG]   Has existing sleeve at placement: {hasExistingSleeveAtPlacement}");
                            
                            // Check if placement point is inside any cluster sleeve's bounding box
                            bool isInsideClusterSleeve = false;
                            int clusterIndex = 0;
                            foreach (var clusterInfo in existingClusterLocationsWithBounds.Values)
                            {
                                clusterIndex++;
                                DebugLogger.Log($"[SUPPRESSION_DEBUG]   Checking cluster sleeve #{clusterIndex} for Y-duct:");
                                DebugLogger.Log($"[SUPPRESSION_DEBUG]     Cluster location: {clusterInfo.Location}");
                                
                                if (clusterInfo.BoundingBox != null)
                                {
                                    // Expand cluster bounding box slightly for tolerance
                                    double clusterTolerance = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm tolerance - MATCHES cluster command tolerance
                                    XYZ minExpanded = clusterInfo.BoundingBox.Min - new XYZ(clusterTolerance, clusterTolerance, clusterTolerance);
                                    XYZ maxExpanded = clusterInfo.BoundingBox.Max + new XYZ(clusterTolerance, clusterTolerance, clusterTolerance);
                                    
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Expanded BoundingBox: Min={minExpanded}, Max={maxExpanded}");
                                    
                                    bool pointInBounds = IsPointInBounds(sleevePoint, minExpanded, maxExpanded);
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Point in expanded bounds: {pointInBounds}");
                                    
                                    if (pointInBounds)
                                    {
                                        isInsideClusterSleeve = true;
                                        DebugLogger.Log($"[SUPPRESSION_DEBUG] *** SUPPRESSED *** Y-duct ID={duct.Id.IntegerValue}: placement point inside cluster sleeve bounding box, skipping individual sleeve placement");
                                        break;
                                    }
                                }
                                else
                                {
                                    // Fallback: use distance check if bounding box is not available
                                    double distanceToCluster = sleevePoint.DistanceTo(clusterInfo.Location);
                                    double clusterRadius = UnitUtils.ConvertToInternalUnits(200.0, UnitTypeId.Millimeters); // 200mm radius
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG]     Distance to cluster: {UnitUtils.ConvertFromInternalUnits(distanceToCluster, UnitTypeId.Millimeters):F1}mm");
                                    
                                    if (distanceToCluster <= clusterRadius)
                                    {
                                        isInsideClusterSleeve = true;
                                        DebugLogger.Log($"[SUPPRESSION_DEBUG] *** SUPPRESSED *** Y-duct ID={duct.Id.IntegerValue}: placement point within {UnitUtils.ConvertFromInternalUnits(clusterRadius, UnitTypeId.Millimeters):F0}mm of cluster sleeve, skipping individual sleeve placement");
                                        break;
                                    }
                                }
                            }
                            
                            DebugLogger.Log($"[SUPPRESSION_DEBUG]   Y-duct final result - isInsideClusterSleeve: {isInsideClusterSleeve}");
                            
                            if (hasExistingSleeveAtPlacement || isInsideClusterSleeve)
                            {
                                if (hasExistingSleeveAtPlacement)
                                {
                                    DebugLogger.Log($"[SUPPRESSION_DEBUG] *** SUPPRESSED *** Y-duct ID={duct.Id.IntegerValue}: existing sleeve found within {UnitUtils.ConvertFromInternalUnits(sleeveCheckRadius, UnitTypeId.Millimeters):F0}mm at placement point, skipping Y-orientation placement");
                                }
                                skippedExistingCount++;
                                placedY = true; // Mark as placed to avoid counting as missing
                                break; // Exit the wall loop for this duct
                            }
                            
                            DebugLogger.Log($"[SUPPRESSION_DEBUG] *** PROCEEDING *** Y-duct ID={duct.Id.IntegerValue}: No suppression conditions met, proceeding with sleeve placement");
                            
                            DebugLogger.Log($"[DuctSleeveCommand] Y-duct: Placing sleeve at {sleevePoint} in wall {hostWall.Id.IntegerValue}");
                            placer.PlaceDuctSleeve(duct, sleevePoint, sw, sh, rayDirection, ductSymbol, hostWall, hostWall.Orientation.Normalize());
                            processedDucts.Add(duct.Id); // Mark as processed to prevent duplicate placement
                            placedCount++; // Only increment when actual placement occurs
                            placedY = true;
                            break; // Only place one sleeve per Y-duct
                        }
                        if (!placedY) missingCount++;
                        continue;
                        // --- END COMBINED LOGIC FOR Y-DUCTS ---
                    }
                }
                tx.Commit();
                DebugLogger.Log($"DuctSleeveCommand summary: Total={totalDucts}, Intersections={intersectionCount}, Placed={placedCount}, Missing={missingCount}, DamperSkipped={damperSkippedCount}, SkippedExisting={skippedExistingCount}");
                // DuctSleevePlacer.LogDuctSleeveTable();  // Not needed for FINAL version
            }

            TaskDialog.Show("Done", "Duct sleeves placed.");
            return Result.Succeeded;        }

        // Helper: Robust point-in-solid test with tolerance and dual-direction ray casting
        private static bool IsPointInsideSolid(Solid solid, XYZ point, XYZ direction, double tolMm = 2.0)
        {
            // Validate direction vector to prevent issues - be more strict
            if (direction.GetLength() < 1e-9) // More strict threshold
            {
                DebugLogger.Log($"[IsPointInsideSolid] WARNING: Direction vector too small ({direction.GetLength():E6}), using default X direction");
                direction = new XYZ(1, 0, 0);
            }
            
            double tol = UnitUtils.ConvertToInternalUnits(tolMm, UnitTypeId.Millimeters);
            // 1. Accept if point is within tol of any face
            foreach (Face face in solid.Faces)
            {
                var result = face.Project(point);
                if (result != null && result.Distance <= tol)
                {
                    DebugLogger.Log($"[IsPointInsideSolid] Point near face: dist={result.Distance}, tol={tol}");
                    return true;
                }
            }
            // 2. Ray cast in both directions
            bool inside1 = RayParityTest(solid, point, direction);
            bool inside2 = RayParityTest(solid, point, direction.Negate());
            DebugLogger.Log($"[IsPointInsideSolid] Point=({point.X},{point.Y},{point.Z}), Dir1=({direction.X},{direction.Y},{direction.Z}), Dir2=({-direction.X},{-direction.Y},{-direction.Z}), Inside1={inside1}, Inside2={inside2}");
            return inside1 || inside2;
        }        private static bool RayParityTest(Solid solid, XYZ point, XYZ direction)
        {
            // Validate direction vector to prevent zero-length curves
            if (direction.GetLength() < 1e-9) // More strict threshold
            {
                DebugLogger.Log($"[RayParityTest] WARNING: Direction vector too small ({direction.GetLength():E6}), using default X direction");
                direction = new XYZ(1, 0, 0);
            }
            
            // Normalize the direction safely
            XYZ normalizedDirection;
            try
            {
                normalizedDirection = direction.Normalize();
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[RayParityTest] ERROR: Failed to normalize direction ({direction.X}, {direction.Y}, {direction.Z}): {ex.Message}, using default X direction");
                normalizedDirection = new XYZ(1, 0, 0);
            }
            
            // Create ray endpoint with safe distance
            XYZ rayEndPoint = point + normalizedDirection.Multiply(1000);
            
            // Critical validation: Check the actual distance between points that will be used for Line.CreateBound
            double rayLength = point.DistanceTo(rayEndPoint);
            if (rayLength < 0.01) // 0.01 feet ≈ 3mm - Revit's minimum tolerance
            {
                DebugLogger.Log($"[RayParityTest] WARNING: Ray too short ({rayLength:F9}), recalculating with guaranteed length");
                // Force a minimum ray length by using a unit vector and scaling appropriately
                rayEndPoint = point + new XYZ(1, 0, 0).Multiply(10); // 10 feet guaranteed length
                rayLength = point.DistanceTo(rayEndPoint);
                DebugLogger.Log($"[RayParityTest] Forced ray length: {rayLength:F6}");
            }
              Line ray;
            try
            {
                ray = Line.CreateBound(point, rayEndPoint);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[RayParityTest] ERROR: Failed to create ray line from ({point.X:F6}, {point.Y:F6}, {point.Z:F6}) to ({rayEndPoint.X:F6}, {rayEndPoint.Y:F6}, {rayEndPoint.Z:F6}), distance={rayLength:F9}: {ex.Message}");
                return false; // Default to outside if we can't create the ray
            }
            
            int intersectionCount = 0;
            foreach (Face face in solid.Faces)
            {
                IntersectionResultArray results = null;
                SetComparisonResult res = face.Intersect(ray, out results);
                if (res == SetComparisonResult.Overlap && results != null)
                {
                    intersectionCount += results.Size;
                }
            }
            bool isInside = (intersectionCount % 2) == 1;
            DebugLogger.Log($"[RayParityTest] Point=({point.X},{point.Y},{point.Z}), Dir=({direction.X},{direction.Y},{direction.Z}), Intersections={intersectionCount}, Inside={isInside}");
            return isInside;
        }

        // Helper: Check if point is within bounding box
        private static bool IsPointInBounds(XYZ point, XYZ min, XYZ max)
        {
            return point.X >= min.X && point.X <= max.X &&
                   point.Y >= min.Y && point.Y <= max.Y &&
                   point.Z >= min.Z && point.Z <= max.Z;
        }

        // Helper function for bounding box check
        private static bool IsPointInsideBoundingBox(XYZ point, BoundingBoxXYZ bb)
        {
            if (bb == null) return false;
            return point.X >= bb.Min.X && point.X <= bb.Max.X &&
                   point.Y >= bb.Min.Y && point.Y <= bb.Max.Y &&
                   point.Z >= bb.Min.Z && point.Z <= bb.Max.Z;
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
