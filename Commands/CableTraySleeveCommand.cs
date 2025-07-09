using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CableTraySleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DebugLogger.Log("Starting CableTraySleeveCommand"); // log start
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Select default CT# sleeve symbol
            var ctSymbol = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym =>
                    string.Equals(sym.Family.Name, "CableTrayOpeningOnWall", System.StringComparison.OrdinalIgnoreCase)
                    && sym.Name.StartsWith("CT#"));
            // Find cluster symbols for suppression (CableTrayOpeningOnWall and CableTrayOpeningOnSlab)
            var clusterSymbolWall = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family != null && sym.Family.Name.Equals("CableTrayOpeningOnWall", System.StringComparison.OrdinalIgnoreCase));
            var clusterSymbolSlab = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(sym => sym.Family != null && sym.Family.Name.Equals("CableTrayOpeningOnSlab", System.StringComparison.OrdinalIgnoreCase));
            if (ctSymbol == null)
            {
                TaskDialog.Show("Error", "No CT# family symbol found for CableTray sleeves.");
                return Result.Failed;
            }

            // Activate Cable Tray Symbol
            using (var txActivate = new Transaction(doc, "Activate CT Symbol"))
            {
                txActivate.Start();
                if (!ctSymbol.IsActive)
                    ctSymbol.Activate();
                txActivate.Commit();
            }

            // Find a non-template 3D view
            var view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D))
                .Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate);
            if (view3D == null)
            {
                TaskDialog.Show("Error", "No non-template 3D view available.");
                return Result.Failed;
            }

            // Prepare wall filter
            ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            // Use Element target and include linked documents for wall intersections
            var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Element, view3D)
            {
                FindReferencesInRevitLinks = true
            };

            // Place sleeves for each cable tray
            var placer = new CableTraySleevePlacer(doc);
            var trays = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_CableTray)
                .OfClass(typeof(CableTray))
                .Cast<CableTray>();

            // SUPPRESSION: Collect existing cable tray sleeves to avoid duplicates
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("CT#"))
                .ToList();

            // Also collect cable tray cluster sleeves to prevent individual sleeves from being placed inside them
            var existingClusterSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Name != null && 
                           (string.Equals(fi.Symbol.Family.Name, "CableTrayOpeningOnWall", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(fi.Symbol.Family.Name, "CableTrayOpeningOnSlab", StringComparison.OrdinalIgnoreCase)))
                .ToList();

            DebugLogger.Log($"Found {existingSleeves.Count} existing cable tray sleeves and {existingClusterSleeves.Count} cable tray cluster sleeves in the model");

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
            int totalTrays = trays.Count();
            int intersectionCount = 0;
            int placedCount = 0;
            int missingCount = 0;
            int skippedExistingCount = 0; // Counter for trays with existing sleeves
            DebugLogger.Log($"Found {totalTrays} cable trays to process");

            // HashSet for duplicate suppression (robust, like pipes/ducts)
            HashSet<ElementId> processedCableTrays = new HashSet<ElementId>();

            using (var tx = new Transaction(doc, "Place Cable Tray Sleeves"))
            {
                DebugLogger.Log("Transaction started for cable tray sleeve placement");
                tx.Start();
                DebugLogger.Log("Starting wall intersection tests for cable trays");
                foreach (var tray in trays)
                {
                    if (processedCableTrays.Contains(tray.Id))
                    {
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: already processed, skipping to prevent duplicate sleeve");
                        continue;
                    }
                    DebugLogger.Log($"Processing CableTray ID={tray.Id.IntegerValue}");
                    var curve = (tray.Location as LocationCurve)?.Curve as Line;
                    if (curve == null) {
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: no valid location curve, skipping");
                        continue;
                    }

                    // Get tray dimensions for sleeve placement
                    double width = tray.LookupParameter("Width")?.AsDouble() ?? 0.0;
                    double height = tray.LookupParameter("Height")?.AsDouble() ?? 0.0;

                    // Enhanced multi-direction raycasting - cast from multiple points along cable tray to find actual wall intersection
                    XYZ rayDir = curve.Direction;

                    // Cast rays from multiple points along the cable tray (start, 25%, 50%, 75%, end)
                    var testPoints = new List<XYZ>
                     {
                         curve.GetEndPoint(0),           // Start point
                         curve.Evaluate(0.25, true),    // 25% along
                         curve.Evaluate(0.5, true),     // Midpoint  
                         curve.Evaluate(0.75, true),    // 75% along
                         curve.GetEndPoint(1)           // End point
                     };

                    var allHits = new List<(ReferenceWithContext hit, XYZ direction, XYZ rayOrigin)>();

                    foreach (var testPoint in testPoints)
                    {
                        // Cast in primary direction (forward/backward)
                        var hitsFwd = refIntersector.Find(testPoint, rayDir)?.Where(h => h != null).OrderBy(h => h.Proximity).ToList();
                        var hitsBack = refIntersector.Find(testPoint, rayDir.Negate())?.Where(h => h != null).OrderBy(h => h.Proximity).ToList();

                        // Also cast in perpendicular directions to catch walls parallel to cable tray
                        var perpDir1 = new XYZ(-rayDir.Y, rayDir.X, 0).Normalize();
                        var perpDir2 = perpDir1.Negate();
                        var hitsPerp1 = refIntersector.Find(testPoint, perpDir1)?.Where(h => h != null).OrderBy(h => h.Proximity).ToList();
                        var hitsPerp2 = refIntersector.Find(testPoint, perpDir2)?.Where(h => h != null).OrderBy(h => h.Proximity).ToList();

                        // Add all hits with their ray origin point
                        if (hitsFwd?.Any() == true) allHits.AddRange(hitsFwd.Select(h => (h, rayDir, testPoint)));
                        if (hitsBack?.Any() == true) allHits.AddRange(hitsBack.Select(h => (h, rayDir.Negate(), testPoint)));
                        if (hitsPerp1?.Any() == true) allHits.AddRange(hitsPerp1.Select(h => (h, perpDir1, testPoint)));
                        if (hitsPerp2?.Any() == true) allHits.AddRange(hitsPerp2.Select(h => (h, perpDir2, testPoint)));
                    }

                    ReferenceWithContext bestHit = null;
                    XYZ bestDir = null;
                    XYZ bestRayOrigin = null;

                    if (allHits.Any())
                    {
                        // Find the hit closest to the cable tray path (prioritize hits from cable tray endpoints)
                        var closestHit = allHits.OrderBy(x => x.hit.Proximity).First();
                        bestHit = closestHit.hit;
                        bestDir = closestHit.direction;
                        bestRayOrigin = closestHit.rayOrigin;
                    }

                    int hitCount = allHits.Count;
                    DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: multi-point ray hits from {testPoints.Count} points, total={hitCount}");
                    DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: direction={rayDir}, Start={curve.GetEndPoint(0)}, End={curve.GetEndPoint(1)}");
                    if (bestHit == null)
                    {
                        missingCount++;
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: no wall intersection detected in either direction, skipping");
                        continue;
                    }
                    // Determine host wall
                    var r = bestHit.GetReference();
                    var linkInst = doc.GetElement(r.ElementId) as RevitLinkInstance;
                    var targetDoc = linkInst != null ? linkInst.GetLinkDocument() : doc;
                    ElementId elemId = linkInst != null ? r.LinkedElementId : r.ElementId;
                    Wall hostWall = targetDoc?.GetElement(elemId) as Wall;
                    if (hostWall == null)
                    {
                        missingCount++;
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: hit but no valid wall element, skipping");
                        continue;
                    }
                    // True intersection point (wall face) - use the actual intersection from the closest ray hit
                    XYZ intersectionPt = r.GlobalPoint;
                    DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: raw intersection at {intersectionPt} from ray origin {bestRayOrigin}");

                    // Calculate wall centerline placement point (same logic as CableTraySleevePlacer)
                    double wallThickness = hostWall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? hostWall.Width;

                    // Determine cable tray direction for wall normal calculation
                    XYZ cableTrayDirection = curve.Direction.Normalize();
                    bool isYAxisCableTray = Math.Abs(cableTrayDirection.Y) > Math.Abs(cableTrayDirection.X);

                    // *** CRITICAL: COPY EXACT PIPE LOGIC - USE INTERSECTION AS-IS FOR ALL CABLE TRAYS ***
                    // *** NO X-AXIS OR Y-AXIS SPECIAL HANDLING - PIPES WORK PERFECTLY THIS WAY ***
                    XYZ finalPlacementPoint = intersectionPt; // *** EXACT PIPE SLEEVE WORKING LOGIC ***
                    DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: using intersection point as-is (EXACT PIPE LOGIC) - {finalPlacementPoint}");

                    // SUPPRESSION CHECK: Skip if sleeve already exists at this placement point
                    double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm precision tolerance
                    bool hasExistingSleeveAtPlacement = existingSleeveLocations.Values.Any(sleeveLocation =>
                        finalPlacementPoint.DistanceTo(sleeveLocation) <= sleeveCheckRadius);

                    // Check if placement point is inside any cluster sleeve's bounding box
                    bool isInsideClusterSleeve = false;
                    foreach (var clusterInfo in existingClusterLocationsWithBounds.Values)
                    {
                        if (clusterInfo.BoundingBox != null)
                        {
                            // Expand cluster bounding box slightly for tolerance
                            double clusterToleranceLocal = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm tolerance - MATCHES cluster command tolerance
                            XYZ minExpanded = clusterInfo.BoundingBox.Min - new XYZ(clusterToleranceLocal, clusterToleranceLocal, clusterToleranceLocal);
                            XYZ maxExpanded = clusterInfo.BoundingBox.Max + new XYZ(clusterToleranceLocal, clusterToleranceLocal, clusterToleranceLocal);
                            
                            if (IsPointInBounds(finalPlacementPoint, minExpanded, maxExpanded))
                            {
                                isInsideClusterSleeve = true;
                                DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: placement point inside cluster sleeve bounding box, skipping individual sleeve placement");
                                break;
                            }
                        }
                        else
                        {
                            // Fallback: use distance check if bounding box is not available
                            double distanceToCluster = finalPlacementPoint.DistanceTo(clusterInfo.Location);
                            double clusterRadius = UnitUtils.ConvertToInternalUnits(200.0, UnitTypeId.Millimeters); // 200mm radius
                            if (distanceToCluster <= clusterRadius)
                            {
                                isInsideClusterSleeve = true;
                                DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: placement point within {UnitUtils.ConvertFromInternalUnits(clusterRadius, UnitTypeId.Millimeters):F0}mm of cluster sleeve, skipping individual sleeve placement");
                                break;
                            }
                        }
                    }

                    if (hasExistingSleeveAtPlacement || isInsideClusterSleeve)
                    {
                        if (hasExistingSleeveAtPlacement)
                        {
                            DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: existing sleeve found within {UnitUtils.ConvertFromInternalUnits(sleeveCheckRadius, UnitTypeId.Millimeters):F0}mm at placement point, skipping");
                        }
                        skippedExistingCount++;
                        continue;
                    }
                    // --- CLUSTER SUPPRESSION: Use helper to check for any cluster at this location ---
                    double clusterTolerance = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                    bool clusterFound = false;
                    if (clusterSymbolWall != null && OpeningDuplicationChecker.IsClusterAtLocation(doc, finalPlacementPoint, clusterTolerance, clusterSymbolWall))
                        clusterFound = true;
                    if (!clusterFound && clusterSymbolSlab != null && OpeningDuplicationChecker.IsClusterAtLocation(doc, finalPlacementPoint, clusterTolerance, clusterSymbolSlab))
                        clusterFound = true;
                    if (clusterFound)
                    {
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Skipping sleeve placement because cluster sleeve found at location (helper check)");
                        skippedExistingCount++;
                        continue;
                    }
                    // Check again for duplicate before placement (paranoia)
                    if (processedCableTrays.Contains(tray.Id))
                    {
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: detected duplicate processing just before placement! This should not happen. Skipping.");
                        continue;
                    }

                    // Log sleeve dimensions
                    DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: sleeve width={width}, height={height}");
                    DebugLogger.Log($"Requesting sleeve placement at final placement point {finalPlacementPoint}");
                    XYZ direction = bestDir;
                    // Pass final placement point (wall centerline corrected for X-axis, intersection for Y-axis)
                    placer.PlaceCableTraySleeve(tray, finalPlacementPoint, width, height, direction, ctSymbol, hostWall);
                    placedCount++; // Only increment when actual placement occurs
                    processedCableTrays.Add(tray.Id); // Mark as processed
                    DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: sleeve placed and marked as processed");
                    continue;
                }
                tx.Commit();
                DebugLogger.Log("Transaction committed for cable tray sleeve placement");
                // Summary log
                DebugLogger.Log($"CableTraySleeveCommand summary: Total={totalTrays}, Intersections={intersectionCount}, Placed={placedCount}, Missing={missingCount}, SkippedExisting={skippedExistingCount}");
            }

            DebugLogger.Log("Cable tray sleeves placement completed.");
            return Result.Succeeded;
        }

        private XYZ GetWallNormal(Wall wall, XYZ point)
        {
            // Used for orientation only, not for hosting. Do not use wall as host.
            try
            {
                // Get the wall's location curve
                LocationCurve locationCurve = wall.Location as LocationCurve;
                if (locationCurve?.Curve is Line line)
                {
                    XYZ direction = line.Direction.Normalize();
                    XYZ normal = new XYZ(-direction.Y, direction.X, 0).Normalize();
                    return normal;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CableTraySleeveCommand] Failed to get wall normal: {ex.Message}");
            }

            // Default to X-axis normal
            return new XYZ(1, 0, 0);
        }
        
        private static bool IsPointInBounds(XYZ pt, XYZ min, XYZ max)
        {
            return pt.X >= min.X && pt.X <= max.X &&
                   pt.Y >= min.Y && pt.Y <= max.Y &&
                   pt.Z >= min.Z && pt.Z <= max.Z;
        }
    }
}
