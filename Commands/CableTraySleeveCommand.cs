using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class CableTraySleeveCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // ...existing code...

            // Log available families for debugging
            var allFamilySymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family != null)
                .ToList();
            
            StructuralElementLogger.LogStructuralElement("DIAGNOSTIC", 0, "CT_FAMILY_SEARCH", $"Found {allFamilySymbols.Count} family symbols in project");
            
            var cableTrayFamilies = allFamilySymbols
                .Where(sym => sym.Family.Name.IndexOf("CableTray", System.StringComparison.OrdinalIgnoreCase) >= 0 ||
                              sym.Family.Name.IndexOf("CT", System.StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
            
            foreach (var fam in cableTrayFamilies.Take(10)) // Log first 10 cable tray families
            {
                StructuralElementLogger.LogStructuralElement("DIAGNOSTIC", (int)fam.Id.Value, "CT_FAMILY_FOUND", $"Family: {fam.Family.Name}, Symbol: {fam.Name}");
            }


            // Select all wall and slab family symbols for cable tray sleeves by family name only
            var ctWallSymbols = allFamilySymbols
                .Where(sym => string.Equals(sym.Family.Name, "CableTrayOpeningOnWall", System.StringComparison.OrdinalIgnoreCase))
                .ToList();
            var ctSlabSymbols = allFamilySymbols
                .Where(sym => string.Equals(sym.Family.Name, "CableTrayOpeningOnSlab", System.StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (!ctWallSymbols.Any() && !ctSlabSymbols.Any())
            {
                StructuralElementLogger.LogStructuralElement("ERROR", 0, "MISSING_CT_FAMILY", "Could not find cable tray sleeve family (CableTrayOpeningOnWall or CableTrayOpeningOnSlab)");
                TaskDialog.Show("Error", "Please load cable tray sleeve opening families (wall and slab).");
                return Result.Failed;
            }
            foreach (var sym in ctWallSymbols)
                StructuralElementLogger.LogStructuralElement("SUCCESS", (int)sym.Id.Value, "CT_FAMILY_FOUND", $"Using cable tray wall sleeve family: {sym.Family.Name}, Symbol: {sym.Name}");
            foreach (var sym in ctSlabSymbols)
                StructuralElementLogger.LogStructuralElement("SUCCESS", (int)sym.Id.Value, "CT_FAMILY_FOUND", $"Using cable tray slab sleeve family: {sym.Family.Name}, Symbol: {sym.Name}");

            // Activate all symbols if needed
            using (var txActivate = new Transaction(doc, "Activate CT Symbols"))
            {
                txActivate.Start();
                foreach (var sym in ctWallSymbols)
                    if (!sym.IsActive) sym.Activate();
                foreach (var sym in ctSlabSymbols)
                    if (!sym.IsActive) sym.Activate();
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

            // Prepare wall filter only (keep existing wall detection working)
            ElementFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            var refIntersector = new ReferenceIntersector(wallFilter, FindReferenceTarget.Element, view3D)
            {
                FindReferencesInRevitLinks = true
            };
            
            // Collect cable trays from both host and visible linked models
            var mepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
            var trays = mepElements
                .Where(tuple => tuple.Item1 is CableTray)
                .Select(tuple => tuple.Item1 as CableTray)
                .Where(t => t != null)
                .ToList();
            if (trays.Count == 0)
            {
                TaskDialog.Show("Info", "No cable trays found in host or linked models.");
                return Result.Succeeded;
            }

            // DEBUG: Let's also test if walls would be detected using direct solid intersection like structural elements
            StructuralElementLogger.LogStructuralElement("DEBUG_WALLS", 0, "WALL_SOLID_TEST", "Testing wall detection using direct solid approach like structural elements");
            
            // Collect walls from linked documents for comparison
            var linkedWalls = new List<(Element, Transform)>();
            foreach (var linkInstance in new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().Where(link => link.GetLinkDocument() != null))
            {
                var linkDoc = linkInstance.GetLinkDocument();
                var linkTransform = linkInstance.GetTotalTransform();
                
                var walls = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .ToList();
                    
                linkedWalls.AddRange(walls.Select(wall => (wall as Element, linkTransform)));
                StructuralElementLogger.LogStructuralElement("DEBUG_WALLS", (int)linkInstance.Id.Value, "LINKED_WALLS_FOUND", $"Found {walls.Count} walls in {linkInstance.Name}");
            }
            
            // Test first cable tray against walls using our direct solid method
            var firstTray = trays.FirstOrDefault();
            if (firstTray != null && linkedWalls.Any())
            {
                var wallIntersections = FindDirectStructuralIntersections(firstTray, linkedWalls);
            StructuralElementLogger.LogStructuralElement("DEBUG_WALLS", (int)firstTray.Id.Value, "WALL_SOLID_RESULT", $"Direct solid method found {wallIntersections.Count} wall intersections");
            }

            // Collect structural elements for direct solid intersection (replace ReferenceIntersector approach)
            var directStructuralElements = CollectStructuralElementsForDirectIntersection(doc);

            // Place sleeves for each cable tray
            var placer = new CableTraySleevePlacer(doc);

            // SUPPRESSION: Collect existing cable tray sleeves to avoid duplicates
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("CT#"))
                .ToList();

            DebugLogger.Log($"Found {existingSleeves.Count} existing cable tray sleeves in the model");

            // Create a map of existing sleeve locations for quick lookup
            var existingSleeveLocations = existingSleeves.ToDictionary(
                sleeve => sleeve,
                sleeve => (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin
            );

            // Initialize counters for detailed logging
            int totalTrays = trays.Count();
            int intersectionCount = 0;
            int placedCount = 0;
            int missingCount = 0;
            int skippedExistingCount = 0; // Counter for trays with existing sleeves
            int structuralElementsDetected = 0; // Counter for structural elements
            int structuralSleevesPlacer = 0; // Counter for successful structural sleeve placements
            DebugLogger.Log($"Found {totalTrays} cable trays to process");
            StructuralElementLogger.LogStructuralElement("SYSTEM", 0, "PROCESSING STARTED", $"Total cable trays to process: {totalTrays}");

            // HashSet for duplicate suppression (robust, like pipes/ducts)
            HashSet<ElementId> processedCableTrays = new HashSet<ElementId>();

            using (var tx = new Transaction(doc, "Place Cable Tray Sleeves"))
            {
                DebugLogger.Log("Transaction started for cable tray sleeve placement");
                tx.Start();
                DebugLogger.Log("Starting wall intersection tests for cable trays");
                foreach (var tray in trays)
                {
                    if (tray == null)
                        continue;
                    if (processedCableTrays.Contains(tray.Id))
                    {
                        DebugLogger.Log($"CableTray ID={tray.Id.Value}: already processed, skipping to prevent duplicate sleeve");
                        continue;
                    }
                    DebugLogger.Log($"Processing CableTray ID={tray.Id.Value}");
                    var curve = (tray.Location as LocationCurve)?.Curve as Line;
                    if (curve == null) {
                        DebugLogger.Log($"CableTray ID={tray.Id.Value}: no valid location curve, skipping");
                        continue;
                    }

                    // Get tray dimensions for sleeve placement
                    double width = tray.LookupParameter("Width")?.AsDouble() ?? 0.0;
                    double height = tray.LookupParameter("Height")?.AsDouble() ?? 0.0;

                    // Enhanced multi-direction raycasting for WALLS ONLY (keep existing wall detection working)
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

                    var allWallHits = new List<(ReferenceWithContext hit, XYZ direction, XYZ rayOrigin)>();

                    // Test wall intersections using existing ReferenceIntersector (keep this working)
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
                        if (hitsFwd?.Any() == true) allWallHits.AddRange(hitsFwd.Select(h => (h, rayDir, testPoint)));
                        if (hitsBack?.Any() == true) allWallHits.AddRange(hitsBack.Select(h => (h, rayDir.Negate(), testPoint)));
                        if (hitsPerp1?.Any() == true) allWallHits.AddRange(hitsPerp1.Select(h => (h, perpDir1, testPoint)));
                        if (hitsPerp2?.Any() == true) allWallHits.AddRange(hitsPerp2.Select(h => (h, perpDir2, testPoint)));
                    }

                    // Also check for structural element intersections using direct solid approach
                    var structuralIntersections = FindDirectStructuralIntersections(tray, directStructuralElements);

                    // Process wall intersections (existing logic, keep working)
                    ReferenceWithContext? bestWallHit = null;
                    XYZ? bestWallDir = null;
                    XYZ? bestWallRayOrigin = null;

                    if (allWallHits.Any())
                    {
                        // Find the hit closest to the cable tray path (prioritize hits from cable tray endpoints)
                        var closestHit = allWallHits.OrderBy(x => x.hit.Proximity).First();
                        bestWallHit = closestHit.hit;
                        bestWallDir = closestHit.direction;
                        bestWallRayOrigin = closestHit.rayOrigin;
                    }

                    int wallHitCount = allWallHits.Count;
                    int structuralHitCount = structuralIntersections.Count;
                    DebugLogger.Log($"CableTray ID={tray.Id.Value}: Wall hits: {wallHitCount}, Structural hits: {structuralHitCount}");

                    // Process wall intersections first (preserve existing wall functionality)
                    if (bestWallHit != null)
                    {
                        var r = bestWallHit.GetReference();
                        var linkInst = doc.GetElement(r.ElementId) as RevitLinkInstance;
                        var targetDoc = linkInst != null ? linkInst.GetLinkDocument() : doc;
                        ElementId elemId = linkInst != null ? r.LinkedElementId : r.ElementId;
                        Element? linkedReferenceElement = targetDoc?.GetElement(elemId);

                        if (linkedReferenceElement != null)
                        {
                            string elementTypeName = linkedReferenceElement.Category?.Name ?? "NO_CATEGORY";
                            DebugLogger.Log($"CableTray ID={tray.Id.Value}: detected wall element: {elementTypeName}, ID={linkedReferenceElement.Id.Value}");
                            StructuralElementLogger.LogStructuralElement(elementTypeName, (int)linkedReferenceElement.Id.Value, "ELEMENT DETECTED", $"Hit by cable tray {tray.Id.Value}");

                            // Calculate intersection point for wall
                            XYZ intersectionPoint = bestWallRayOrigin + bestWallDir * bestWallHit.Proximity;
                            StructuralElementLogger.LogStructuralElement("CableTray-WALL INTERSECTION", 0, "INTERSECTION_DETAILS", $"CableTray ID={tray.Id.Value}, Wall ID={linkedReferenceElement.Id.Value}, Position=({intersectionPoint.X:F9}, {intersectionPoint.Y:F9}, {intersectionPoint.Z:F9})");

                            // Always use CableTrayOpeningOnWall family for wall/framing
                            var wallFamilySymbol = ctWallSymbols.FirstOrDefault();
                            if (wallFamilySymbol == null)
                            {
                                TaskDialog.Show("Error", "Cable tray wall sleeve family (CableTrayOpeningOnWall) is missing. Please load it and rerun the command.");
                                StructuralElementLogger.LogStructuralElement(elementTypeName, (int)linkedReferenceElement.Id.Value, "SLEEVE_FAILED", "Reason: Wall family symbol not found");
                                skippedExistingCount++;
                                processedCableTrays.Add(tray.Id);
                                continue;
                            }
                            DebugLogger.Log($"CableTray ID={tray.Id.Value}: Using wall family: {wallFamilySymbol.Family.Name}, Symbol: {wallFamilySymbol.Name}");
                            if (PlaceCableTraySleeveAtLocation_Wall(doc, wallFamilySymbol, linkedReferenceElement, intersectionPoint, bestWallDir ?? XYZ.BasisX, width, height, tray.Id))
                            {
                                placedCount++;
                                processedCableTrays.Add(tray.Id);
                                DebugLogger.Log($"CableTray ID={tray.Id.Value}: Wall sleeve successfully placed at {intersectionPoint}");
                            }
                            else
                            {
                                skippedExistingCount++;
                                processedCableTrays.Add(tray.Id); // Still mark as processed to avoid duplicates
                                StructuralElementLogger.LogStructuralElement(elementTypeName, (int)linkedReferenceElement.Id.Value, "SLEEVE_FAILED", "Reason: Existing sleeve found at location");
                            }
                        }
                    }

                    // Process structural intersections using new direct approach
                    foreach (var (structuralElement, intersectionPoint) in structuralIntersections)
                    {
                        if (processedCableTrays.Contains(tray.Id))
                        {
                            // Skip if already placed wall sleeve for this tray
                            continue;
                        }

                        structuralElementsDetected++;
                        string elementTypeName = structuralElement.Category?.Name ?? "STRUCTURAL";
                        DebugLogger.Log($"CableTray ID={tray.Id.Value}: detected structural element: {elementTypeName}, ID={structuralElement.Id.Value}");
                        StructuralElementLogger.LogStructuralElement(elementTypeName, (int)structuralElement.Id.Value, "STRUCTURAL DETECTED", $"Hit by cable tray {tray.Id.Value}");
                        StructuralElementLogger.LogStructuralElement("CableTray-STRUCTURAL INTERSECTION", 0, "INTERSECTION_DETAILS", $"CableTray ID={tray.Id.Value}, Structural ID={structuralElement.Id.Value}, Position=({intersectionPoint.X:F9}, {intersectionPoint.Y:F9}, {intersectionPoint.Z:F9})");

                        // For structural elements, calculate direction based on cable tray direction
                        XYZ sleeveDirection = rayDir;

                        // Decide which family to use based on host type: OnWall for structural framing, OnSlab for floors
                        FamilySymbol? familySymbolToUse = null;
                        string linkedReferenceType = "UNKNOWN";
                        if (structuralElement is Floor)
                        {
                            familySymbolToUse = ctSlabSymbols.FirstOrDefault();
                            linkedReferenceType = "FLOOR";
                        }
                        else if (structuralElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                        {
                            // Always use CableTrayOpeningOnWall family for structural framing
                            familySymbolToUse = ctWallSymbols.FirstOrDefault();
                            linkedReferenceType = "STRUCTURAL FRAMING (CABLETRAY FAMILY)";
                        }
                        else
                        {
                            // Fallback: use slab symbol if nothing else
                            familySymbolToUse = ctSlabSymbols.FirstOrDefault();
                            linkedReferenceType = structuralElement.Category?.Name ?? "UNKNOWN";
                        }
                        DebugLogger.Log($"CableTray ID={tray.Id.Value}: Linked reference type detected: {linkedReferenceType}");
                        if (familySymbolToUse == null)
                        {
                            TaskDialog.Show("Error", $"Cable tray sleeve family for {linkedReferenceType} is missing. Please load it and rerun the command.");
                            StructuralElementLogger.LogStructuralElement(elementTypeName, (int)structuralElement.Id.Value, "SLEEVE_FAILED", $"Reason: {linkedReferenceType} family symbol not found");
                            skippedExistingCount++;
                            processedCableTrays.Add(tray.Id);
                            continue;
                        }
                        DebugLogger.Log($"CableTray ID={tray.Id.Value}: Using family: {familySymbolToUse.Family.Name}, Symbol: {familySymbolToUse.Name} for linked reference type {linkedReferenceType}");
                        if (PlaceCableTraySleeveAtLocation_Structural(doc, familySymbolToUse, structuralElement, intersectionPoint, sleeveDirection, width, height, tray))
                        {
                            structuralSleevesPlacer++;
                            placedCount++;
                            processedCableTrays.Add(tray.Id);
                            DebugLogger.Log($"CableTray ID={(int)tray.Id.Value}: Structural sleeve successfully placed at {intersectionPoint}");
                            break; // Only place one sleeve per cable tray
                        }
                        else
                        {
                            skippedExistingCount++;
                            processedCableTrays.Add(tray.Id); // Still mark as processed to avoid duplicates
                            StructuralElementLogger.LogStructuralElement(elementTypeName, (int)structuralElement.Id.Value, "SLEEVE_FAILED", "Reason: Existing sleeve found at location or placement failed");
                        }
                    }

                    // If no wall or structural intersection found
                    if (bestWallHit == null && structuralIntersections.Count == 0)
                    {
                        missingCount++;
                        DebugLogger.Log($"CableTray ID={(int)tray.Id.Value}: no wall or structural intersection detected, skipping");
                        continue;
                    }
                }
                
                tx.Commit();
                DebugLogger.Log("Transaction committed for cable tray sleeve placement");
                
                // Summary log with structural elements details
                DebugLogger.Log($"CableTraySleeveCommand summary: Total={totalTrays}, Intersections={intersectionCount}, Placed={placedCount}, Missing={missingCount}, SkippedExisting={skippedExistingCount}, StructuralDetected={structuralElementsDetected}, StructuralSleeves={structuralSleevesPlacer}");
                
                // Log structural summary to dedicated logger
                StructuralElementLogger.LogSummary("CableTraySleeveCommand", totalTrays, structuralElementsDetected, structuralSleevesPlacer, (structuralElementsDetected - structuralSleevesPlacer));
                StructuralElementLogger.LogStructuralElement("SYSTEM", 0, "COMMAND COMPLETED", $"Structural sleeve placement finished. Log file: {StructuralElementLogger.GetLogFilePath()}");
            }

            DebugLogger.Log("Cable tray sleeves placement completed.");
            return Result.Succeeded;
        }

        /// <summary>
        /// Places a cable tray wall sleeve at the specified location with duplication checking
        /// </summary>
        private bool PlaceCableTraySleeveAtLocation_Wall(Document doc, FamilySymbol ctWallSymbol, Element hostElement, XYZ placementPoint, XYZ direction, double width, double height, ElementId trayId)
        {
            try
            {
                double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                double clusterExpansion = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);

                // --- Robust duplicate suppression and logging (like DuctSleeveCommand) ---
                string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placementPoint, sleeveCheckRadius);
                DebugLogger.Log($"[CableTraySleeveCommand] DUPLICATION CHECK at {placementPoint}:\n{sleeveSummary}");
                var indivDup = OpeningDuplicationChecker.FindIndividualSleevesAtLocation(doc, placementPoint, sleeveCheckRadius);
                var clusterDup = OpeningDuplicationChecker.FindAllClusterSleevesAtLocation(doc, placementPoint, sleeveCheckRadius);
                if (indivDup.Any() || clusterDup.Any())
                {
                    string msg = $"SKIP: CableTray {trayId.Value} duplicate sleeve (individual or cluster) exists near {placementPoint}";
                    DebugLogger.Log($"[CableTraySleeveCommand] {msg}");
                    return false;
                }

                var placer = new CableTraySleevePlacer(doc);
                if (hostElement is Wall wall)
                {
                    double wallDepth = 0.0;
                    var wallType = wall.WallType;
                    var bParam = wallType.LookupParameter("b");
                    if (bParam != null && bParam.StorageType == StorageType.Double)
                    {
                        wallDepth = bParam.AsDouble();
                    }
                    else
                    {
                        DebugLogger.Log($"CableTray ID={(int)trayId.Value}: Wall type parameter 'b' not found or not double, using default depth");
                    }

                    // Bounding box extent log for cable tray sleeve (wall)
                    double minX = placementPoint.X - width / 2.0;
                    double maxX = placementPoint.X + width / 2.0;
                    double minY = placementPoint.Y - height / 2.0;
                    double maxY = placementPoint.Y + height / 2.0;
                    DebugLogger.Log($"CableTray ID={(int)trayId.Value}: [BBOX-DEBUG] Placement at ({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3}), BBox X=({UnitUtils.ConvertFromInternalUnits(minX, UnitTypeId.Millimeters):F1}, {UnitUtils.ConvertFromInternalUnits(maxX, UnitTypeId.Millimeters):F1}), Y=({UnitUtils.ConvertFromInternalUnits(minY, UnitTypeId.Millimeters):F1}, {UnitUtils.ConvertFromInternalUnits(maxY, UnitTypeId.Millimeters):F1}), Width={UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm");

                    FamilyInstance? sleeveInstance = placer.PlaceCableTraySleeve((CableTray)null!, placementPoint, width, height, direction, ctWallSymbol, wall);
                    if (sleeveInstance != null)
                    {
                        // --- Set Depth parameter as before ---
                        var depthParam = sleeveInstance.LookupParameter("Depth") ?? sleeveInstance.LookupParameter("d");
                        if (depthParam != null && depthParam.StorageType == StorageType.Double && !depthParam.IsReadOnly)
                        {
                            depthParam.Set(wallDepth);
                            DebugLogger.Log($"CableTray ID={(int)trayId.Value}: Set sleeve depth to wall thickness: {UnitUtils.ConvertFromInternalUnits(wallDepth, UnitTypeId.Millimeters):F0}mm");
                        }
                        else
                        {
                            DebugLogger.Log($"CableTray ID={(int)trayId.Value}: Sleeve depth parameter not found or not writable");
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"CableTray ID={(int)trayId.Value}: error placing wall sleeve: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Places a cable tray structural sleeve at the specified location with duplication checking
        /// </summary>
          private bool PlaceCableTraySleeveAtLocation_Structural(
            Document doc,
            FamilySymbol ctSlabSymbol,
            Element hostElement,
            XYZ placementPoint,
            XYZ direction,
            double width,
            double height,
            CableTray tray)
        {
            try
            {
                double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                // --- Robust duplicate suppression and logging (like DuctSleeveCommand) ---
                string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placementPoint, sleeveCheckRadius);
                DebugLogger.Log($"[CableTraySleeveCommand] DUPLICATION CHECK at {placementPoint}:\n{sleeveSummary}");
                var indivDup = OpeningDuplicationChecker.FindIndividualSleevesAtLocation(doc, placementPoint, sleeveCheckRadius);
                var clusterDup = OpeningDuplicationChecker.FindAllClusterSleevesAtLocation(doc, placementPoint, sleeveCheckRadius);
                if (indivDup.Any() || clusterDup.Any())
                {
                    string msg = $"SKIP: CableTray {tray.Id.Value} duplicate sleeve (individual or cluster) exists near {placementPoint}";
                    DebugLogger.Log($"[CableTraySleeveCommand] {msg}");
                    return false;
                }

                var placer = new CableTraySleevePlacer(doc);
                var sleeveInstance = placer.PlaceCableTraySleeve(
                    tray, // pass the actual tray object!
                    placementPoint,
                    width,
                    height,
                    direction,
                    ctSlabSymbol,
                    hostElement);
                if (sleeveInstance != null)
                {
                    DebugLogger.Log($"CableTray ID={(int)tray.Id.Value}: Structural sleeve successfully placed at {placementPoint}");
                    return true;
                }
                else
                {
                    DebugLogger.Log($"CableTray ID={(int)tray.Id.Value}: Failed to place structural sleeve at {placementPoint}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"CableTray ID={(int)tray.Id.Value}: error placing structural sleeve: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Collects structural elements from both active and linked documents for direct solid intersection
        /// Returns tuples of (element, linkTransform) where linkTransform is null for active document elements
        /// </summary>
        private List<(Element element, Transform linkTransform)> CollectStructuralElementsForDirectIntersection(Document doc)
        {
            var structuralElements = new List<(Element, Transform)>();
            
            StructuralElementLogger.LogStructuralElement("DIRECT_COLLECTION", 0, "STARTING_COLLECTION", "Collecting structural elements for direct solid intersection");
            
            // Collect from linked documents (where structural elements actually exist)
            var revitLinks = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(link => link.GetLinkDocument() != null)
                .ToList();
                
            foreach (var linkInstance in revitLinks)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                var linkName = linkInstance.Name;
                var linkTransform = linkInstance.GetTotalTransform(); // Critical: get the transformation matrix
                
                // Collect structural floors from linked document
                var linkedFloors = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType()
                    .Cast<Floor>()
                    .Where(floor => floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.AsInteger() == 1)
                    .ToList();
                    
                // Collect structural framing from linked document
                var linkedFraming = new FilteredElementCollector(linkDoc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(frame => frame.StructuralType != Autodesk.Revit.DB.Structure.StructuralType.NonStructural)
                    .ToList();
                    
                // Add with transform information
                structuralElements.AddRange(linkedFloors.Select(floor => (floor as Element, linkTransform)));
                structuralElements.AddRange(linkedFraming.Select(frame => (frame as Element, linkTransform)));
                
                StructuralElementLogger.LogStructuralElement("DIRECT_COLLECTION", (int)linkInstance.Id.Value, "LINKED_COLLECTED", 
                    $"From {linkName}: {linkedFloors.Count} structural floors, {linkedFraming.Count} structural framing, Transform: {linkTransform != null}");
            }
            
            // Also collect from active document (though we know there are none in this case)
            var activeFloors = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Floors)
                .WhereElementIsNotElementType()
                .Cast<Floor>()
                .Where(floor => floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)?.AsInteger() == 1)
                .ToList();
                
            var activeFraming = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
                .Cast<FamilyInstance>()
                .Where(frame => frame.StructuralType != Autodesk.Revit.DB.Structure.StructuralType.NonStructural)
                .ToList();
                
            // Add active document elements with null transform (no transformation needed)
            structuralElements.AddRange(activeFloors.Select(floor => (floor as Element, (Transform)null!)));
            structuralElements.AddRange(activeFraming.Select(frame => (frame as Element, (Transform)null!)));
            
            StructuralElementLogger.LogStructuralElement("DIRECT_COLLECTION", 0, "COLLECTION_COMPLETE", 
                $"Total collected: {structuralElements.Count} structural elements ({activeFloors.Count + activeFraming.Count} from active doc)");
                
            return structuralElements;
        }

        /// <summary>
        /// Performs direct solid intersection check between cable tray and structural elements
        /// </summary>
        private List<(Element structuralElement, XYZ intersectionPoint)> FindDirectStructuralIntersections(
            CableTray cableTray, List<(Element element, Transform linkTransform)> structuralElements)
        {
            var intersections = new List<(Element, XYZ)>();
            StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)cableTray.Id.Value, "STARTING_INTERSECTION", $"Testing cable tray against {structuralElements.Count} structural elements");
            try
            {
                // Get cable tray centerline as Line
                var curve = (cableTray.Location as LocationCurve)?.Curve as Line;
                if (curve == null)
                {
                    StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)cableTray.Id.Value, "NO_CURVE", "Cable tray has no valid centerline");
                    return intersections;
                }
                // For each structural element, get solid and check face/line intersection
                int elementIndex = 0;
                foreach (var (structuralElement, linkTransform) in structuralElements)
                {
                    elementIndex++;
                    try
                    {
                        bool isLinkedElement = linkTransform != null;
                        StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)structuralElement.Id.Value, "TESTING_ELEMENT", $"Testing element {elementIndex}/{structuralElements.Count}: {structuralElement.GetType().Name}, IsLinked: {isLinkedElement}");
                        var structuralOptions = new Options();
                        var structuralGeometry = structuralElement.get_Geometry(structuralOptions);
            Solid? structuralSolid = null;
                        foreach (var geomObj in structuralGeometry)
                        {
                            if (geomObj is Solid solid && solid.Volume > 0)
                            {
                                structuralSolid = solid;
                                break;
                            }
                            else if (geomObj is GeometryInstance instance)
                            {
                                foreach (var instObj in instance.GetInstanceGeometry())
                                {
                                    if (instObj is Solid instSolid && instSolid.Volume > 0)
                                    {
                                        structuralSolid = instSolid;
                                        break;
                                    }
                                }
                                if (structuralSolid != null) break;
                            }
                        }
                        if (structuralSolid == null)
                        {
                            StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)structuralElement.Id.Value, "NO_SOLID", "No solid geometry found in structural element");
                            continue;
                        }
                        // Apply linkTransform if this is a linked element
                        if (isLinkedElement && linkTransform != null)
                        {
                            structuralSolid = SolidUtils.CreateTransformed(structuralSolid, linkTransform);
                            StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)structuralElement.Id.Value, "APPLIED_TRANSFORM", $"Applied linkTransform: {linkTransform}");
                        }
                        // Log bounding box for structural solid
                        var structBbox = structuralSolid.GetBoundingBox();
                        StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)structuralElement.Id.Value, "STRUCT_BBOX", $"Min=({structBbox.Min.X:F2},{structBbox.Min.Y:F2},{structBbox.Min.Z:F2}), Max=({structBbox.Max.X:F2},{structBbox.Max.Y:F2},{structBbox.Max.Z:F2})");
                        // Check intersection using face.Intersect(line, out ira)
                        var intersectionPoints = new List<XYZ>();
                        foreach (Face face in structuralSolid.Faces)
                        {
                            IntersectionResultArray? ira = null;
                            SetComparisonResult res = face.Intersect(curve, out ira);
                            if (res == SetComparisonResult.Overlap && ira != null)
                            {
                                foreach (IntersectionResult ir in ira)
                                {
                                    intersectionPoints.Add(ir.XYZPoint);
                                }
                            }
                        }
                        if (intersectionPoints.Count > 0)
                        {
                            // If two or more intersection points, use the midpoint between the two furthest apart (entry/exit)
                            if (intersectionPoints.Count >= 2)
                            {
                                // Find the two points with the maximum distance between them
                                double maxDist = double.MinValue;
                                XYZ? ptA = null; XYZ? ptB = null;
                                for (int i = 0; i < intersectionPoints.Count - 1; i++)
                                {
                                    for (int j = i + 1; j < intersectionPoints.Count; j++)
                                    {
                                        double dist = intersectionPoints[i].DistanceTo(intersectionPoints[j]);
                                        if (dist > maxDist)
                                        {
                                            maxDist = dist;
                                            ptA = intersectionPoints[i];
                                            ptB = intersectionPoints[j];
                                        }
                                    }
                                }
                                if (ptA != null && ptB != null)
                                {
                                    var midpoint = new XYZ((ptA.X + ptB.X) / 2, (ptA.Y + ptB.Y) / 2, (ptA.Z + ptB.Z) / 2);
                                    intersections.Add((structuralElement, midpoint));
                                    StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)structuralElement.Id.Value, "INTERSECTION_POINT", $"Entry={ptA}, Exit={ptB}, Midpoint={midpoint}");
                                }
                            }
                            else
                            {
                                // Only one intersection point, use as is
                                var pt = intersectionPoints[0];
                                intersections.Add((structuralElement, pt));
                                StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)structuralElement.Id.Value, "INTERSECTION_POINT", $"Intersection at {pt}");
                            }
                        }
                        else
                        {
                            StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)structuralElement.Id.Value, "NO_INTERSECTION", "No intersection found with cable tray centerline");
                        }
                    }
                    catch (Exception ex)
                    {
                        StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)structuralElement.Id.Value, "ELEMENT_ERROR", $"Error testing element: {ex.Message}");
                    }
                }
                StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)cableTray.Id.Value, "INTERSECTION_COMPLETE", $"Tested {structuralElements.Count} elements, found {intersections.Count} intersections");
            }
            catch (Exception ex)
            {
                StructuralElementLogger.LogStructuralElement("DIRECT_INTERSECTION", (int)cableTray.Id.Value, "SOLID_ERROR", $"Error in direct intersection: {ex.Message}");
            }
            return intersections;
        }
    }
}