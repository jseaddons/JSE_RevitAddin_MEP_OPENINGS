using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using Autodesk.Revit.ApplicationServices;
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

            // Collect cable trays from both host and visible linked models
            var mepElements = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementCollectorHelper.CollectMepElementsVisibleOnly(doc);
            var trayTuples = mepElements
                .Where(tuple => tuple.Item1 is CableTray)
                .Select(tuple => (tuple.Item1 as CableTray, tuple.Item2))
                .Where(t => t.Item1 != null)
                .ToList();
            if (trayTuples.Count == 0)
            {
                TaskDialog.Show("Info", "No cable trays found in host or linked models.");
                return Result.Succeeded;
            }

            // Place sleeves for each cable tray
            var placer = new CableTraySleevePlacer(doc);

            // SUPPRESSION: Collect existing cable tray sleeves to avoid duplicates
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") && fi.Symbol.Name.StartsWith("CT#"))
                .ToList();

            DebugLogger.Log($"Found {existingSleeves.Count} existing cable tray sleeves in the model");

            // Log details of existing cable tray sleeves
            var allExistingCTSleeves = OpeningDuplicationChecker.FindCableTraySleeves(doc);
            if (allExistingCTSleeves.Any())
            {
                DebugLogger.Log($"--- Details of Existing Cable Tray Sleeves ---");
                foreach (var sleeve in allExistingCTSleeves)
                {
                    XYZ? sleeveLocation = (sleeve.Location as LocationPoint)?.Point;
                    double width = sleeve.LookupParameter("Width")?.AsDouble() ?? 0.0;
                    double height = sleeve.LookupParameter("Height")?.AsDouble() ?? 0.0;
                    double depth = sleeve.LookupParameter("Depth")?.AsDouble() ?? 0.0;

                    DebugLogger.Log($"  Sleeve ID: {sleeve.Id.Value}, Location: {sleeveLocation?.ToString() ?? "N/A"}, Width: {UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, Height: {UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm, Depth: {UnitUtils.ConvertFromInternalUnits(depth, UnitTypeId.Millimeters):F1}mm");
                }
                DebugLogger.Log($"------------------------------------------");
            }

            // Create a map of existing sleeve locations for quick lookup
            var existingSleeveLocations = existingSleeves.ToDictionary(
                sleeve => sleeve,
                sleeve => (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin
            );

            // Initialize counters for detailed logging
            int totalTrays = trayTuples.Count();
            int intersectionCount = 0;
            int placedCount = 0;
            int missingCount = 0;
            int skippedExistingCount = 0; // Counter for trays with existing sleeves
            int structuralElementsDetected = 0; // Counter for structural elements
            int structuralSleevesPlacer = 0; // Counter for successful structural sleeve placements
            int totalTrayTuples = trayTuples.Count();
            DebugLogger.Log($"Found {totalTrayTuples} cable trays to process");
            StructuralElementLogger.LogStructuralElement("SYSTEM", 0, "PROCESSING STARTED", $"Total cable trays to process: {totalTrayTuples}");

            // HashSet for duplicate suppression (robust, like pipes/ducts)
            HashSet<ElementId> processedCableTrays = new HashSet<ElementId>();

            using (var tx = new Transaction(doc, "Place Cable Tray Sleeves"))
            {
                DebugLogger.Log("Transaction started for cable tray sleeve placement");
                tx.Start();
                DebugLogger.Log("Starting wall intersection tests for cable trays");
                foreach (var tuple in trayTuples)
                {
                    var tray = tuple.Item1;
                    var transform = tuple.Item2;
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
                    // Transform geometry if from a linked model
                    Line hostLine = curve;
                    if (transform != null)
                    {
                        hostLine = Line.CreateBound(
                            transform.OfPoint(curve.GetEndPoint(0)),
                            transform.OfPoint(curve.GetEndPoint(1))
                        );
                    }
                    // Get tray dimensions for sleeve placement
                    double width = tray.LookupParameter("Width")?.AsDouble() ?? 0.0;
                    double height = tray.LookupParameter("Height")?.AsDouble() ?? 0.0;
                    // Enhanced multi-direction raycasting using efficient intersection service
                    XYZ rayDir = hostLine.Direction;
                    // Cast rays from multiple points along the cable tray (start, 25%, 50%, 75%, end)
                    var testPoints = new List<XYZ>
                     {
                         hostLine.GetEndPoint(0),           // Start point
                         hostLine.Evaluate(0.25, true),    // 25% along
                         hostLine.Evaluate(0.5, true),     // Midpoint  
                         hostLine.Evaluate(0.75, true),    // 75% along
                         hostLine.GetEndPoint(1)           // End point
                     };

                    // Use efficient intersection service for wall intersections
                    var allWallHits = EfficientIntersectionService.FindWallIntersections(tray, hostLine, view3D, testPoints, rayDir);

                    // Use efficient intersection service for structural intersections
                    var structuralIntersections = EfficientIntersectionService.FindStructuralIntersections(tray, hostLine, view3D);

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

                    // PRIORITIZE STRUCTURAL INTERSECTIONS FIRST (floors and beams)
                    // Process structural intersections using new direct approach
                    bool structuralSleeveePlaced = false;
                    // EXTENDED: Log all detected floor elements for this tray intersection
                    var allFloors = structuralIntersections.Where(t => t.Item1 is Floor).ToList();
                    if (allFloors.Count > 0)
                    {
                        DebugLogger.Log($"[CableTraySleeveCommand] CableTray {tray.Id.Value}: Detected {allFloors.Count} floor(s) at intersection:");
                        foreach (var floorTuple in allFloors)
                        {
                            var floorElem = (Floor)floorTuple.Item1;
                            var floorDoc = floorElem.Document;
                            var floorLoc = floorElem.Location as LocationPoint;
                            string locStr = floorLoc != null ? $"({floorLoc.Point.X:F3}, {floorLoc.Point.Y:F3}, {floorLoc.Point.Z:F3})" : "<no location>";
                            DebugLogger.Log($"[CableTraySleeveCommand]   Floor ID={floorElem.Id.Value}, Doc={floorDoc.Title}, Location={locStr}");
                        }
                    }
                    // Continue with normal intersection processing
                    foreach (var intersectionTuple in structuralIntersections)
                    {
                        if (processedCableTrays.Contains(tray.Id))
                        {
                            // Skip if already placed sleeve for this tray
                            continue;
                        }
                        // Fix tuple deconstruction for 3-tuple
                        var structuralElement = intersectionTuple.Item1;
                        var bbox = intersectionTuple.Item2;
                        var intersectionPoint = intersectionTuple.Item3;
                        structuralElementsDetected++;
                        string elementTypeName = structuralElement.Category?.Name ?? "STRUCTURAL";
                        DebugLogger.Log($"CableTray ID={tray.Id.Value}: detected structural element: {elementTypeName}, ID={structuralElement.Id.Value}");
                        StructuralElementLogger.LogStructuralElement(elementTypeName, (int)structuralElement.Id.Value, "STRUCTURAL DETECTED", $"Hit by cable tray {tray.Id.Value}");
                        StructuralElementLogger.LogStructuralElement("CableTray-STRUCTURAL INTERSECTION", 0, "INTERSECTION_DETAILS", $"CableTray ID={tray.Id.Value}, Structural ID={structuralElement.Id.Value}, Position=({intersectionPoint.X:F9}, {intersectionPoint.Y:F9}, {intersectionPoint.Z:F9})");
                        // For wall intersections, always use wall family
                        XYZ sleeveDirection = rayDir;
                        FamilySymbol? familySymbolToUse = null;
                        string linkedReferenceType = "UNKNOWN";
                        if (structuralElement is Wall)
                        {
                            familySymbolToUse = ctWallSymbols.FirstOrDefault();
                            linkedReferenceType = "WALL";
                        }
                        else if (structuralElement is Floor floor)
                        {
                            familySymbolToUse = ctSlabSymbols.FirstOrDefault();
                            linkedReferenceType = "FLOOR";

                            // EXTENDED FLOOR DEBUG LOGGING
                            var isStructuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                            bool isStructural = isStructuralParam != null && isStructuralParam.AsInteger() == 1;
                            string floorStructuralStatus = isStructural ? "STRUCTURAL" : "NON-STRUCTURAL";
                            string docTitle = floor.Document.Title;
                            string linkInfo = (transform != null) ? $"LINKED (Transform: {transform.Origin.X:F2},{transform.Origin.Y:F2},{transform.Origin.Z:F2})" : "HOST";
                            string floorDebugMsg = $"CABLETRAY FLOOR DEBUG: CableTray {tray.Id.Value} intersects Floor {floor.Id.Value} [{docTitle}] - Status: {floorStructuralStatus}, {linkInfo}";
                            DebugLogger.Log($"[CableTraySleeveCommand] {floorDebugMsg}");
                            if (isStructuralParam != null)
                                DebugLogger.Log($"[CableTraySleeveCommand] FLOOR_PARAM_IS_STRUCTURAL value: {isStructuralParam.AsInteger()} (1=structural, 0=non-structural)");
                            else
                                DebugLogger.Log($"[CableTraySleeveCommand] FLOOR_PARAM_IS_STRUCTURAL parameter is NULL");
                            if (transform != null)
                                DebugLogger.Log($"[CableTraySleeveCommand] Floor transform: Origin=({transform.Origin.X:F2},{transform.Origin.Y:F2},{transform.Origin.Z:F2})");
                            // END EXTENDED FLOOR DEBUG LOGGING

                            if (!isStructural)
                            {
                                string skipMsg = $"SKIP: CableTray {tray.Id.Value} host Floor {floor.Id.Value} [{docTitle}] is NON-STRUCTURAL. Sleeve will NOT be placed.";
                                DebugLogger.Log($"[CableTraySleeveCommand] {skipMsg}");
                                skippedExistingCount++;
                                // Don't mark as processed yet - let it try other floors
                                continue;
                            }

                            // Check floor family symbol availability  
                            if (familySymbolToUse != null)
                            {
                                string symbolMsg = $"CABLETRAY FLOOR SYMBOL: Found floor symbol: {familySymbolToUse.Family.Name} - {familySymbolToUse.Name}";
                                DebugLogger.Log($"[CableTraySleeveCommand] {symbolMsg}");
                            }
                            else
                            {
                                string noSymbolMsg = $"CABLETRAY FLOOR SYMBOL ERROR: No floor sleeve symbol available for CableTray {tray.Id.Value}";
                                DebugLogger.Log($"[CableTraySleeveCommand] {noSymbolMsg}");
                            }
                        }
                        else if (structuralElement is FamilyInstance famInst2 && famInst2.Category != null && famInst2.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                        {
                            familySymbolToUse = ctWallSymbols.FirstOrDefault();
                            linkedReferenceType = "STRUCTURAL FRAMING (CABLETRAY FAMILY)";
                        }
                        else
                        {
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

                        // For wall: always use wall family, no orientation logic
                        if (structuralElement is Wall || (structuralElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming))
                        {
                            // For wall and framing: use orientation/rotation logic (bounding box, width/height swap, rotation)
                            if (PlaceCableTraySleeveAtLocation_StructuralWithOrientation(doc, ctWallSymbols.FirstOrDefault(), structuralElement, intersectionPoint, sleeveDirection, width, height, tray))
                            {
                                structuralSleevesPlacer++;
                                placedCount++;
                                processedCableTrays.Add(tray.Id);
                                structuralSleeveePlaced = true;
                                DebugLogger.Log($"CableTray ID={(int)tray.Id.Value}: Structural sleeve successfully placed at {intersectionPoint} with width-based orientation");
                                break; // Only place one sleeve per cable tray
                            }
                        }
                        // For floor: use slab family and orientation logic
                        else if (structuralElement is Floor)
                        {
                            if (PlaceCableTraySleeveAtLocation_StructuralWithOrientation(doc, familySymbolToUse, structuralElement, intersectionPoint, sleeveDirection, width, height, tray))
                            {
                                structuralSleevesPlacer++;
                                placedCount++;
                                processedCableTrays.Add(tray.Id);
                                structuralSleeveePlaced = true;
                                DebugLogger.Log($"CableTray ID={(int)tray.Id.Value}: Structural sleeve successfully placed at {intersectionPoint} with width-based orientation");
                                break; // Only place one sleeve per cable tray
                            }
                        }
                        else
                        {
                            skippedExistingCount++;
                            processedCableTrays.Add(tray.Id); // Still mark as processed to avoid duplicates
                            StructuralElementLogger.LogStructuralElement(elementTypeName, (int)structuralElement.Id.Value, "SLEEVE_FAILED", "Reason: Existing sleeve found at location or placement failed");
                        }
                    }

                    // ONLY check walls if NO structural sleeve was placed
                    if (!structuralSleeveePlaced && allWallHits.Any())
                    {
                        // Find the hit closest to the cable tray path (prioritize hits from cable tray endpoints)
                        var closestHit = allWallHits.OrderBy(x => x.hit.Proximity).First();
                        var fallbackWallHit = closestHit.hit;
                        var fallbackWallDir = closestHit.direction;
                        var fallbackWallRayOrigin = closestHit.rayOrigin;

                        var r = fallbackWallHit.GetReference();
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
                            XYZ intersectionPoint = fallbackWallRayOrigin + fallbackWallDir * fallbackWallHit.Proximity;
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
                            if (PlaceCableTraySleeveAtLocation_Wall(doc, wallFamilySymbol, linkedReferenceElement, intersectionPoint, fallbackWallDir ?? XYZ.BasisX, width, height, tray.Id))
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
                    // Calculate wall depth using proper wall thickness parameter (not "b" which is for structural framing)
                    double wallDepth = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                    if (wallDepth <= 0.0)
                    {
                        // Fallback: try to get thickness from wall type
                        var wallType = wall.WallType;
                        var widthParam = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                        if (widthParam != null && widthParam.StorageType == StorageType.Double)
                        {
                            wallDepth = widthParam.AsDouble();
                        }
                        else
                        {
                            wallDepth = UnitUtils.ConvertToInternalUnits(200.0, UnitTypeId.Millimeters); // 200mm fallback
                            DebugLogger.Log($"CableTray ID={(int)trayId.Value}: Wall thickness parameter not found, using 200mm fallback");
                        }
                    }
                    DebugLogger.Log($"CableTray ID={(int)trayId.Value}: Wall thickness calculated: {UnitUtils.ConvertFromInternalUnits(wallDepth, UnitTypeId.Millimeters):F1}mm");

                    // Bounding box extent log for cable tray sleeve (wall)
                    double minX = placementPoint.X - width / 2.0;
                    double maxX = placementPoint.X + width / 2.0;
                    double minY = placementPoint.Y - height / 2.0;
                    double maxY = placementPoint.Y + height / 2.0;
                    DebugLogger.Log($"CableTray ID={(int)trayId.Value}: [BBOX-DEBUG] Placement at ({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3}), BBox X=({UnitUtils.ConvertFromInternalUnits(minX, UnitTypeId.Millimeters):F1}, {UnitUtils.ConvertFromInternalUnits(maxX, UnitTypeId.Millimeters):F1}), Y=({UnitUtils.ConvertFromInternalUnits(minY, UnitTypeId.Millimeters):F1}, {UnitUtils.ConvertFromInternalUnits(maxY, UnitTypeId.Millimeters):F1}), Width={UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm");

                    // ALWAYS use XYZ.BasisX for wall/framing, never pass cable tray direction
                    FamilyInstance? sleeveInstance = placer.PlaceCableTraySleeve((CableTray)null!, placementPoint, width, height, XYZ.BasisX, ctWallSymbol, wall);
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
            XYZ _direction, // unused, always use default
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
                // For wall/framing, always use XYZ.BasisX (default orientation), never pass cable tray direction
                var sleeveInstance = placer.PlaceCableTraySleeve(
                    tray, // pass the actual tray object!
                    placementPoint,
                    width,
                    height,
                    XYZ.BasisX, // always use default orientation for wall/framing
                    ctSlabSymbol,
                    hostElement);
                if (sleeveInstance != null)
                {
                    // --- Set clearance/depth for floors and beams (structural framing) ---
                    double clearance = 0.0;
                    if (hostElement is Floor floor)
                    {
                        var thicknessParam = floor.LookupParameter("Default Thickness");
                        if (thicknessParam != null && thicknessParam.StorageType == StorageType.Double)
                        {
                            clearance = thicknessParam.AsDouble();
                        }
                    }
                    else if (hostElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                    {
                        var bParam = famInst.Symbol.LookupParameter("b");
                        if (bParam != null && bParam.StorageType == StorageType.Double)
                        {
                            clearance = bParam.AsDouble();
                        }
                    }
                    // Set the clearance/depth parameter if found
                    if (clearance > 0.0)
                    {
                        var depthParam = sleeveInstance.LookupParameter("Depth") ?? sleeveInstance.LookupParameter("d");
                        if (depthParam != null && depthParam.StorageType == StorageType.Double && !depthParam.IsReadOnly)
                        {
                            depthParam.Set(clearance);
                        }
                    }
                    return true;
                }
                else
                {
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
        /// Places a cable tray structural sleeve with width-based orientation for floors (similar to duct logic)
        /// </summary>
        private bool PlaceCableTraySleeveAtLocation_StructuralWithOrientation(
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
                int cableTrayId = (int)tray.Id.Value;
                DebugLogger.Log($"[CableTraySleeveCommand] ORIENTATION ANALYSIS for CableTray {cableTrayId}");
                
                // For framing, use location curve direction for orientation
                XYZ? preCalculatedOrientation = null;
                double widthToUse = width;
                double heightToUse = height;
                bool isFraming = hostElement is FamilyInstance famInst1 && famInst1.Category != null && famInst1.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming;
                if (isFraming)
                {
                    var curve = (tray.Location as LocationCurve)?.Curve as Line;
                    if (curve != null)
                    {
                        var dir = curve.Direction;
                        DebugLogger.Log($"[CableTraySleeveCommand] [FRAMING] Tray location curve direction: ({dir.X:F6},{dir.Y:F6},{dir.Z:F6})");
                        // If running along Y (abs(Y) > abs(X)), treat as Y-oriented
                        if (Math.Abs(dir.Y) > Math.Abs(dir.X))
                        {
                            preCalculatedOrientation = XYZ.BasisY;
                            widthToUse = height;
                            heightToUse = width;
                            DebugLogger.Log($"[CableTraySleeveCommand] [FRAMING] Detected Y-oriented tray. Swapping width/height. Orientation=Y");
                        }
                        else
                        {
                            preCalculatedOrientation = XYZ.BasisX;
                            DebugLogger.Log($"[CableTraySleeveCommand] [FRAMING] Detected X-oriented tray. Orientation=X");
                        }
                    }
                    else
                    {
                        DebugLogger.Log($"[CableTraySleeveCommand] [FRAMING] ERROR: Could not get location curve for tray");
                    }
                }
                else
                {
                    // For floors and all other hosts, use bounding box width alignment (GetCableTrayWidthOrientation)
                    var orientationResult = JSE_RevitAddin_MEP_OPENINGS.Helpers.MepElementOrientationHelper.GetCableTrayWidthOrientation(tray);
                    string orientationStatus = orientationResult.orientation;
                    XYZ widthDirection = orientationResult.widthDirection;
                    DebugLogger.Log($"[CableTraySleeveCommand] Flow Direction: ({direction.X:F6},{direction.Y:F6},{direction.Z:F6})");
                    DebugLogger.Log($"[CableTraySleeveCommand] Width Direction: ({widthDirection.X:F6},{widthDirection.Y:F6},{widthDirection.Z:F6})");
                    DebugLogger.Log($"[CableTraySleeveCommand] CableTray {cableTrayId}: {orientationStatus}");
                    // For floors, set preCalculatedOrientation and swap width/height if Y-oriented
                    bool isFloor = hostElement is Floor;
                    if (isFloor)
                    {
                        if (orientationStatus == "Y-ORIENTED")
                        {
                            preCalculatedOrientation = XYZ.BasisY;
                            DebugLogger.Log($"[CableTraySleeveCommand] [FLOOR] Detected Y-oriented tray (bounding box). Passing Y orientation for rotation. No width/height swap.");
                        }
                        else if (orientationStatus == "X-ORIENTED")
                        {
                            preCalculatedOrientation = XYZ.BasisX;
                            DebugLogger.Log($"[CableTraySleeveCommand] [FLOOR] Detected X-oriented tray (bounding box). Orientation=X");
                        }
                    }
                }
                
                double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                // --- Robust duplicate suppression and logging (like DuctSleeveCommand) ---
                string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placementPoint, sleeveCheckRadius);
                DebugLogger.Log($"[CableTraySleeveCommand] DUPLICATION CHECK at {placementPoint}:\n{sleeveSummary}");
                var indivDup = OpeningDuplicationChecker.FindIndividualSleevesAtLocation(doc, placementPoint, sleeveCheckRadius);
                var clusterDup = OpeningDuplicationChecker.FindAllClusterSleevesAtLocation(doc, placementPoint, sleeveCheckRadius);
                if (indivDup.Any() || clusterDup.Any())
                {
                    string msg = $"SKIP: CableTray {cableTrayId} duplicate sleeve (individual or cluster) exists near {placementPoint}";
                    DebugLogger.Log($"[CableTraySleeveCommand] {msg}");
                    return false;
                }

                var placer = new CableTraySleevePlacer(doc);
                var sleeveInstance = placer.PlaceCableTraySleeveWithOrientation(
                    tray,
                    placementPoint,
                    widthToUse,
                    heightToUse,
                    direction,
                    preCalculatedOrientation,
                    ctSlabSymbol,
                    hostElement);
                    
                if (sleeveInstance != null)
                {
                    // --- Set clearance/depth for floors and beams (structural framing) ---
                    double clearance = 0.0;
                    if (hostElement is Floor floor)
                    {
                        var thicknessParam = floor.LookupParameter("Default Thickness");
                        if (thicknessParam != null && thicknessParam.StorageType == StorageType.Double)
                        {
                            clearance = thicknessParam.AsDouble();
                        }
                    }
                else if (hostElement is FamilyInstance famInst2 && famInst2.Category != null && famInst2.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                    {
                        var bParam = famInst2.Symbol.LookupParameter("b");
                        if (bParam != null && bParam.StorageType == StorageType.Double)
                        {
                            clearance = bParam.AsDouble();
                        }
                    }
                    // Set the clearance/depth parameter if found
                    if (clearance > 0.0)
                    {
                        var depthParam = sleeveInstance.LookupParameter("Depth") ?? sleeveInstance.LookupParameter("d");
                        if (depthParam != null && depthParam.StorageType == StorageType.Double && !depthParam.IsReadOnly)
                        {
                            depthParam.Set(clearance);
                        }
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"CableTray ID={(int)tray.Id.Value}: error placing structural sleeve with orientation: {ex.Message}");
                return false;
            }
        }


    }
}