using System.Linq;
using System.IO;
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

            // Initialize cable tray log file for diagnostics
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.SetCableTrayLogFile();
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.InitLogFile();

            // ...existing code...

            // Log available families for debugging
            var allFamilySymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Where(sym => sym.Family != null)
                .ToList();
            
            StructuralElementLogger.LogStructuralElement("DIAGNOSTIC", new Autodesk.Revit.DB.ElementId((BuiltInParameter)0L), "CT_FAMILY_SEARCH", $"Found {allFamilySymbols.Count} family symbols in project");
            
            var cableTrayFamilies = allFamilySymbols
                .Where(sym => sym.Family.Name.IndexOf("CableTray", System.StringComparison.OrdinalIgnoreCase) >= 0 ||
                              sym.Family.Name.IndexOf("CT", System.StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
            
            foreach (var fam in cableTrayFamilies.Take(10)) // Log first 10 cable tray families
            {
                StructuralElementLogger.LogStructuralElement("DIAGNOSTIC", fam.Id, "CT_FAMILY_FOUND", $"Family: {fam.Family.Name}, Symbol: {fam.Name}");
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
                StructuralElementLogger.LogStructuralElement("ERROR", new Autodesk.Revit.DB.ElementId((BuiltInParameter)0L), "MISSING_CT_FAMILY", "Could not find cable tray sleeve family (CableTrayOpeningOnWall or CableTrayOpeningOnSlab)");
                TaskDialog.Show("Error", "Please load cable tray sleeve opening families (wall and slab).");
                return Result.Failed;
            }
            foreach (var sym in ctWallSymbols)
                StructuralElementLogger.LogStructuralElement("SUCCESS", sym.Id, "CT_FAMILY_FOUND", $"Using cable tray wall sleeve family: {sym.Family.Name}, Symbol: {sym.Name}");
            foreach (var sym in ctSlabSymbols)
                StructuralElementLogger.LogStructuralElement("SUCCESS", sym.Id, "CT_FAMILY_FOUND", $"Using cable tray slab sleeve family: {sym.Family.Name}, Symbol: {sym.Name}");

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
            var allSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => (fi.Symbol.Family.Name.Contains("OpeningOnWall") || fi.Symbol.Family.Name.Contains("OpeningOnSlab")))
                .ToList();

            BoundingBoxXYZ? sectionBox = null;
            try
            {
                if (doc.ActiveView is View3D vb)
                    sectionBox = SectionBoxHelper.GetSectionBoxBounds(vb);
            }
            catch { /* ignore */ }

            if (sectionBox != null)
            {
                allSleeves = allSleeves.Where(s =>
                {
                    var bb = s.get_BoundingBox(null);
                    return bb != null && BoundingBoxesIntersect(bb, sectionBox);
                }).ToList();
            }

            var sleeveGrid = new SleeveSpatialGrid(allSleeves);
            DebugLogger.Log($"Found {allSleeves.Count} existing cable tray sleeves visible in the active section box (or total if fallback)");

            // Log details of existing cable tray sleeves (doc-wide helper - kept for diagnostic value)
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

                    DebugLogger.Log($"  Sleeve ID: {sleeve.Id.IntegerValue}, Location: {sleeveLocation?.ToString() ?? "N/A"}, Width: {UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, Height: {UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm, Depth: {UnitUtils.ConvertFromInternalUnits(depth, UnitTypeId.Millimeters):F1}mm");
                }
                DebugLogger.Log($"------------------------------------------");
            }

            // Create a map of existing sleeve locations for quick lookup
            var existingSleeveLocations = allSleeves.ToDictionary(
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
            StructuralElementLogger.LogStructuralElement("SYSTEM", new Autodesk.Revit.DB.ElementId((BuiltInParameter)0L), "PROCESSING STARTED", $"Total cable trays to process: {totalTrayTuples}");

            // HashSet for duplicate suppression (robust, like pipes/ducts)
            HashSet<ElementId> processedCableTrays = new HashSet<ElementId>();

            void Log(string msg) => DebugLogger.Log(msg);

            // Collect structural elements using section box filtering (same as ducts/pipes)
            var structuralElements = JSE_RevitAddin_MEP_OPENINGS.Services.MepIntersectionService.CollectStructuralElementsForDirectIntersectionVisibleOnly(doc, Log);

            var spatialService = new SpatialPartitioningService(structuralElements);

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
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: already processed, skipping to prevent duplicate sleeve");
                        continue;
                    }
                    DebugLogger.Log($"Processing CableTray ID={tray.Id.IntegerValue}");
                    var curve = (tray.Location as LocationCurve)?.Curve as Line;
                    if (curve == null) {
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: no valid location curve, skipping");
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

                    
                    // Transform bounding box if from linked model (critical for spatial filtering)
                    var trayBBox = tray.get_BoundingBox(null);
                    if (transform != null && trayBBox != null)
                    {
                        var transformedMin = transform.OfPoint(trayBBox.Min);
                        var transformedMax = transform.OfPoint(trayBBox.Max);
                        trayBBox = new BoundingBoxXYZ
                        {
                            Min = new XYZ(Math.Min(transformedMin.X, transformedMax.X), Math.Min(transformedMin.Y, transformedMax.Y), Math.Min(transformedMin.Z, transformedMax.Z)),
                            Max = new XYZ(Math.Max(transformedMin.X, transformedMax.X), Math.Max(transformedMin.Y, transformedMax.Y), Math.Max(transformedMin.Z, transformedMax.Z))
                        };
                    }
                    
                    var nearbyStructuralElements = spatialService.GetNearbyElements(tray);
                    if (!nearbyStructuralElements.Any())
                    {
                        DebugLogger.Log($"WARNING: CableTray {tray.Id} no nearby structural elements found via spatial partitioning. Falling back to all structural elements.");
                        
                        // FALLBACK: Use all structural elements when spatial partitioning fails
                        // This ensures the system works regardless of coordinate/precision issues
                        nearbyStructuralElements = structuralElements;
                        DebugLogger.Log($"Fallback: Using all {nearbyStructuralElements.Count} structural elements for cable tray {tray.Id}");
                    }

                    // Use MepIntersectionService for wall intersections (like duct logic)
                    var wallIntersections = JSE_RevitAddin_MEP_OPENINGS.Services.MepIntersectionService.FindIntersections(hostLine, trayBBox, nearbyStructuralElements, Log);

                    // Convert wallIntersections to the same format as allWallHits for downstream code
                    var allWallHits = wallIntersections
                        .Select(t => (
                            hit: (ReferenceWithContext?)null, // Not used downstream, so can be null
                            direction: rayDir,
                            rayOrigin: t.Item3 // intersectionPoint
                        ))
                        .ToList();
                    DebugLogger.Log($"[CableTraySleeveCommand] CableTray {tray.Id.IntegerValue}: allWallHits count = {allWallHits.Count}");

                    // Use shared intersection service (same pattern as pipe/duct): prefer host-line overload for linked trays
                    List<(Element, BoundingBoxXYZ, XYZ)> structuralIntersections;
                    if (transform != null)
                    {
                        structuralIntersections = JSE_RevitAddin_MEP_OPENINGS.Services.MepIntersectionService.FindIntersections(hostLine, trayBBox, nearbyStructuralElements, Log);
                    }
                    else
                    {
                        structuralIntersections = JSE_RevitAddin_MEP_OPENINGS.Services.MepIntersectionService.FindIntersections(tray, nearbyStructuralElements, Log);
                    }
                    DebugLogger.Log($"[CableTraySleeveCommand] CableTray {tray.Id.IntegerValue}: structuralIntersections count = {structuralIntersections.Count}");

                    // Process wall intersections (existing logic, keep working)
                    ReferenceWithContext? bestWallHit = null;
                    XYZ? bestWallDir = null;
                    XYZ? bestWallRayOrigin = null;

                    if (allWallHits.Any())
                    {
                        // Find the hit closest to the cable tray path (prioritize hits from cable tray endpoints)
                        var closestHit = allWallHits
                            .Where(x => x.hit != null)
                            .OrderBy(x => x.hit!.Proximity)
                            .FirstOrDefault();

                        if (closestHit.hit != null)
                        {
                            bestWallHit = closestHit.hit;
                            bestWallDir = closestHit.direction;
                            bestWallRayOrigin = closestHit.rayOrigin;
                        }
                    }

                    int wallHitCount = allWallHits.Count;
                    int structuralHitCount = structuralIntersections.Count;
                    DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Wall hits: {wallHitCount}, Structural hits: {structuralHitCount}");

                    // PRIORITIZE STRUCTURAL INTERSECTIONS FIRST (floors and beams)
                    // Process structural intersections using new direct approach
                    bool structuralSleeveePlaced = false;
                    // EXTENDED: Log all detected floor elements for this tray intersection
                    var allFloors = structuralIntersections.Where(t => t.Item1 is Floor).ToList();
                    if (allFloors.Count > 0)
                    {
                        DebugLogger.Log($"[CableTraySleeveCommand] CableTray {tray.Id.IntegerValue}: Detected {allFloors.Count} floor(s) at intersection:");
                        foreach (var floorTuple in allFloors)
                        {
                            var floorElem = (Floor)floorTuple.Item1;
                            var floorDoc = floorElem.Document;
                            var floorLoc = floorElem.Location as LocationPoint;
                            string locStr = floorLoc != null ? $"({floorLoc.Point.X:F3}, {floorLoc.Point.Y:F3}, {floorLoc.Point.Z:F3})" : "<no location>";
                            DebugLogger.Log($"[CableTraySleeveCommand]   Floor ID={floorElem.Id.IntegerValue}, Doc={floorDoc.Title}, Location={locStr}");
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
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: detected structural element: {elementTypeName}, ID={structuralElement.Id.IntegerValue}");
                        StructuralElementLogger.LogStructuralElement(elementTypeName, structuralElement.Id, "STRUCTURAL DETECTED", $"Hit by cable tray {tray.Id.IntegerValue}");
                        StructuralElementLogger.LogStructuralElement("CableTray-STRUCTURAL INTERSECTION", new Autodesk.Revit.DB.ElementId((BuiltInParameter)0L), "INTERSECTION_DETAILS", $"CableTray ID={tray.Id.IntegerValue}, Structural ID={structuralElement.Id.IntegerValue}, Position=({intersectionPoint.X:F9}, {intersectionPoint.Y:F9}, {intersectionPoint.Z:F9})");
                        // For wall intersections, always use wall family
                        XYZ sleeveDirection = rayDir;
                        FamilySymbol? familySymbolToUse = null;
                        string linkedReferenceType = "UNKNOWN";

                        // If the structural element lives in a linked document, transform the
                        // intersection point into the link-local coordinate space so the
                        // placer and centerline projection operate in the same coordinates
                        // as the host element (matches wall handling above).
                        XYZ intersectionToPass = intersectionPoint;
                        XYZ dirToPass = rayDir;
                        try
                        {
                            // IMPORTANT: MepIntersectionService already returns intersection points in host/document coordinates
                            // Do NOT invert back into link-local coordinates; doing so and then creating a FamilyInstance in the
                            // active document mixes coordinate spaces and results in placement in empty space.
                            // Always pass the host-space intersection straight to the placer.
                            intersectionToPass = intersectionPoint;
                            dirToPass = rayDir;
                            DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Passing host-space intersection to placer: ({intersectionToPass.X:F6},{intersectionToPass.Y:F6},{intersectionToPass.Z:F6})");
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Error handling structural intersection coordinates: {ex.Message}");
                        }
                        if (structuralElement is Wall)
                        {
                            familySymbolToUse = ctWallSymbols.FirstOrDefault();
                            if (familySymbolToUse == null)
                            {
                                TaskDialog.Show("Error", "Cable tray wall sleeve family (CableTrayOpeningOnWall) is missing. Please load it and rerun the command.");
                                return Result.Failed;
                            }
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
                            string floorDebugMsg = $"CABLETRAY FLOOR DEBUG: CableTray {tray.Id.IntegerValue} intersects Floor {floor.Id.IntegerValue} [{docTitle}] - Status: {floorStructuralStatus}, {linkInfo}";
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
                                string skipMsg = $"SKIP: CableTray {tray.Id.IntegerValue} host Floor {floor.Id.IntegerValue} [{docTitle}] is NON-STRUCTURAL. Sleeve will NOT be placed.";
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
                                string noSymbolMsg = $"CABLETRAY FLOOR SYMBOL ERROR: No floor sleeve symbol available for CableTray {tray.Id.IntegerValue}";
                                DebugLogger.Log($"[CableTraySleeveCommand] {noSymbolMsg}");
                            }
                        }
                        else if (structuralElement is FamilyInstance famInst2 && famInst2.Category != null && famInst2.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                        {
                            familySymbolToUse = ctWallSymbols.FirstOrDefault();
                            linkedReferenceType = "STRUCTURAL FRAMING (CABLETRAY FAMILY)";
                        }
                        else
                        {
                            familySymbolToUse = ctSlabSymbols.FirstOrDefault();
                            linkedReferenceType = structuralElement?.Category?.Name ?? "UNKNOWN";
                        }
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Linked reference type detected: {linkedReferenceType}");
                        if (familySymbolToUse == null)
                        {
                            TaskDialog.Show("Error", $"Cable tray sleeve family for {linkedReferenceType} is missing. Please load it and rerun the command.");
                            StructuralElementLogger.LogStructuralElement(elementTypeName, structuralElement?.Id ?? new Autodesk.Revit.DB.ElementId(0), "SLEEVE_FAILED", $"Reason: {linkedReferenceType} family symbol not found");
                            skippedExistingCount++;
                            processedCableTrays.Add(tray.Id);
                            continue;
                        }
                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Using family: {familySymbolToUse.Family.Name}, Symbol: {familySymbolToUse.Name} for linked reference type {linkedReferenceType}");

                        // For wall: always use wall family, no orientation logic
                            if (structuralElement is Wall || (structuralElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming))
                        {
                        // For wall and framing: use orientation/rotation logic (bounding box, width/height swap, rotation)
                        if (PlaceCableTraySleeveAtLocation_StructuralWithOrientation(doc, sleeveGrid, ctWallSymbols.FirstOrDefault(), structuralElement, intersectionToPass, dirToPass, width, height, tray))
                        {
                            structuralSleevesPlacer++;
                                placedCount++;
                                processedCableTrays.Add(tray.Id);
                                structuralSleeveePlaced = true;
                                DebugLogger.Log($"CableTray ID={(int)tray.Id.IntegerValue}: Structural sleeve successfully placed at {intersectionPoint} with width-based orientation");
                                break; // Only place one sleeve per cable tray
                            }
                        }
                        // For floor: use slab family and orientation logic
                        else if (structuralElement is Floor)
                        {
                            if (PlaceCableTraySleeveAtLocation_StructuralWithOrientation(doc, sleeveGrid, familySymbolToUse, structuralElement, intersectionToPass, dirToPass, width, height, tray))
                            {
                                structuralSleevesPlacer++;
                                placedCount++;
                                processedCableTrays.Add(tray.Id);
                                structuralSleeveePlaced = true;
                                DebugLogger.Log($"CableTray ID={(int)tray.Id.IntegerValue}: Structural sleeve successfully placed at {intersectionPoint} with width-based orientation");
                                break; // Only place one sleeve per cable tray
                            }
                        }
                        else
                        {
                            skippedExistingCount++;
                            processedCableTrays.Add(tray.Id); // Still mark as processed to avoid duplicates
                            StructuralElementLogger.LogStructuralElement(elementTypeName, structuralElement?.Id ?? new Autodesk.Revit.DB.ElementId(0), "SLEEVE_FAILED", "Reason: Existing sleeve found at location or placement failed");
                        }
                    }

                    // ONLY check walls if NO structural sleeve was placed
                    if (!structuralSleeveePlaced && allWallHits.Any())
                    {
                        // Find the hit closest to the cable tray path (prioritize hits from cable tray endpoints)
                        var closestHit = allWallHits
                            .Where(x => x.hit != null)
                            .OrderBy(x => x.hit!.Proximity)
                            .FirstOrDefault();

                        if (closestHit.hit == null)
                        {
                            DebugLogger.Log("No valid wall hit found for cable tray, skipping.");
                            continue;
                        }
                        var fallbackWallHit = closestHit.hit;
                        var fallbackWallDir = closestHit.direction;
                        var fallbackWallRayOrigin = closestHit.rayOrigin;

                        if (fallbackWallHit != null)
                        {
                            var r = fallbackWallHit.GetReference();
                            var linkInst = doc.GetElement(r.ElementId) as RevitLinkInstance;
                            var targetDoc = linkInst != null ? linkInst.GetLinkDocument() : doc;
                            ElementId elemId = linkInst != null ? r.LinkedElementId : r.ElementId;
                            Element? linkedReferenceElement = targetDoc?.GetElement(elemId);

                            if (linkedReferenceElement != null)
                            {
                                string elementTypeName = linkedReferenceElement.Category?.Name ?? "NO_CATEGORY";
                                DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: detected wall element: {elementTypeName}, ID={linkedReferenceElement.Id.IntegerValue}");
                                StructuralElementLogger.LogStructuralElement(elementTypeName, linkedReferenceElement.Id, "ELEMENT DETECTED", $"Hit by cable tray {tray.Id.IntegerValue}");

                                // Calculate intersection point for wall (in host coordinates)
                                XYZ intersectionPoint = fallbackWallRayOrigin + fallbackWallDir * fallbackWallHit.Proximity;

                                // If the wall is in a linked document, transform the intersection into the link-local coordinate
                                // space before passing to the placer so projection onto wall centerline is correct.
                                XYZ intersectionForHost = intersectionPoint;
                                XYZ intersectionForHostLog = intersectionPoint;
                                XYZ intersectionToPass = intersectionPoint;
                                XYZ dirToPass = fallbackWallDir ?? XYZ.BasisX;
                                if (linkInst != null)
                                {
                                    var linkTransform = linkInst.GetTotalTransform();
                                    try
                                    {
                                        var inv = linkTransform.Inverse;
                                        intersectionToPass = inv.OfPoint(intersectionPoint);
                                        dirToPass = inv.OfVector(fallbackWallDir ?? XYZ.BasisX);
                                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Transformed intersection into link-local coords: host=({intersectionForHost.X:F6},{intersectionForHost.Y:F6},{intersectionForHost.Z:F6}) -> local=({intersectionToPass.X:F6},{intersectionToPass.Y:F6},{intersectionToPass.Z:F6})");
                                    }
                                    catch (Exception ex)
                                    {
                                        DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Error transforming intersection into link-local coords: {ex.Message}");
                                    }
                                }

                                StructuralElementLogger.LogStructuralElement("CableTray-WALL INTERSECTION", new Autodesk.Revit.DB.ElementId((BuiltInParameter)0L), "INTERSECTION_DETAILS", $"CableTray ID={tray.Id.IntegerValue}, Wall ID={linkedReferenceElement.Id.IntegerValue}, Position=({intersectionForHostLog.X:F9}, {intersectionForHostLog.Y:F9}, {intersectionForHostLog.Z:F9})");

                                // Always use CableTrayOpeningOnWall family for wall/framing
                                var wallFamilySymbol = ctWallSymbols.FirstOrDefault();
                                if (wallFamilySymbol == null)
                                {
                                    TaskDialog.Show("Error", "Cable tray wall sleeve family (CableTrayOpeningOnWall) is missing. Please load it and rerun the command.");
                                    StructuralElementLogger.LogStructuralElement(elementTypeName, linkedReferenceElement.Id, "SLEEVE_FAILED", "Reason: Wall family symbol not found");
                                    skippedExistingCount++;
                                    processedCableTrays.Add(tray.Id);
                                    continue;
                                }
                                DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Using wall family: {wallFamilySymbol.Family.Name}, Symbol: {wallFamilySymbol.Name}");
                                if (PlaceCableTraySleeveAtLocation_Wall(doc, sleeveGrid, wallFamilySymbol, linkedReferenceElement, intersectionToPass, dirToPass, width, height, tray.Id))
                                {
                                    placedCount++;
                                    processedCableTrays.Add(tray.Id);
                                    DebugLogger.Log($"CableTray ID={tray.Id.IntegerValue}: Wall sleeve successfully placed at {intersectionPoint}");
                                }
                                else
                                {
                                    skippedExistingCount++;
                                    processedCableTrays.Add(tray.Id); // Still mark as processed to avoid duplicates
                                    StructuralElementLogger.LogStructuralElement(elementTypeName, linkedReferenceElement.Id, "SLEEVE_FAILED", "Reason: Existing sleeve found at location");
                                }
                            }
                        }
                    }

                    // If no wall or structural intersection found
                    if (bestWallHit == null && structuralIntersections.Count == 0)
                    {
                        missingCount++;
                        DebugLogger.Log($"CableTray ID={(int)tray.Id.IntegerValue}: no wall or structural intersection detected, skipping");
                        continue;
                    }
                }
                
                tx.Commit();
                DebugLogger.Log("Transaction committed for cable tray sleeve placement");
                
                // Summary log with structural elements details
                DebugLogger.Log($"CableTraySleeveCommand summary: Total={totalTrays}, Intersections={intersectionCount}, Placed={placedCount}, Missing={missingCount}, SkippedExisting={skippedExistingCount}, StructuralDetected={structuralElementsDetected}, StructuralSleeves={structuralSleevesPlacer}");
                
                // Log structural summary to dedicated logger
                StructuralElementLogger.LogSummary("CableTraySleeveCommand", totalTrays, structuralElementsDetected, structuralSleevesPlacer, (structuralElementsDetected - structuralSleevesPlacer));
                StructuralElementLogger.LogStructuralElement("SYSTEM", new Autodesk.Revit.DB.ElementId((BuiltInParameter)0L), "COMMAND COMPLETED", $"Structural sleeve placement finished. Log file: {StructuralElementLogger.GetLogFilePath()}");
            }

            // Show status prompt to user
            string summary = $"CABLE TRAY SLEEVE SUMMARY: Total={totalTrays}, Placed={placedCount}, Missing={missingCount}, Skipped={skippedExistingCount}, Structural={structuralSleevesPlacer}";
            TaskDialog.Show("Cable Tray Sleeve Placement", summary);

            DebugLogger.Log("Cable tray sleeves placement completed.");
            return Result.Succeeded;
        }

        /// <summary>
        /// Places a cable tray wall sleeve at the specified location with duplication checking
        /// </summary>
        private bool PlaceCableTraySleeveAtLocation_Wall(Document doc, SleeveSpatialGrid sleeveGrid, FamilySymbol ctWallSymbol, Element hostElement, XYZ placementPoint, XYZ direction, double width, double height, ElementId trayId)
        {
            try
            {
                double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                double clusterExpansion = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);

                // If the host element lives in a linked document, transform the placement point
                // back into the host document coordinates before running duplication checks.
                XYZ placementPointHostCoords = placementPoint;
                try
                {
                    if (hostElement?.Document != null && hostElement.Document.IsLinked)
                    {
                        var allLinks = new FilteredElementCollector(doc)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>()
                            .ToList();
                        var linkInstance = allLinks.FirstOrDefault(l => l.GetLinkDocument() == hostElement.Document
                            || (l.GetLinkDocument()?.Title ?? string.Empty) == (hostElement.Document.Title ?? string.Empty)
                            || (l.GetLinkDocument()?.PathName ?? string.Empty) == (hostElement.Document.PathName ?? string.Empty));
                        if (linkInstance != null)
                        {
                            placementPointHostCoords = linkInstance.GetTotalTransform().OfPoint(placementPoint);
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[CableTraySleeveCommand] Warning transforming placement point for duplication check: {ex.Message}");
                }

                // --- Optimized duplicate suppression and logging ---
                string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placementPointHostCoords, sleeveCheckRadius);
                DebugLogger.Log($"[CableTraySleeveCommand] DUPLICATION CHECK at (hostCoords) {placementPointHostCoords} (link-local supplied: {placementPoint}):\n{sleeveSummary}");
                BoundingBoxXYZ? sectionBox = null;
                try
                {
                    if (doc.ActiveView is View3D vb)
                        sectionBox = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(vb);
                }
                catch { }

                string hostTypeFilter = hostElement is Wall ? "OpeningOnWall" : (hostElement is Floor ? "OpeningOnSlab" : "OpeningOnWall");
                DebugLogger.Log($"[CableTraySleeveCommand] Using optimized duplication checker hostType={hostTypeFilter}");

                var nearbySleeves = sleeveGrid.GetNearbySleeves(placementPointHostCoords, sleeveCheckRadius);
                bool duplicateExists = OpeningDuplicationChecker.IsAnySleeveAtLocationOptimized(placementPointHostCoords, sleeveCheckRadius, clusterExpansion, nearbySleeves, hostTypeFilter);

                if (duplicateExists)
                {
                    string msg = $"SKIP: CableTray {trayId.IntegerValue} duplicate sleeve (individual or cluster) exists near {placementPoint}";
                    DebugLogger.Log($"[CableTraySleeveCommand] {msg} (optimized)");
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
                            DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: Wall thickness parameter not found, using 200mm fallback");
                        }
                    }
                    DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: Wall thickness calculated: {UnitUtils.ConvertFromInternalUnits(wallDepth, UnitTypeId.Millimeters):F1}mm");

                    // Bounding box extent log for cable tray sleeve (wall)
                    double minX = placementPoint.X - width / 2.0;
                    double maxX = placementPoint.X + width / 2.0;
                    double minY = placementPoint.Y - height / 2.0;
                    double maxY = placementPoint.Y + height / 2.0;
                    DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: [BBOX-DEBUG] Placement at ({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3}), BBox X=({UnitUtils.ConvertFromInternalUnits(minX, UnitTypeId.Millimeters):F1}, {UnitUtils.ConvertFromInternalUnits(maxX, UnitTypeId.Millimeters):F1}), Y=({UnitUtils.ConvertFromInternalUnits(minY, UnitTypeId.Millimeters):F1}, {UnitUtils.ConvertFromInternalUnits(maxY, UnitTypeId.Millimeters):F1}), Width={UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, Height={UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm");

                    // Compute preCalculatedOrientation for wall: determine whether wall normal is X or Y oriented
                    XYZ? preCalcForWall = null;
                    try
                    {
                        // Attempt to compute wall normal from wall location curve if available
                        XYZ wallNormal = XYZ.BasisY;
                        try
                        {
                            var loc = wall.Location as LocationCurve;
                            if (loc != null && loc.Curve is Line ln)
                            {
                                // Wall normal is line direction crossed with Z (approx)
                                wallNormal = ln.Direction.CrossProduct(XYZ.BasisZ).Normalize();
                            }
                            else
                            {
                                // Fallback to using wall orientation by inspecting wall bounding box
                                var bbox = wall.get_BoundingBox(null);
                                if (bbox != null)
                                {
                                    double dx = Math.Abs(bbox.Max.X - bbox.Min.X);
                                    double dy = Math.Abs(bbox.Max.Y - bbox.Min.Y);
                                    wallNormal = dx > dy ? XYZ.BasisX : XYZ.BasisY;
                                }
                            }
                        }
                        catch { wallNormal = XYZ.BasisY; }

                        double absXn = Math.Abs(wallNormal.X);
                        double absYn = Math.Abs(wallNormal.Y);
                        // Only rotate for walls that are Y-oriented (wall runs along Y axis)
                        if (absYn > absXn)
                        {
                            preCalcForWall = XYZ.BasisY;
                            DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: Wall is Y-oriented - will request rotation (preCalc orientation Y)");
                        }
                        else
                        {
                            preCalcForWall = null; // do not rotate for X-oriented walls
                            DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: Wall is X-oriented - no rotation requested");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: Failed to compute wall orientation: {ex.Message}. No rotation requested");
                        preCalcForWall = null;
                    }

                    FamilyInstance? sleeveInstance = placer.PlaceCableTraySleeve((CableTray)null!, placementPoint, width, height, XYZ.BasisX, ctWallSymbol, wall, preCalcForWall);
                    if (sleeveInstance != null)
                    {
                        // --- Set Depth parameter as before ---
                        var depthParam = sleeveInstance.LookupParameter("Depth") ?? sleeveInstance.LookupParameter("d");
                        if (depthParam != null && depthParam.StorageType == StorageType.Double && !depthParam.IsReadOnly)
                        {
                            depthParam.Set(wallDepth);
                            DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: Set sleeve depth to wall thickness: {UnitUtils.ConvertFromInternalUnits(wallDepth, UnitTypeId.Millimeters):F0}mm");
                        }
                        else
                        {
                            DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: Sleeve depth parameter not found or not writable");
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"CableTray ID={(int)trayId.IntegerValue}: error placing wall sleeve: {ex.Message}");
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
                // If the host element lives in a linked document, transform the placement point
                // back into the host document coordinates before running duplication checks.
                XYZ placementPointHostCoords = placementPoint;
                try
                {
                    if (hostElement?.Document != null && hostElement.Document.IsLinked)
                    {
                        var allLinks = new FilteredElementCollector(doc)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>()
                            .ToList();
                        var linkInstance = allLinks.FirstOrDefault(l => l.GetLinkDocument() == hostElement.Document
                            || (l.GetLinkDocument()?.Title ?? string.Empty) == (hostElement.Document.Title ?? string.Empty)
                            || (l.GetLinkDocument()?.PathName ?? string.Empty) == (hostElement.Document.PathName ?? string.Empty));
                        if (linkInstance != null)
                        {
                            placementPointHostCoords = linkInstance.GetTotalTransform().OfPoint(placementPoint);
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[CableTraySleeveCommand] Warning transforming placement point for duplication check: {ex.Message}");
                }

                // --- Robust duplicate suppression and logging (optimized) ---
                string sleeveSummary = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placementPointHostCoords, sleeveCheckRadius);
                DebugLogger.Log($"[CableTraySleeveCommand] DUPLICATION CHECK at (hostCoords) {placementPointHostCoords} (link-local supplied: {placementPoint}):\n{sleeveSummary}");
                BoundingBoxXYZ? sectionBoxDoc = null;
                try { if (doc.ActiveView is View3D vb) sectionBoxDoc = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(vb); } catch { }
                string hostTypeFilter = hostElement is Wall ? "OpeningOnWall" : (hostElement is Floor ? "OpeningOnSlab" : "OpeningOnWall");
                DebugLogger.Log($"[CableTraySleeveCommand] Structural path duplication check using optimized checker hostType={hostTypeFilter}, sectionBoxProvided={(sectionBoxDoc!=null)}");
                bool duplicateExists = OpeningDuplicationChecker.IsAnySleeveAtLocationEnhanced(doc, placementPointHostCoords, sleeveCheckRadius, clusterExpansion: UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters), ignoreIds: null, hostType: hostTypeFilter, sectionBox: sectionBoxDoc);
                if (duplicateExists)
                {
                    string msg = $"SKIP: CableTray {tray.Id.IntegerValue} duplicate sleeve (individual or cluster) exists near {placementPoint}";
                    DebugLogger.Log($"[CableTraySleeveCommand] {msg} (optimized)");
                    return false;
                }

                var placer = new CableTraySleevePlacer(doc);
                // Guard: ensure hostElement is not null before passing into the placer
                if (hostElement == null)
                {
                    DebugLogger.Log($"[CableTraySleeveCommand] Cannot place structural sleeve: hostElement is NULL for CableTray {tray?.Id.IntegerValue.ToString() ?? "unknown"}");
                    return false;
                }
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
                    else if (hostElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
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
                DebugLogger.Log($"CableTray ID={(int)tray.Id.IntegerValue}: error placing structural sleeve: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Places a cable tray structural sleeve with width-based orientation for floors (similar to duct logic)
        /// </summary>
        private bool PlaceCableTraySleeveAtLocation_StructuralWithOrientation(
            Document doc,
            SleeveSpatialGrid sleeveGrid,
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
                int cableTrayId = (int)tray.Id.IntegerValue;
                DebugLogger.Log($"[CableTraySleeveCommand] ORIENTATION ANALYSIS for CableTray {cableTrayId}");
                
                // For framing, use location curve direction for orientation
                XYZ? preCalculatedOrientation = null;
                double widthToUse = width;
                double heightToUse = height;
                bool isFraming = hostElement is FamilyInstance famInst1 && famInst1.Category != null && famInst1.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming;
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
                    // For walls, compute orientation from wall geometry if not already set
                    if (hostElement is Wall && preCalculatedOrientation == null)
                    {
                        try
                        {
                            var theWall = hostElement as Wall;
                            XYZ wallNormal = XYZ.BasisY;
                            var loc = theWall?.Location as LocationCurve;
                            if (loc != null && loc.Curve is Line ln)
                            {
                                wallNormal = ln.Direction.CrossProduct(XYZ.BasisZ).Normalize();
                            }
                            else
                            {
                                var bb = theWall?.get_BoundingBox(null);
                                if (bb != null)
                                {
                                    double dx = Math.Abs(bb.Max.X - bb.Min.X);
                                    double dy = Math.Abs(bb.Max.Y - bb.Min.Y);
                                    wallNormal = dx > dy ? XYZ.BasisX : XYZ.BasisY;
                                }
                            }

                            double absX = Math.Abs(wallNormal.X);
                            double absY = Math.Abs(wallNormal.Y);
                            // Only request rotation for walls that are Y-oriented (wall runs along Y axis)
                            if (absY > absX)
                            {
                                preCalculatedOrientation = XYZ.BasisY;
                                DebugLogger.Log($"[CableTraySleeveCommand] [WALL] Wall is Y-oriented - requesting rotation (Y)");
                            }
                            else
                            {
                                preCalculatedOrientation = null; // do not rotate for X-oriented walls
                                DebugLogger.Log($"[CableTraySleeveCommand] [WALL] Wall is X-oriented - no rotation requested");
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Log($"[CableTraySleeveCommand] [WALL] Failed to compute wall orientation: {ex.Message}");
                        }
                    }
                }
                
                double sleeveCheckRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters);
                // --- Optimized duplicate suppression and logging ---
                string sleeveSummary2 = OpeningDuplicationChecker.GetSleevesSummaryAtLocation(doc, placementPoint, sleeveCheckRadius);
                DebugLogger.Log($"[CableTraySleeveCommand] DUPLICATION CHECK at {placementPoint}:\n{sleeveSummary2}");
                BoundingBoxXYZ? sectionBox2 = null;
                try
                {
                    if (doc.ActiveView is View3D vb2)
                        sectionBox2 = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(vb2);
                }
                catch { }

                string hostTypeFilter2 = hostElement is Wall ? "OpeningOnWall" : (hostElement is Floor ? "OpeningOnSlab" : "OpeningOnWall");
                DebugLogger.Log($"[CableTraySleeveCommand] Using optimized duplication checker hostType={hostTypeFilter2}");

                var nearbySleeves2 = sleeveGrid.GetNearbySleeves(placementPoint, sleeveCheckRadius);
                bool duplicateExists2 = OpeningDuplicationChecker.IsAnySleeveAtLocationOptimized(placementPoint, sleeveCheckRadius, UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters), nearbySleeves2, hostTypeFilter2);

                if (duplicateExists2)
                {
                    string msg = $"SKIP: CableTray {cableTrayId} duplicate sleeve (individual or cluster) exists near {placementPoint}";
                    DebugLogger.Log($"[CableTraySleeveCommand] {msg} (optimized)");
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
                else if (hostElement is FamilyInstance famInst2 && famInst2.Category != null && famInst2.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
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
                DebugLogger.Log($"CableTray ID={(int)tray.Id.IntegerValue}: error placing structural sleeve with orientation: {ex.Message}");
                return false;
            }
        }

        private static bool BoundingBoxesIntersect(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            if (a == null || b == null) return false;
            return !(a.Max.X < b.Min.X || a.Min.X > b.Max.X ||
                     a.Max.Y < b.Min.Y || a.Min.Y > b.Max.Y ||
                     a.Max.Z < b.Min.Z || a.Min.Z > b.Max.Z);
        }
    }
}
