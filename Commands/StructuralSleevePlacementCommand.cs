
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Plumbing; // For Pipe
using Autodesk.Revit.DB.Mechanical; // For Duct
using Autodesk.Revit.DB.Electrical; // For CableTray
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;


namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class StructuralSleevePlacementCommand : IExternalCommand
    {
        // Helper: Get the sleeve family symbol that would actually be placed for this MEP/structural pair
        private static FamilySymbol GetSleeveFamilySymbolForPlacement(Document doc, Element mepElement, Element structuralElement)
        {
            var allFamilySymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .ToList();

            if (allFamilySymbols.Count == 0)
            {
                System.IO.File.AppendAllText($"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log",
                    $"[CRITICAL][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] No FamilySymbols found in project!\n");
                return null;
            }

            // Strict: Only select the correct sleeve family symbol for placement, no fallback to any symbol
            FamilySymbol sleeveSymbol = null;

            // Determine MEP and structural types
            string mepType = mepElement?.GetType().Name ?? string.Empty;
            string structuralType = string.Empty;
            if (structuralElement != null && structuralElement.Category != null)
            {
                structuralType = structuralElement.Category.Name ?? string.Empty;
            }

            // Map of (MEP type, Structural Category Name) to strict family name
            // Update these mappings as needed to match your loaded families
            var familyMap = new Dictionary<(string mep, string structural), string>
            {
                // Pipe + Floor
                { ("Pipe", "Floors"), "PipeOpeningOnSlabRect" },
                // Pipe + Structural Framing
                { ("Pipe", "Structural Framing"), "PipeOpeningOnWallRect" },
                // Duct + Floor
                { ("Duct", "Floors"), "DuctOpeningOnSlabRect" },
                // Duct + Structural Framing
                { ("Duct", "Structural Framing"), "DuctOpeningOnWallRect" },
                // CableTray + Floor
                { ("CableTray", "Floors"), "CableTrayOpeningOnSlabRect" },
                // CableTray + Structural Framing
                { ("CableTray", "Structural Framing"), "CableTrayOpeningOnWallRect" },
                // Add more mappings as needed for your project
            };

            // Try to find a strict match in the map
            string strictFamilyName = null;
            if (familyMap.TryGetValue((mepType, structuralType), out strictFamilyName))
            {
                sleeveSymbol = allFamilySymbols.FirstOrDefault(sym => sym.Family.Name.Equals(strictFamilyName, StringComparison.OrdinalIgnoreCase));
                if (sleeveSymbol != null)
                {
                    System.IO.File.AppendAllText($"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log",
                        $"[DEBUG][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Strict match: MEP={mepType}, Structural={structuralType}, Family='{strictFamilyName}', Symbol='{sleeveSymbol.Name}'\n");
                }
                else
                {
                    System.IO.File.AppendAllText($"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log",
                        $"[CRITICAL][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] No FamilySymbol found for strict family name '{strictFamilyName}' (MEP={mepType}, Structural={structuralType})\n");
                }
            }
            else
            {
                System.IO.File.AppendAllText($"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log",
                    $"[CRITICAL][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] No strict mapping for MEP={mepType}, Structural={structuralType}. No sleeve will be placed.\n");
            }

            // If not found, log and return null (no fallback)
            if (sleeveSymbol == null)
            {
                System.IO.File.AppendAllText($"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log",
                    $"[CRITICAL][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] No matching sleeve family symbol found for placement.\n");
                return null;
            }
            return sleeveSymbol;
        }

// Remove duplicate namespace and class declarations


    /// <summary>
    /// Suppression Logic Type 1: Bounding Box Suppression for Rectangular Sleeves
    /// ---------------------------------------------------------------
    /// For rectangular sleeves (families like OnWallRect/OnSlabRect),
    /// this logic checks if a candidate point (potential sleeve location)
    /// is inside the transformed bounding box of any existing rectangular sleeve.
    /// - The candidate point is transformed into the local coordinate system of each sleeve (using the sleeve's transform).
    /// - The local point is compared to the local bounding box extents (with tolerance).
    /// - This approach is robust for large or rotated sleeves, ensuring that any point inside the sleeve's geometry is suppressed.
    /// - Used for rectangular sleeves where the sleeve's location point may not represent the full area covered by the sleeve.
    /// </summary>
    private static FamilyInstance IsSuppressedByRectBoundingBox(XYZ point, List<FamilyInstance> rectSleeves)
    {
        DebugLogger.Info($"[SUPPRESSION][TRACE] IsSuppressedByRectBoundingBox called: point={point}, rectSleeves.Count={(rectSleeves != null ? rectSleeves.Count.ToString() : "null")}");
        double tolCluster = 50.0; // 50mm for cluster sleeves
        double tolCandidate = 10.0; // 10mm for candidate sleeves (center-to-center)
        if (rectSleeves == null || rectSleeves.Count == 0)
        {
            DebugLogger.Info($"[SUPPRESSION][RECT] No rect sleeves to check for suppression at {point}");
            return null;
        }
        // IDs to debug in detail
        var debugElementIds = new HashSet<int> { 1935079,1935473,1935656,1944806,1944807,1946130,1946134,1946139 };
        string shortDebug = null;
        foreach (var rectSleeve in rectSleeves)
        {
            var bbox = rectSleeve.get_BoundingBox(null);
            if (bbox == null)
                continue;
            string famName = rectSleeve.Symbol?.Family?.Name ?? "";
            bool isCluster = famName.IndexOf("cluster", StringComparison.OrdinalIgnoreCase) >= 0;
            int sleeveId = rectSleeve.Id.IntegerValue;
            bool isDebug = debugElementIds.Contains(sleeveId);
            if (isDebug) {
                DebugLogger.Info($"[SUPPRESSION][SHORT] Checking {sleeveId} fam='{famName}' cluster={isCluster} point={point}");
            }
            // If cluster, use bbox center/corners with 50mm
            if (isCluster)
            {
                XYZ center = (bbox.Min + bbox.Max) / 2.0;
                double distToCenter = UnitUtils.ConvertFromInternalUnits(point.DistanceTo(center), UnitTypeId.Millimeters);
                if (isDebug)
                {
                    DebugLogger.Info($"[SUPPRESSION][SHORT]  - centerDist={distToCenter:F1}mm (<=50mm?)");
                }
                if (distToCenter <= tolCluster)
                {
                    if (isDebug)
                        DebugLogger.Info($"[SUPPRESSION][SHORT]  => YES (center)");
                    DebugLogger.Info($"[SUPPRESSION][RECT] Suppressed by CLUSTER bbox center: SleeveId={rectSleeve.Id}, Center={center}, Dist={distToCenter:F1}mm");
                    return rectSleeve;
                }
                var corners = new List<XYZ>
                {
                    new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                    new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                    new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                    new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                    new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                    new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                    new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                    new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z)
                };
                int cornerIdx = 0;
                foreach (var corner in corners)
                {
                    double distToCorner = UnitUtils.ConvertFromInternalUnits(point.DistanceTo(corner), UnitTypeId.Millimeters);
                    if (isDebug)
                    {
                        DebugLogger.Info($"[SUPPRESSION][SHORT]  - corner{cornerIdx}Dist={distToCorner:F1}mm (<=50mm?)");
                    }
                    if (distToCorner <= tolCluster)
                    {
                        if (isDebug)
                            DebugLogger.Info($"[SUPPRESSION][SHORT]  => YES (corner {cornerIdx})");
                        DebugLogger.Info($"[SUPPRESSION][RECT] Suppressed by CLUSTER bbox corner: SleeveId={rectSleeve.Id}, Corner={corner}, Dist={distToCorner:F1}mm");
                        return rectSleeve;
                    }
                    cornerIdx++;
                }
            }
            else
            {
                // Candidate sleeve: center-to-center check (10mm)
                var locPt = (rectSleeve.Location as LocationPoint)?.Point ?? rectSleeve.GetTransform().Origin;
                double distToCenter = UnitUtils.ConvertFromInternalUnits(point.DistanceTo(locPt), UnitTypeId.Millimeters);
                if (isDebug)
                {
                    DebugLogger.Info($"[SUPPRESSION][SHORT]  - centerDist={distToCenter:F1}mm (<=10mm?)");
                }
                if (distToCenter <= tolCandidate)
                {
                    if (isDebug)
                        DebugLogger.Info($"[SUPPRESSION][SHORT]  => YES");
                    DebugLogger.Info($"[SUPPRESSION][RECT] Suppressed by CANDIDATE center: SleeveId={rectSleeve.Id}, Center={locPt}, Dist={distToCenter:F1}mm");
                    return rectSleeve;
                }
            }
            if (isDebug)
            {
                DebugLogger.Info($"[SUPPRESSION][SHORT]  => NO");
            }
        }
        DebugLogger.Info($"[SUPPRESSION][RECT] Checked {rectSleeves.Count} rect sleeves, NO suppression at {point}");
        return null;
    }
    /// <summary>
    /// Suppression Logic Type 2: Center-to-Center Suppression for Standard (Round) Sleeves
    /// ---------------------------------------------------------------
    /// For standard (typically round) sleeves, this logic checks if a candidate point
    /// is within a small distance (10mm) of the center/placement point of any existing sleeve.
    /// - The candidate point is compared directly to the location point of each sleeve.
    /// - If the distance is less than or equal to 10mm, the point is considered suppressed.
    /// - This is appropriate for round sleeves where the location point is always at the center.
    /// - Used for non-rectangular sleeves (e.g., PipeOpeningOnSlab, DuctOpeningOnSlab, etc.).
    /// </summary>
    private static FamilyInstance IsSuppressedByExistingSleeve(XYZ point, Dictionary<FamilyInstance, XYZ> sleeveLocations)
    {
        const double tolMm = 10.0;
        int checkedCount = 0;
        // Center-to-center suppression logic for round sleeves only (NO LOGGING)
        foreach (var kvp in sleeveLocations)
        {
            var sleeve = kvp.Key;
            var existingPoint = kvp.Value;
            string thisType = GetRectSleeveTypeByFamilyName(sleeve.Symbol?.Family?.Name);
            // Only allow suppression by sleeves of the same type (wall/slab) or both non-rectangular
            if (!string.IsNullOrEmpty(thisType))
            {
                // Only suppress if candidate and existing are same rect type
                string candidateType = GetRectSleeveTypeByFamilyName(sleeve.Symbol?.Family?.Name);
                if (thisType != candidateType)
                    continue;
            }
            double distMm = UnitUtils.ConvertFromInternalUnits(point.DistanceTo(existingPoint), UnitTypeId.Millimeters);
            checkedCount++;
            if (distMm <= tolMm)
            {
                return sleeve;
            }
        }
        return null;
    }

    // Helper: Get the type of rectangular sleeve (wall/slab) from family name
    private static string GetRectSleeveTypeByFamilyName(string familyName)
    {
        if (string.IsNullOrEmpty(familyName)) return string.Empty;

        // Normalize for robust matching
        string name = familyName.ToLowerInvariant();

        // Recognize cluster/rectangular sleeves: ends with Rect, or contains 'cluster' and 'wall'/'slab'
        if (name.EndsWith("wallrect") || name.EndsWith("rect") || (name.Contains("cluster") && name.Contains("wall"))) return "WALL";
        if (name.EndsWith("slabrect") || name.EndsWith("rect") || (name.Contains("cluster") && name.Contains("slab"))) return "SLAB";

        // Recognize standard sleeves: ends with Wall/Slab (but not Rect)
        if (name.EndsWith("wall")) return "WALL";
        if (name.EndsWith("slab")) return "SLAB";

        // Also recognize common typos and pluralization
        if (name.Contains("openingsonwallrect") || name.Contains("openinigsonwallrect")) return "WALL";
        if (name.Contains("openingsonslabrect") || name.Contains("openinigsonslabrect")) return "SLAB";

        return string.Empty;
    }
        // 🎯 TARGET: Track these 9 specific MEP elements that should intersect
        private static readonly HashSet<int> EXPECTED_MEP_ELEMENTS = new HashSet<int>
        {
            // FLOORS (2 elements):
            1892943, // Pipe → should intersect floor
            1933792, // Duct → should intersect floor
            
            // STRUCTURAL FRAMING (7 elements):
            1935283, // Cable Tray → should intersect framing
            1935473, // Cable Tray → should intersect framing  
            1934835, // Pipe → should intersect framing
            1935108, // Pipe → should intersect framing
            1935654, // Pipe → should intersect framing
            1935656, // Pipe → should intersect framing
            1935079  // Duct → should intersect framing
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // UNCONDITIONAL LOG: Confirm Execute is called and logger works
            string logPath = $"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log";
            try {
                // Overwrite the log file at the start of every run
                System.IO.File.WriteAllText(logPath, $"[BUILD][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Build started for StructuralSleevePlacementCommand\n");
            } catch (Exception ex) {
                Autodesk.Revit.UI.TaskDialog.Show("Log Error", $"Failed to write to log: {ex.Message}");
            }
            try {
                // Also log the build time to a separate file for build tracking
                string buildLogPath = $"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Build.log";
                System.IO.File.WriteAllText(buildLogPath, $"[BUILD][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Build started and log overwritten.\n");
            } catch { }
            DebugLogger.InitLogFileOverwrite("StructuralSleevePlacement_Debug"); // Ensure log file is overwritten
            DebugLogger.Info("Starting StructuralSleevePlacementCommand execution.");
            DebugLogger.Info($"🎯 TRACKING 9 EXPECTED MEP ELEMENTS: {string.Join(", ", EXPECTED_MEP_ELEMENTS)}");

            // --- Suppression logic helpers (must be in scope for all placement logic) ---

            // Use the main document for all suppression logic
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            var existingSleeves = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol != null && (
                    fi.Symbol.Family.Name.Contains("PipeOpeningOnSlab") ||
                    fi.Symbol.Family.Name.Contains("DuctOpeningOnSlab") ||
                    fi.Symbol.Family.Name.Contains("OnWall")))
                .ToList();

            var existingSleeveLocations = existingSleeves.ToDictionary(
                sleeve => sleeve,
                sleeve => (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin);

            try
            {
                // Restore: Collect all MEP and structural elements from active and visible linked documents (no view/crop filtering)
                DebugLogger.Info($"Collecting MEP elements from active and visible linked documents");
                var allMEPElements = MEPElementCollector.CollectMEPElements(doc);
                DebugLogger.Info($"Total MEP elements collected: {allMEPElements.Count}");

                DebugLogger.Info($"Collecting structural elements from active and visible linked documents");
                var structuralElements = CollectStructuralElements(doc);
                DebugLogger.Info($"Found {structuralElements.Count} structural elements.");

                // Activate required family symbols
                var allFamilySymbols = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .ToList();

                // Update filter to match family names based on README guidelines
                var familySymbol = allFamilySymbols
                    .FirstOrDefault(sym => sym.Family.Name.Contains("OnWall", StringComparison.OrdinalIgnoreCase) ||
                                           sym.Family.Name.Contains("PipeOpeningOnSlab", StringComparison.OrdinalIgnoreCase) ||
                                           sym.Family.Name.Contains("DuctOpeningOnSlab", StringComparison.OrdinalIgnoreCase));

                if (familySymbol == null)
                {
                    TaskDialog.Show("Error", "No suitable sleeve family found. Please ensure the required families are loaded as per the README guidelines.");
                    return Result.Failed;
                }

                using (var txActivate = new Transaction(doc, "Activate Structural Sleeve Symbol"))
                {
                    txActivate.Start();
                    if (!familySymbol.IsActive)
                        familySymbol.Activate();
                    txActivate.Commit();
                }

                // Select a non-template 3D view
                var view3D = new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate);

                if (view3D == null)
                {
                    TaskDialog.Show("Error", "No non-template 3D view found.");
                    return Result.Failed;
                }

                // Define structural filters for floors and framing
                ElementFilter structuralFilter = new LogicalOrFilter(
                    new ElementCategoryFilter(BuiltInCategory.OST_Floors),
                    new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming));


            // Use ReferenceIntersector to detect intersections with structural elements
            var refIntersector = new ReferenceIntersector(structuralFilter, FindReferenceTarget.Face, view3D)
            {
                FindReferencesInRevitLinks = true
            };

            // Filter for only truly intersecting elements using robust geometric validation
            var intersectingElements = new List<Element>();
            int intersectionCount = 0;
            int noIntersectionCount = 0;

            foreach (var mepElement in allMEPElements)
            {
                int mepId = mepElement.Id.IntegerValue;
                bool isTargetElement = EXPECTED_MEP_ELEMENTS.Contains(mepId);

                if (isTargetElement)
                {
                    DebugLogger.Info($"🎯 FOUND TARGET ELEMENT: {mepElement.Category.Name} {mepId}");
                }

                var locCurve = mepElement.Location as LocationCurve;
                if (locCurve?.Curve is Line line)
                {
                    // Filter for elements based on their orientation and target structural type
                    bool shouldProcess = true;
                    string orientationReason = "";

                    if (mepElement is Pipe)
                    {
                        var dir = line.Direction.Normalize();
                        bool isVertical = Math.Abs(dir.Z) >= 0.7; // Z-component significant
                        bool isHorizontal = Math.Abs(dir.Z) < 0.3; // Z-component minimal

                        // Check if we have floors or framing to intersect with
                        bool hasFloors = structuralElements.Any(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors);
                        bool hasFraming = structuralElements.Any(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming);

                        if (isVertical && hasFloors)
                        {
                            orientationReason = $"vertical pipe (Z={Math.Abs(dir.Z):F3}) → will check floor intersections";
                            shouldProcess = true;
                        }
                        else if (isHorizontal && hasFraming)
                        {
                            orientationReason = $"horizontal pipe (Z={Math.Abs(dir.Z):F3}) → will check structural framing intersections";
                            shouldProcess = true;
                        }
                        else if (isVertical && !hasFloors)
                        {
                            orientationReason = $"vertical pipe but no floors available";
                            shouldProcess = false;
                        }
                        else if (isHorizontal && !hasFraming)
                        {
                            orientationReason = $"horizontal pipe but no structural framing available";
                            shouldProcess = false;
                        }
                        else
                        {
                            orientationReason = $"diagonal pipe (Z={Math.Abs(dir.Z):F3}) → will check all structural elements";
                            shouldProcess = true; // Process diagonal pipes against all structural elements
                        }

                        if (shouldProcess)
                        {
                            DebugLogger.Info($"Processing pipe {mepElement.Id.IntegerValue}: {orientationReason}");
                        }
                        else
                        {
                            DebugLogger.Debug($"Skipping pipe {mepElement.Id.IntegerValue}: {orientationReason}");
                            continue;
                        }
                    }
                    // Note: Ducts and cable trays are processed regardless of orientation

                    bool foundIntersection = false;
                    string mepType = mepElement.GetType().Name;

                    // Filter structural elements based on MEP element orientation
                    var targetStructuralElements = new List<Element>();
                    if (mepElement is Pipe)
                    {
                        var dir = line.Direction.Normalize();
                        bool isVertical = Math.Abs(dir.Z) >= 0.7;
                        bool isHorizontal = Math.Abs(dir.Z) < 0.3;

                        if (isVertical)
                        {
                            // Vertical pipes should intersect floors
                            targetStructuralElements = structuralElements
                                .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                                .ToList();
                            DebugLogger.Debug($"Vertical pipe {mepElement.Id.IntegerValue}: targeting {targetStructuralElements.Count} floors");
                        }
                        else if (isHorizontal)
                        {
                            // Horizontal pipes should intersect structural framing
                            targetStructuralElements = structuralElements
                                .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                                .ToList();
                            DebugLogger.Debug($"Horizontal pipe {mepElement.Id.IntegerValue}: targeting {targetStructuralElements.Count} structural framing elements");
                        }
                        else
                        {
                            // Diagonal pipes - check against all structural elements
                            targetStructuralElements = structuralElements.ToList();
                            DebugLogger.Debug($"Diagonal pipe {mepElement.Id.IntegerValue}: targeting all {targetStructuralElements.Count} structural elements");
                        }
                    }
                    else
                    {
                        // Ducts and cable trays - check against all structural elements
                        targetStructuralElements = structuralElements.ToList();
                        DebugLogger.Debug($"{mepType} {mepElement.Id.IntegerValue}: targeting all {targetStructuralElements.Count} structural elements");
                    }

                    // First try EXACT working command intersection logic (adapted for structural elements)
                    foreach (var structuralElement in targetStructuralElements)
                    {
                        // Task 2: Check if intersection exists and calculate placement point
                        var sleeveLocation = CalculateSleeveLocation(mepElement, structuralElement, line, existingSleeveLocations);
                        if (sleeveLocation != null)
                        {
                            DebugLogger.Info($"CONFIRMED INTERSECTION: {mepType} {mepElement.Id.IntegerValue} intersects {structuralElement.Category.Name} Id={structuralElement.Id.IntegerValue} - Sleeve placement: {sleeveLocation}");
                            foundIntersection = true;
                            intersectionCount++;
                            break;
                        }
                    }

                    // Fallback: Ray casting with strict geometric validation (like working commands)
                    if (!foundIntersection)
                    {
                        var sampleFractions = new[] { 0.0, 0.25, 0.5, 0.75, 1.0 };

                        foreach (double t in sampleFractions)
                        {
                            var samplePt = line.Evaluate(t, true);
                            var elementDirection = line.Direction.Normalize();

                            var directions = new[]
                            {
                                elementDirection,                      // Forward along element
                                elementDirection.Negate(),             // Backward along element
                                XYZ.BasisZ,                           // Up (for floors above)
                                XYZ.BasisZ.Negate(),                  // Down (for floors below)
                                new XYZ(-elementDirection.Y, elementDirection.X, 0).Normalize(), // Perpendicular 1
                                new XYZ(elementDirection.Y, -elementDirection.X, 0).Normalize()  // Perpendicular 2
                            };

                            foreach (var direction in directions)
                            {
                                var hits = refIntersector.Find(samplePt, direction);
                                if (hits != null && hits.Count > 0)
                                {
                                    foreach (var hit in hits)
                                    {
                                        // Proximity check - filter out hits that are too far away (following working command pattern)
                                        if (hit.Proximity > 0.33) continue; // 0.33 feet ≈ 10cm, matching PipeSleeveCommand

                                        var reference = hit.GetReference();
                                        Element structuralElement = null;

                                        // Handle both linked and non-linked elements
                                        var linkInst = doc.GetElement(reference.ElementId) as RevitLinkInstance;
                                        if (linkInst != null)
                                        {
                                            var linkedDoc = linkInst.GetLinkDocument();
                                            if (linkedDoc != null)
                                            {
                                                structuralElement = linkedDoc.GetElement(reference.LinkedElementId);
                                            }
                                        }
                                        else
                                        {
                                            structuralElement = doc.GetElement(reference.ElementId);
                                        }

                                        if (structuralElement != null && 
                                            (structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors ||
                                             structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming))
                                        {
                                            // Final validation using EXACT working command pattern
                                            if (ValidateStructuralIntersectionExact(mepElement, structuralElement, line))
                                            {
                                                DebugLogger.Info($"VALIDATED: {mepType} {mepElement.Id.IntegerValue} intersects {structuralElement.Category.Name} " +
                                                               $"Id={structuralElement.Id.IntegerValue} at proximity {hit.Proximity:F3}. Link: {linkInst != null}");
                                                foundIntersection = true;
                                                intersectionCount++;
                                                break;
                                            }
                                            else
                                            {
                                                DebugLogger.Debug($"REJECTED: {mepType} {mepElement.Id.IntegerValue} ray hit {structuralElement.Category.Name} " +
                                                                $"Id={structuralElement.Id.IntegerValue} at proximity {hit.Proximity:F3} but geometric validation failed");
                                            }
                                        }
                                    }

                                    if (foundIntersection) break;
                                }
                            }

                            if (foundIntersection) break;
                        }
                    }

                    if (foundIntersection)
                    {
                        intersectingElements.Add(mepElement);
                        DebugLogger.Info($"Intersection FOUND for MEP element: {mepType} Id={mepElement.Id.IntegerValue}");
                    }
                    else
                    {
                        noIntersectionCount++;
                        DebugLogger.Info($"No intersection found for MEP element: {mepType} Id={mepElement.Id.IntegerValue}");
                    }
                }
            }

            DebugLogger.Info($"Total intersecting MEP elements found: {intersectingElements.Count} (from {intersectionCount} intersection hits).");
            DebugLogger.Info($"Total MEP elements with NO intersection: {noIntersectionCount}");

            // 🎯 SUMMARY: Check which of the 9 expected elements were found
            var foundTargetElements = intersectingElements.Where(e => EXPECTED_MEP_ELEMENTS.Contains(e.Id.IntegerValue)).ToList();
            var missingTargetElements = EXPECTED_MEP_ELEMENTS.Except(intersectingElements.Select(e => e.Id.IntegerValue)).ToList();

            DebugLogger.Info($"🎯 TARGET ELEMENTS SUMMARY:");
            DebugLogger.Info($"   ✅ Found {foundTargetElements.Count}/9 expected elements: {string.Join(", ", foundTargetElements.Select(e => e.Id.IntegerValue))}");
            DebugLogger.Info($"   ❌ Missing {missingTargetElements.Count}/9 expected elements: {string.Join(", ", missingTargetElements)}");

            // Log details of intersecting MEP elements
            foreach (var element in intersectingElements)
            {
                string elementType = element.GetType().Name;
                string elementId = element.Id.ToString();
                string location = (element.Location as LocationCurve)?.Curve.GetEndPoint(0).ToString() ?? "Unknown";
                DebugLogger.Info($"Intersecting MEP element: {elementType} Id={elementId} at {location}");
            }

                // Collect existing sleeves based on the correct family names from README

                // (Removed duplicate/inner declarations of suppression logic and sleeve collections)

                if (existingSleeves.Count > 0)
                {
                    DebugLogger.Info($"Found {existingSleeves.Count} existing sleeves.");

                    // Log existing sleeves details
                    foreach (var sleeve in existingSleeves)
                    {
                        string familyName = sleeve.Symbol.Family.Name;
                        string sleeveId = sleeve.Id.ToString();
                        string sleeveLocation = existingSleeveLocations[sleeve]?.ToString() ?? "Unknown";
                        DebugLogger.Info($"Found existing sleeve: {familyName} Id={sleeveId} at {sleeveLocation}");
                    }
                }
                else
                {
                    DebugLogger.Info("No existing sleeves found.");
                }

                // Test structural element collection
                var testStructuralElements = CollectStructuralElements(doc);
                DebugLogger.Info($"Test: Found {testStructuralElements.Count} total structural elements.");
                
                // Log breakdown by type
                var floorCount = testStructuralElements.Count(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors);
                var framingCount = testStructuralElements.Count(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming);
                DebugLogger.Info($"Test: {floorCount} floors, {framingCount} structural framing elements.");

                DebugLogger.Info("StructuralSleevePlacementCommand execution completed successfully.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error during StructuralSleevePlacementCommand execution: {ex.Message}");
                message = ex.Message;
                return Result.Failed;
            }
        }

        public static List<Element> CollectStructuralElements(Document doc)
        {
            try
            {
                DebugLogger.Debug("Starting structural element collection from active document and linked files.");

                var structuralElements = new List<Element>();

                // Collect Floor elements from active document
                var floors = new FilteredElementCollector(doc)
                    .OfClass(typeof(Floor))
                    .WhereElementIsNotElementType()
                    .ToList();
                structuralElements.AddRange(floors);
                DebugLogger.Debug($"Collected {floors.Count} Floor elements from the active document.");

                // Collect ALL Structural Framing elements from active document (regardless of host)
                // NOTE: We use unhosted generic model families for sleeves, as per project standard and repeated user instruction.
                // (User has requested this many times: generic model sleeves are NOT hosted, do NOT filter by Host!)
                var framingElements = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                    .Cast<FamilyInstance>()
                    .ToList();
                structuralElements.AddRange(framingElements);
                DebugLogger.Debug($"Collected {framingElements.Count} Structural Framing elements from the active document (hosted and unhosted).");

                // Collect structural elements from linked documents (VISIBLE ONLY)
                var linkedInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();

                foreach (var linkInstance in linkedInstances)
                {
                    // FILTER: Only process VISIBLE linked models
                    if (!linkInstance.IsHidden(doc.ActiveView))
                    {
                        var linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            var linkedElements = new List<Element>();

                            // Collect Floor elements from linked document
                            var linkedFloors = new FilteredElementCollector(linkedDoc)
                                .OfClass(typeof(Floor))
                                .WhereElementIsNotElementType()
                                .ToList();
                            linkedElements.AddRange(linkedFloors);

                            // Collect ALL Structural Framing elements from linked document (regardless of host)
                            // NOTE: We use unhosted generic model families for sleeves, as per project standard and repeated user instruction.
                            var linkedFraming = new FilteredElementCollector(linkedDoc)
                                .OfClass(typeof(FamilyInstance))
                                .WhereElementIsNotElementType()
                                .Where(e => e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                                .Cast<FamilyInstance>()
                                .ToList();
                            linkedElements.AddRange(linkedFraming);

                            structuralElements.AddRange(linkedElements);
                            DebugLogger.Debug($"Collected {linkedElements.Count} structural elements from VISIBLE linked document: {linkedDoc.Title}.");
                        }
                    }
                    else
                    {
                        DebugLogger.Debug($"Skipped HIDDEN linked document: {linkInstance.Name}");
                    }
                }

                DebugLogger.Debug($"Total structural elements collected: {structuralElements.Count}.");
                return structuralElements;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error collecting structural elements: {ex.Message}");
                return new List<Element>();
            }
        }



        public static List<IntersectionResult> FindIntersections(List<Element> mepElements, List<Element> structuralElements)
        {
            try
            {
                DebugLogger.Debug("Starting intersection detection.");

                var intersections = new List<IntersectionResult>();
                double proximityThreshold = 0.1; // Example threshold in feet

                foreach (var mepElement in mepElements)
                {
                    foreach (var structuralElement in structuralElements)
                    {
                        // Get bounding boxes
                        var boundingBoxMEP = mepElement.get_BoundingBox(null);
                        var boundingBoxStructural = structuralElement.get_BoundingBox(null);

                        if (boundingBoxMEP != null && boundingBoxStructural != null)
                        {
                            // Check for overlap between bounding boxes
                            if (boundingBoxMEP.Min.X <= boundingBoxStructural.Max.X && boundingBoxMEP.Max.X >= boundingBoxStructural.Min.X &&
                                boundingBoxMEP.Min.Y <= boundingBoxStructural.Max.Y && boundingBoxMEP.Max.Y >= boundingBoxStructural.Min.Y &&
                                boundingBoxMEP.Min.Z <= boundingBoxStructural.Max.Z && boundingBoxMEP.Max.Z >= boundingBoxStructural.Min.Z)
                            {
                                // Calculate proximity
                                double distance = boundingBoxMEP.Min.DistanceTo(boundingBoxStructural.Min);
                                if (distance <= proximityThreshold)
                                {
                                    intersections.Add(new IntersectionResult
                                    {
                                        MEPElement = mepElement,
                                        StructuralElement = structuralElement,
                                        IntersectionPoint = boundingBoxMEP.Min // Example point
                                    });

                                    DebugLogger.Debug($"Intersection added: MEP Element ID {mepElement.Id}, Structural Element ID {structuralElement.Id}, Distance: {distance}");
                                }
                                else
                                {
                                    DebugLogger.Debug($"Filtered out due to proximity: MEP Element ID {mepElement.Id}, Structural Element ID {structuralElement.Id}, Distance: {distance}");
                                }
                            }
                            else
                            {
                                DebugLogger.Debug($"No bounding box overlap: MEP Element ID {mepElement.Id}, Structural Element ID {structuralElement.Id}");
                            }
                        }
                        else
                        {
                            DebugLogger.Debug($"Null bounding box detected: MEP Element ID {mepElement.Id}, Structural Element ID {structuralElement.Id}");
                        }
                    }
                }

                DebugLogger.Debug($"Found {intersections.Count} intersections.");

                // Log intersection detection results
                DebugLogger.LogDetailed($"Intersection detection completed - {(intersections.Count > 0 ? "Intersections found." : "No intersections detected.")}");

                return intersections;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error finding intersections: {ex.Message}");
                return new List<IntersectionResult>();
            }
        }

        /// <summary>
        /// Validates actual geometric intersection using solid operations - matching working command pattern
        /// </summary>
        public static bool ValidateGeometricIntersection(Element mepElement, List<Element> structuralElements, Document doc)
        {
            try
            {
                // Get the solid geometry of the MEP element
                var mepOptions = new Options();
                mepOptions.DetailLevel = ViewDetailLevel.Fine;
                mepOptions.IncludeNonVisibleObjects = false;
                
                var mepGeometry = mepElement.get_Geometry(mepOptions);
                if (mepGeometry == null) return false;
                
                Solid mepSolid = null;
                foreach (GeometryObject geomObj in mepGeometry)
                {
                    if (geomObj is Solid solid && solid.Volume > 0)
                    {
                        mepSolid = solid;
                        break;
                    }
                    else if (geomObj is GeometryInstance instance)
                    {
                        foreach (GeometryObject instObj in instance.GetInstanceGeometry())
                        {
                            if (instObj is Solid instSolid && instSolid.Volume > 0)
                            {
                                mepSolid = instSolid;
                                break;
                            }
                        }
                        if (mepSolid != null) break;
                    }
                }
                
                if (mepSolid == null || mepSolid.Volume <= 0) return false;
                
                // Check against each structural element
                foreach (var structuralElement in structuralElements)
                {
                    var structOptions = new Options();
                    structOptions.DetailLevel = ViewDetailLevel.Fine;
                    structOptions.IncludeNonVisibleObjects = false;
                    
                    var structGeometry = structuralElement.get_Geometry(structOptions);
                    if (structGeometry == null) continue;
                    
                    foreach (GeometryObject geomObj in structGeometry)
                    {
                        Solid structSolid = null;
                        if (geomObj is Solid solid && solid.Volume > 0)
                        {
                            structSolid = solid;
                        }
                        else if (geomObj is GeometryInstance instance)
                        {
                            foreach (GeometryObject instObj in instance.GetInstanceGeometry())
                            {
                                if (instObj is Solid instSolid && instSolid.Volume > 0)
                                {
                                    structSolid = instSolid;
                                    break;
                                }
                            }
                        }
                        
                        if (structSolid != null && structSolid.Volume > 0)
                        {
                            try
                            {
                                // Perform boolean intersection to check for actual overlap
                                var intersection = BooleanOperationsUtils.ExecuteBooleanOperation(
                                    mepSolid, structSolid, BooleanOperationsType.Intersect);
                                
                                if (intersection != null && intersection.Volume > 1e-6) // Small tolerance for numerical precision
                                {
                                    DebugLogger.Info($"GEOMETRIC VALIDATION: MEP {mepElement.Id.IntegerValue} actually intersects " +
                                                   $"structural {structuralElement.Id.IntegerValue} with intersection volume {intersection.Volume:F6}");
                                    return true;
                                }
                            }
                            catch (Exception ex)
                            {
                                // Boolean operations can fail, fall back to distance check
                                DebugLogger.LogDetailed($"Boolean operation failed - MEP {mepElement.Id.IntegerValue} vs Structural {structuralElement.Id.IntegerValue}: {ex.Message}");
                                
                                // Fallback: check if solids are very close
                                var mepBounds = mepSolid.GetBoundingBox();
                                var structBounds = structSolid.GetBoundingBox();
                                
                                if (mepBounds != null && structBounds != null)
                                {
                                    var distance = mepBounds.Min.DistanceTo(structBounds.Max);
                                    if (distance < 0.01) // Very close tolerance
                                    {
                                        DebugLogger.Info($"PROXIMITY VALIDATION: MEP {mepElement.Id.IntegerValue} very close to " +
                                                       $"structural {structuralElement.Id.IntegerValue} (distance: {distance:F3})");
                                        return true;
                                    }
                                }
                            }
                        }
                    }
                }
                
                return false;
            }
            catch (Exception ex)
            {
                DebugLogger.LogDetailed($"Geometric validation failed - Element {mepElement.Id}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// TASK 2: Calculate sleeve placement point as midpoint between entry and exit faces
        /// Following the working command pattern from PipeSleeveCommand/DuctSleeveCommand
        /// </summary>
        private XYZ CalculateSleeveLocation(Element mepElement, Element structuralElement, Line line, Dictionary<FamilyInstance, XYZ> existingSleeveLocations)
            // (Removed misplaced log line)
        {
            try
            {
                // ONE-LINE DEBUG: Log every call to CalculateSleeveLocation
                System.IO.File.AppendAllText($"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log", $"[DEBUG][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Entered CalculateSleeveLocation for MEP {mepElement.Id} ({mepElement.GetType().Name}) and Structural {structuralElement.Id} ({structuralElement.GetType().Name})\n");
                DebugLogger.Debug($"TASK 2: Calculating sleeve location for {mepElement.Category.Name} {mepElement.Id.IntegerValue} and {structuralElement.Category.Name} {structuralElement.Id.IntegerValue}");

            // --- Only log the necessary details for debugging ---
            var doc = mepElement.Document;

                // --- Use the actual sleeve symbol that would be placed ---


                // For placement, use the selected sleeve symbol as before
                var sleeveSymbol = GetSleeveFamilySymbolForPlacement(doc, mepElement, structuralElement);
                string sleeveFamilyName = sleeveSymbol?.Family?.Name ?? string.Empty;
                if (sleeveSymbol == null)
                {
                    System.IO.File.AppendAllText($"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log",
                        $"[CRITICAL][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] CalculateSleeveLocation: sleeveSymbol is NULL for MEP={mepElement.Id} ({mepElement.GetType().Name}), Structural={structuralElement.Id} ({structuralElement.GetType().Name})\n");
                    return null;
                }
                if (sleeveSymbol.Family == null || string.IsNullOrEmpty(sleeveSymbol.Family.Name))
                {
                    System.IO.File.AppendAllText($"C:\\Users\\{Environment.UserName}\\Documents\\DuplicationSuppression_Debug.log",
                        $"[CRITICAL][{DateTime.Now:yyyy-MM-dd HH:mm:ss}] CalculateSleeveLocation: sleeveSymbol has EMPTY family name! Symbol='{sleeveSymbol.Name}' for MEP={mepElement.Id} ({mepElement.GetType().Name}), Structural={structuralElement.Id} ({structuralElement.GetType().Name})\n");
                    return null;
                }
                // For suppression, use the candidate sleeve family name only (not MEP element)
                bool isWallRectFamily = sleeveFamilyName.EndsWith("OnWallRect", StringComparison.OrdinalIgnoreCase);
                bool isSlabRectFamily = sleeveFamilyName.EndsWith("OnSlabRect", StringComparison.OrdinalIgnoreCase);
                bool isRectFamily = isWallRectFamily || isSlabRectFamily;
                // Suppression logic should NOT reference or log structural elements
                DebugLogger.Info($"[SUPPRESSION] Family: '{sleeveFamilyName}', isRectFamily={isRectFamily}, isWallRectFamily={isWallRectFamily}, isSlabRectFamily={isSlabRectFamily}");
                // Only collect and log sleeves, not structural elements
                List<FamilyInstance> rectSleeves = new List<FamilyInstance>();
                if (isWallRectFamily)
                {
                    rectSleeves = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol != null && fi.Symbol.Family.Name.EndsWith("OnWallRect", StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }
                else if (isSlabRectFamily)
                {
                    rectSleeves = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => fi.Symbol != null && fi.Symbol.Family.Name.EndsWith("OnSlabRect", StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }
                DebugLogger.Info($"[SUPPRESSION] Existing RECT sleeves found: {rectSleeves.Count}");
                foreach (var rect in rectSleeves)
                {
                    DebugLogger.Debug($"[SUPPRESSION] Existing RECT sleeve: {rect.Symbol.Family.Name} (ID: {rect.Id}) at location: {(rect.Location as LocationPoint)?.Point ?? rect.GetTransform().Origin}");
                }

                // Step 1: Extract structural element solid (like working command)
                // ...existing code...
                Solid structuralSolid = null;
                Options geomOptions = new Options { ComputeReferences = true, IncludeNonVisibleObjects = false };
                // ...existing code...
                Transform linkTransform = null;
                bool isLinkedElement = false;
                // ...existing code...
                var linkedInstances = new FilteredElementCollector(mepElement.Document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();
                foreach (var linkInstance in linkedInstances)
                {
                    var linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc != null)
                    {
                        try
                        {
                            var testElement = linkedDoc.GetElement(structuralElement.Id);
                            if (testElement != null)
                            {
                                isLinkedElement = true;
                                linkTransform = linkInstance.GetTotalTransform();
                                DebugLogger.Debug($"Found structural element {structuralElement.Id.IntegerValue} in linked document {linkedDoc.Title} with transform");
                                break;
                            }
                        }
                        catch
                        {
                            // Element not in this linked document, continue searching
                        }
                    }
                }
                GeometryElement geomElem = structuralElement.get_Geometry(geomOptions);
                foreach (GeometryObject obj in geomElem)
                {
                    if (obj is Solid solid && solid.Volume > 0)
                    {
                        // Apply transformation if this is a linked element
                        if (isLinkedElement && linkTransform != null)
                        {
                            structuralSolid = SolidUtils.CreateTransformed(solid, linkTransform);
                            DebugLogger.Debug($"Applied link transformation to structural solid. Original volume: {solid.Volume:F3}, Transformed volume: {structuralSolid.Volume:F3}");
                        }
                        else
                        {
                            structuralSolid = solid;
                        }
                        break;
                    }
                    else if (obj is GeometryInstance geomInstance)
                    {
                        // Handle geometry instances (common in family instances)
                        foreach (GeometryObject instObj in geomInstance.GetInstanceGeometry())
                        {
                            if (instObj is Solid instSolid && instSolid.Volume > 0)
                            {
                                // Apply both instance transform and link transform if needed
                                Transform combinedTransform = geomInstance.Transform;
                                if (isLinkedElement && linkTransform != null)
                                {
                                    combinedTransform = linkTransform.Multiply(geomInstance.Transform);
                                }
                                structuralSolid = SolidUtils.CreateTransformed(instSolid, combinedTransform);
                                DebugLogger.Debug($"Applied combined transformation to geometry instance solid. Volume: {structuralSolid.Volume:F3}");
                                break;
                            }
                        }
                        if (structuralSolid != null) break;
                    }
                }
                if (structuralSolid == null)
                {
                    DebugLogger.Debug($"No solid geometry found for {structuralElement.Category.Name} {structuralElement.Id.IntegerValue}");
                    return null;
                }
                // Step 2: Find entry/exit points through structural element (EXACT copy of working command)
                List<XYZ> intersectionPoints = new List<XYZ>();
                foreach (Face face in structuralSolid.Faces)
                {
                    IntersectionResultArray ira = null;
                    SetComparisonResult res = face.Intersect(line, out ira);
                    if (res == SetComparisonResult.Overlap && ira != null)
                    {
                        foreach (Autodesk.Revit.DB.IntersectionResult ir in ira)
                        {
                            intersectionPoints.Add(ir.XYZPoint);
                        }
                    }
                }
                DebugLogger.Debug($"Found {intersectionPoints.Count} face intersection points");
            // --- DEBUG: Log all intersection points found ---
            DebugLogger.Info($"[DEBUG] CalculateSleeveLocation: MEP {mepElement.Id.IntegerValue}, Structural {structuralElement.Id.IntegerValue}, intersectionPoints: " +
                string.Join(", ", intersectionPoints.Select(pt => pt.ToString())));

                // Step 3: Calculate midpoint between entry and exit points (working command pattern)
                if (intersectionPoints.Count >= 2)
                {
                    // Sort intersection points along the MEP line parameter
                    var sortedPoints = intersectionPoints
                        .Select(pt => new { Point = pt, Parameter = line.Project(pt).Parameter })
                        .OrderBy(x => x.Parameter)
                        .ToList();
                    XYZ entryPoint = sortedPoints.First().Point;
                    XYZ exitPoint = sortedPoints.Last().Point;
                    XYZ midpoint = (entryPoint + exitPoint) / 2.0;
                    DebugLogger.Info($"[DEBUG] CalculateSleeveLocation: Entry={entryPoint}, Exit={exitPoint}, Midpoint={midpoint}");
                    DebugLogger.Info($"TASK 2 SUCCESS: Entry point: {entryPoint}, Exit point: {exitPoint}, Midpoint: {midpoint}");
                    // Suppression: skip if an existing sleeve is already at this point (within 10mm)
                    FamilyInstance suppressor = null;
                    if (isRectFamily)
                    {
                        DebugLogger.Info($"[SUPPRESSION][PLACEMENT] About to check RECT suppression at midpoint: {midpoint} with {rectSleeves.Count} rect sleeves");
                        foreach (var rs in rectSleeves)
                        {
                            var bbox = rs.get_BoundingBox(null);
                            var loc = (rs.Location as LocationPoint)?.Point ?? rs.GetTransform().Origin;
                            DebugLogger.Info($"[SUPPRESSION][PLACEMENT] RectSleeve: Id={rs.Id}, Family={rs.Symbol?.Family?.Name}, Symbol={rs.Symbol?.Name}, Location={loc}, BBoxMin={bbox?.Min}, BBoxMax={bbox?.Max}");
                        }
                        suppressor = IsSuppressedByRectBoundingBox(midpoint, rectSleeves);
                        if (suppressor == null)
                        {
                            DebugLogger.Info($"[SUPPRESSION][PLACEMENT] No RECT suppression at midpoint: {midpoint}");
                        }
                        else
                        {
                            DebugLogger.Info($"[SUPPRESSION][PLACEMENT] SUPPRESSED: Rect sleeve bounding box found at {midpoint}, skipping placement. Existing: {suppressor.Symbol.Family.Name} (ID: {suppressor.Id})");
                            return midpoint;
                        }
                    }
                    else
                    {
                        DebugLogger.Info($"[SUPPRESSION] Checking standard center-to-center suppression at midpoint: {midpoint}");
                        suppressor = IsSuppressedByExistingSleeve(midpoint, existingSleeveLocations);
                        if (suppressor == null)
                        {
                            DebugLogger.Debug($"[SUPPRESSION] No center-to-center suppression at midpoint: {midpoint}");
                        }
                        else
                        {
                            DebugLogger.Info($"SUPPRESSED: Existing sleeve found at {midpoint} (within 10mm), skipping placement. Existing: {suppressor.Symbol.Family.Name} (ID: {suppressor.Id})");
                            return midpoint;
                        }
                    }
                    // TASK 2: Place appropriate sleeve family at midpoint
                    SleevePlacer.PlaceSleeveAtMidpoint(mepElement, structuralElement, midpoint);
                    return midpoint;
                }
                else if (intersectionPoints.Count == 1)
                {
                    // Single intersection point - use as placement location
                    DebugLogger.Info($"[DEBUG] CalculateSleeveLocation: Single intersection point: {intersectionPoints[0]}");
                    DebugLogger.Debug($"Single intersection point found: {intersectionPoints[0]}");
                    // Suppression: skip if an existing sleeve is already at this point (within 10mm)
                    FamilyInstance suppressor = null;
                    if (isRectFamily)
                    {
                        DebugLogger.Info($"[SUPPRESSION][PLACEMENT] About to check RECT suppression at intersection point: {intersectionPoints[0]} with {rectSleeves.Count} rect sleeves");
                        foreach (var rs in rectSleeves)
                        {
                            var bbox = rs.get_BoundingBox(null);
                            var loc = (rs.Location as LocationPoint)?.Point ?? rs.GetTransform().Origin;
                            DebugLogger.Info($"[SUPPRESSION][PLACEMENT] RectSleeve: Id={rs.Id}, Family={rs.Symbol?.Family?.Name}, Symbol={rs.Symbol?.Name}, Location={loc}, BBoxMin={bbox?.Min}, BBoxMax={bbox?.Max}");
                        }
                        suppressor = IsSuppressedByRectBoundingBox(intersectionPoints[0], rectSleeves);
                        if (suppressor == null)
                        {
                            DebugLogger.Info($"[SUPPRESSION][PLACEMENT] No RECT suppression at intersection point: {intersectionPoints[0]}");
                        }
                        else
                        {
                            DebugLogger.Info($"[SUPPRESSION][PLACEMENT] SUPPRESSED: Rect sleeve bounding box found at {intersectionPoints[0]}, skipping placement. Existing: {suppressor.Symbol.Family.Name} (ID: {suppressor.Id})");
                            return intersectionPoints[0];
                        }
                    }
                    else
                    {
                        DebugLogger.Info($"[SUPPRESSION] Checking standard center-to-center suppression at intersection point: {intersectionPoints[0]}");
                        suppressor = IsSuppressedByExistingSleeve(intersectionPoints[0], existingSleeveLocations);
                        if (suppressor == null)
                        {
                            DebugLogger.Debug($"[SUPPRESSION] No center-to-center suppression at intersection point: {intersectionPoints[0]}");
                        }
                        else
                        {
                            DebugLogger.Info($"SUPPRESSED: Existing sleeve found at {intersectionPoints[0]} (within 10mm), skipping placement. Existing: {suppressor.Symbol.Family.Name} (ID: {suppressor.Id})");
                            return intersectionPoints[0];
                        }
                    }
                    // TASK 2: Place appropriate sleeve family at single intersection point
                    SleevePlacer.PlaceSleeveAtMidpoint(mepElement, structuralElement, intersectionPoints[0]);
                    return intersectionPoints[0];
                }
                else
                {
                    DebugLogger.Debug("No face intersections found - trying ray casting fallback");
                    // Fallback: Use ray casting to find intersection point (if geometric method fails)
                    var view3D = new FilteredElementCollector(structuralElement.Document)
                        .OfClass(typeof(View3D))
                        .FirstOrDefault(v => !((View3D)v).IsTemplate) as View3D;
                    if (view3D != null)
                    {
                        var structuralFilter = new ElementCategoryFilter(structuralElement.Category.Id);
                        var refIntersector = new ReferenceIntersector(structuralFilter, FindReferenceTarget.Element, view3D);
                        // Try multiple sample points along the line
                        var sampleFractions = new[] { 0.25, 0.5, 0.75 };
                        foreach (var fraction in sampleFractions)
                        {
                            XYZ samplePoint = line.Evaluate(fraction, true);
                            XYZ direction = line.Direction.Normalize();
                            var hits = refIntersector.Find(samplePoint, direction);
                            var relevantHit = hits?.FirstOrDefault(h => h.GetReference().ElementId == structuralElement.Id);
                            if (relevantHit != null && relevantHit.Proximity <= 0.25)
                            {
                                DebugLogger.Info($"TASK 2 FALLBACK: Ray casting found intersection at {relevantHit.GetReference().GlobalPoint}");
                                // Suppression: skip if an existing sleeve is already at this point (within 10mm)
                                FamilyInstance suppressor = null;
                                if (isRectFamily)
                                {
                                    DebugLogger.Info($"[SUPPRESSION] Checking RECT bounding box suppression at fallback point: {relevantHit.GetReference().GlobalPoint}");
                                    suppressor = IsSuppressedByRectBoundingBox(relevantHit.GetReference().GlobalPoint, rectSleeves);
                                    if (suppressor == null)
                                    {
                                        DebugLogger.Debug($"[SUPPRESSION] No RECT suppression at fallback point: {relevantHit.GetReference().GlobalPoint}");
                                    }
                                    else
                                    {
                                        DebugLogger.Info($"SUPPRESSED: Rect sleeve bounding box found at {relevantHit.GetReference().GlobalPoint}, skipping placement. Existing: {suppressor.Symbol.Family.Name} (ID: {suppressor.Id})");
                                        return relevantHit.GetReference().GlobalPoint;
                                    }
                                }
                                else
                                {
                                    DebugLogger.Info($"[SUPPRESSION] Checking standard center-to-center suppression at fallback point: {relevantHit.GetReference().GlobalPoint}");
                                    suppressor = IsSuppressedByExistingSleeve(relevantHit.GetReference().GlobalPoint, existingSleeveLocations);
                                    if (suppressor == null)
                                    {
                                        DebugLogger.Debug($"[SUPPRESSION] No center-to-center suppression at fallback point: {relevantHit.GetReference().GlobalPoint}");
                                    }
                                    else
                                    {
                                        DebugLogger.Info($"SUPPRESSED: Existing sleeve found at {relevantHit.GetReference().GlobalPoint} (within 10mm), skipping placement. Existing: {suppressor.Symbol.Family.Name} (ID: {suppressor.Id})");
                                        return relevantHit.GetReference().GlobalPoint;
                                    }
                                }
                                // TASK 2: Place appropriate sleeve family at ray casting hit point
                                SleevePlacer.PlaceSleeveAtMidpoint(mepElement, structuralElement, relevantHit.GetReference().GlobalPoint);
                                return relevantHit.GetReference().GlobalPoint;
                            }
                        }
                    }
                    DebugLogger.Debug("No valid intersection points found");
                    return null;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.LogDetailed($"Task 2 sleeve location calculation failed - Element {mepElement.Id}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// EXACT COPY of working command intersection logic adapted for structural elements
        /// This mirrors the proven solid intersection logic from PipeSleeveCommand
        /// </summary>
        private bool ValidateStructuralIntersectionExact(Element mepElement, Element structuralElement, Line line)
        {
            try
            {
                // Check element type to determine validation approach
                bool isStructuralFraming = structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming;
                bool isFloor = structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors;

                // EXACT COPY of PipeSleeveCommand logic adapted for structural elements
                // Step 1: Extract structural element solid (EXACTLY like PipeSleeveCommand extracts wall solid)
                Solid structuralSolid = null;
                try 
                {
                    Options geomOptions = new Options { ComputeReferences = true, IncludeNonVisibleObjects = false };
                    
                    // Check if this is a linked element by looking for it in linked documents
                    Transform linkTransform = null;
                    bool isLinkedElement = false;
                    
                    // Search for the element in linked documents to get the transform
                    var linkedInstances = new FilteredElementCollector(mepElement.Document)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>();
                        
                    foreach (var linkInstance in linkedInstances)
                    {
                        var linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            try
                            {
                                var testElement = linkedDoc.GetElement(structuralElement.Id);
                                if (testElement != null)
                                {
                                    isLinkedElement = true;
                                    linkTransform = linkInstance.GetTotalTransform();
                                    break;
                                }
                            }
                            catch
                            {
                                // Element not in this linked document, continue searching
                            }
                        }
                    }
                    
                    GeometryElement geomElem = structuralElement.get_Geometry(geomOptions);
                    foreach (GeometryObject obj in geomElem) 
                    {
                        if (obj is Solid solid && solid.Volume > 0) 
                        {
                            // Apply transformation if this is a linked element
                            if (isLinkedElement && linkTransform != null)
                            {
                                structuralSolid = SolidUtils.CreateTransformed(solid, linkTransform);
                            }
                            else
                            {
                                structuralSolid = solid;
                            }
                            break;
                        }
                        else if (obj is GeometryInstance geomInstance)
                        {
                            // Handle geometry instances (common in family instances)
                            foreach (GeometryObject instObj in geomInstance.GetInstanceGeometry())
                            {
                                if (instObj is Solid instSolid && instSolid.Volume > 0)
                                {
                                    // Apply both instance transform and link transform if needed
                                    Transform combinedTransform = geomInstance.Transform;
                                    if (isLinkedElement && linkTransform != null)
                                    {
                                        combinedTransform = linkTransform.Multiply(geomInstance.Transform);
                                    }
                                    structuralSolid = SolidUtils.CreateTransformed(instSolid, combinedTransform);
                                    break;
                                }
                            }
                            if (structuralSolid != null) break;
                        }
                    }
                } 
                catch 
                { 
                    structuralSolid = null; 
                }

                if (structuralSolid == null)
                {
                    DebugLogger.Debug($"Could not extract solid from structural element {structuralElement.Id.IntegerValue}");
                    return false;
                }

                // Step 2: Face-by-face intersection testing (EXACTLY like PipeSleeveCommand)
                List<XYZ> intersectionPoints = new List<XYZ>();
                foreach (Face face in structuralSolid.Faces) 
                {
                    IntersectionResultArray ira = null;
                    SetComparisonResult res = face.Intersect(line, out ira);
                    if (res == SetComparisonResult.Overlap && ira != null) 
                    {
                        foreach (Autodesk.Revit.DB.IntersectionResult ir in ira) 
                        {
                            intersectionPoints.Add(ir.XYZPoint);
                        }
                    }
                }

                // Step 3: Endpoint inside/outside testing (EXACTLY like PipeSleeveCommand)
                bool startInside = false;
                bool endInside = false;
                try
                {
                    // Get structural element orientation (adapted from wall.Orientation logic)
                    XYZ structuralOrientation = XYZ.BasisZ; // Default to Z for floors
                    if (isStructuralFraming)
                    {
                        // For structural framing, try to get orientation similar to wall orientation
                        var locCurve = structuralElement.Location as LocationCurve;
                        if (locCurve?.Curve is Line framingLine)
                        {
                            // Use perpendicular to framing direction (similar to wall orientation)
                            var framingDir = framingLine.Direction.Normalize();
                            structuralOrientation = new XYZ(-framingDir.Y, framingDir.X, 0).Normalize();
                            if (structuralOrientation.GetLength() < 1e-6)
                            {
                                structuralOrientation = XYZ.BasisX; // Fallback
                            }
                        }
                        else
                        {
                            structuralOrientation = XYZ.BasisX; // Fallback for framing
                        }
                    }

                    startInside = IsPointInsideSolid(structuralSolid, line.GetEndPoint(0), structuralOrientation);
                    endInside = IsPointInsideSolid(structuralSolid, line.GetEndPoint(1), structuralOrientation);
                }
                catch (Exception ex)
                {
                    DebugLogger.Debug($"IsPointInsideSolid check failed: {ex.Message}");
                    startInside = false;
                    endInside = false;
                }

                // Step 4: Segmented sampling fallback (EXACTLY like PipeSleeveCommand)
                if (intersectionPoints.Count == 0 && !startInside && !endInside) 
                {
                    List<XYZ> altIntersections = new List<XYZ>();
                    int segments = 10;
                    for (int i = 0; i < segments; i++) 
                    {
                        double t1 = (double)i / segments;
                        double t2 = (double)(i + 1) / segments;
                        XYZ pt1 = line.Evaluate(t1, true);
                        XYZ pt2 = line.Evaluate(t2, true);
                        double segmentLength = pt1.DistanceTo(pt2);
                        if (segmentLength < 0.01) continue;
                        Line segment;
                        try { segment = Line.CreateBound(pt1, pt2); } catch { continue; }
                        foreach (Face face in structuralSolid.Faces) 
                        {
                            IntersectionResultArray ira2 = null;
                            SetComparisonResult res2 = face.Intersect(segment, out ira2);
                            if (res2 == SetComparisonResult.Overlap && ira2 != null) 
                            {
                                foreach (Autodesk.Revit.DB.IntersectionResult ir in ira2) 
                                {
                                    if (!altIntersections.Any(pt => pt.DistanceTo(ir.XYZPoint) < UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters))) 
                                    {
                                        altIntersections.Add(ir.XYZPoint);
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

                // Step 5: Final validation with element-specific logic
                bool hasValidIntersection = intersectionPoints.Count > 0 || startInside || endInside;

                if (hasValidIntersection)
                {
                    DebugLogger.Debug($"GEOMETRIC VALIDATION SUCCESS: MEP {mepElement.Id.IntegerValue} intersects {structuralElement.Category.Name} " +
                                    $"Id={structuralElement.Id.IntegerValue} (IntersectionPts: {intersectionPoints.Count}, StartInside: {startInside}, EndInside: {endInside})");
                    return true;
                }

                // PROXIMITY-BASED FALLBACK for both structural framing and floors
                double minDistance = CalculateMinimumDistance(line, structuralSolid);
                double proximityThresholdFraming = 5.0; // feet
                double proximityThresholdFloor = UnitUtils.ConvertToInternalUnits(150.0, UnitTypeId.Millimeters); // 150mm in feet
                bool isTargetElement = EXPECTED_MEP_ELEMENTS.Contains(mepElement.Id.IntegerValue);
                string targetPrefix = isTargetElement ? "🎯 TARGET " : "";

                if (isStructuralFraming)
                {
                    DebugLogger.Debug($"{targetPrefix}PROXIMITY CHECK: MEP {mepElement.Id.IntegerValue} to structural framing " +
                                    $"Id={structuralElement.Id.IntegerValue}, calculated distance: {minDistance:F3} feet, threshold: {proximityThresholdFraming} feet");
                    if (minDistance <= proximityThresholdFraming)
                    {
                        DebugLogger.Debug($"{targetPrefix}PROXIMITY FALLBACK SUCCESS: MEP {mepElement.Id.IntegerValue} near structural framing " +
                                        $"Id={structuralElement.Id.IntegerValue} at distance {minDistance:F3} feet (≤ {proximityThresholdFraming} threshold)");
                        if (isTargetElement)
                        {
                            DebugLogger.Info($"🎯 SUCCESS: Target element {mepElement.Id.IntegerValue} PASSED proximity validation!");
                        }
                        return true;
                    }
                    else
                    {
                        DebugLogger.Debug($"{targetPrefix}PROXIMITY FALLBACK FAILED: MEP {mepElement.Id.IntegerValue} too far from structural framing " +
                                        $"Id={structuralElement.Id.IntegerValue} at distance {minDistance:F3} feet (> {proximityThresholdFraming} threshold)");
                        if (isTargetElement)
                        {
                            DebugLogger.Info($"🎯 FAILED: Target element {mepElement.Id.IntegerValue} FAILED proximity validation - distance too large");
                        }
                    }
                }
                if (isFloor)
                {
                    DebugLogger.Debug($"{targetPrefix}FLOOR PROXIMITY CHECK: MEP {mepElement.Id.IntegerValue} to floor Id={structuralElement.Id.IntegerValue}, calculated distance: {minDistance:F4} feet, threshold: {proximityThresholdFloor:F4} feet");
                    if (minDistance <= proximityThresholdFloor)
                    {
                        DebugLogger.Debug($"{targetPrefix}FLOOR PROXIMITY FALLBACK SUCCESS: MEP {mepElement.Id.IntegerValue} near floor Id={structuralElement.Id.IntegerValue} at distance {minDistance:F4} feet (≤ {proximityThresholdFloor:F4} threshold)");
                        DebugLogger.Info($"FLOOR PROXIMITY FALLBACK TRIGGERED: Proceeding to placement for MEP {mepElement.Id.IntegerValue} and floor {structuralElement.Id.IntegerValue}");
                        if (isTargetElement)
                        {
                            DebugLogger.Info($"🎯 SUCCESS: Target element {mepElement.Id.IntegerValue} PASSED floor proximity validation!");
                        }
                        return true;
                    }
                    else
                    {
                        DebugLogger.Debug($"{targetPrefix}FLOOR PROXIMITY FALLBACK FAILED: MEP {mepElement.Id.IntegerValue} too far from floor Id={structuralElement.Id.IntegerValue} at distance {minDistance:F4} feet (> {proximityThresholdFloor:F4} threshold)");
                        if (isTargetElement)
                        {
                            DebugLogger.Info($"🎯 FAILED: Target element {mepElement.Id.IntegerValue} FAILED floor proximity validation - distance too large");
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                DebugLogger.Debug($"ValidateStructuralIntersectionExact failed for MEP {mepElement.Id.IntegerValue}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Calculate minimum distance between MEP line and structural solid for proximity-based validation
        /// </summary>
        private double CalculateMinimumDistance(Line line, Solid structuralSolid)
        {
            try
            {
                double minDistance = double.MaxValue;
                
                // Sample points along the line at regular intervals
                int samplePoints = 20;
                DebugLogger.Debug($"CalculateMinimumDistance: Sampling {samplePoints + 1} points along line from {line.GetEndPoint(0)} to {line.GetEndPoint(1)}");
                
                for (int i = 0; i <= samplePoints; i++)
                {
                    double parameter = (double)i / samplePoints;
                    XYZ pointOnLine = line.Evaluate(parameter, true);
                    
                    // Calculate distance from this point to the solid
                    double distanceToSolid = CalculateDistanceToSolid(pointOnLine, structuralSolid);
                    if (distanceToSolid < minDistance)
                    {
                        minDistance = distanceToSolid;
                    }
                }
                
                DebugLogger.Debug($"CalculateMinimumDistance: Final minimum distance = {minDistance:F3} feet");
                return minDistance;
            }
            catch (Exception ex)
            {
                DebugLogger.Debug($"CalculateMinimumDistance failed: {ex.Message}");
                return double.MaxValue;
            }
        }

        /// <summary>
        /// Calculate distance from a point to the nearest surface of a solid
        /// </summary>
        private double CalculateDistanceToSolid(XYZ point, Solid solid)
        {
            try
            {
                // Use ray-casting method to check if point is inside solid (already implemented in IsPointInsideSolid)
                // We'll use the solid's bounding box to pick a reasonable normal
                BoundingBoxXYZ bbox = solid.GetBoundingBox();
                XYZ normal = (bbox != null) ? (bbox.Max - bbox.Min).Normalize() : XYZ.BasisZ;
                if (IsPointInsideSolid(solid, point, normal))
                {
                    return 0.0; // Point is inside, distance is zero
                }

                double minDistance = double.MaxValue;

                // Method 1: Try face projection (most accurate when it works)
                foreach (Face face in solid.Faces)
                {
                    try
                    {
                        Autodesk.Revit.DB.IntersectionResult result = face.Project(point);
                        if (result != null && result.XYZPoint != null)
                        {
                            double distance = point.DistanceTo(result.XYZPoint);
                            if (distance < minDistance)
                            {
                                minDistance = distance;
                            }
                        }
                    }
                    catch
                    {
                        // Skip this face if projection fails
                        continue;
                    }
                }

                // Method 2: If face projection didn't work well, use bounding box approximation
                if (minDistance > 10.0) // If distance seems too large, try bounding box method
                {
                    if (bbox != null)
                    {
                        XYZ min = bbox.Min;
                        XYZ max = bbox.Max;

                        // Calculate distance to bounding box
                        double dx = Math.Max(0, Math.Max(min.X - point.X, point.X - max.X));
                        double dy = Math.Max(0, Math.Max(min.Y - point.Y, point.Y - max.Y));
                        double dz = Math.Max(0, Math.Max(min.Z - point.Z, point.Z - max.Z));

                        double bboxDistance = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                        if (bboxDistance < minDistance)
                        {
                            minDistance = bboxDistance;
                        }
                    }
                }

                // Method 3: Last resort - distance to solid centroid
                if (minDistance == double.MaxValue || minDistance > 100.0)
                {
                    XYZ centroid = solid.ComputeCentroid();
                    if (centroid != null)
                    {
                        minDistance = point.DistanceTo(centroid);
                    }
                }

                return minDistance;
            }
            catch (Exception ex)
            {
                DebugLogger.Debug($"CalculateDistanceToSolid failed: {ex.Message}");
                return double.MaxValue;
            }
        }

        /// <summary>
        /// EXACT COPY of IsPointInsideSolid from PipeSleeveCommand
        /// </summary>
        private static bool IsPointInsideSolid(Solid solid, XYZ point, XYZ structuralNormal)
        {
            // Check if a point is inside a solid using a ray-casting method
            // Cast a ray in the direction of the structural element normal
            XYZ rayDirection = structuralNormal.Normalize();
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
    }

    public static class MEPElementCollector
    {
        public static List<Element> CollectMEPElements(Document doc)
        {
            try
            {
                DebugLogger.Debug("Starting MEP element collection from active document and linked files.");
                var mepElements = new List<Element>();

                // Collect from active document
                var pipesActive = new FilteredElementCollector(doc)
                    .OfClass(typeof(Pipe))
                    .WhereElementIsNotElementType()
                    .ToList();
                
                var ductsActive = new FilteredElementCollector(doc)
                    .OfClass(typeof(Duct))
                    .WhereElementIsNotElementType()
                    .ToList();
                    
                var cableTraysActive = new FilteredElementCollector(doc)
                    .OfClass(typeof(CableTray))
                    .WhereElementIsNotElementType()
                    .ToList();

                mepElements.AddRange(pipesActive);
                mepElements.AddRange(ductsActive);
                mepElements.AddRange(cableTraysActive);
                
                DebugLogger.Debug($"Collected from active document: {pipesActive.Count} pipes, {ductsActive.Count} ducts, {cableTraysActive.Count} cable trays.");

                // Collect from linked documents (VISIBLE ONLY)
                var linkInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .WhereElementIsNotElementType()
                    .Cast<RevitLinkInstance>()
                    .ToList();

                foreach (var linkInstance in linkInstances)
                {
                    // FILTER: Only process VISIBLE linked models
                    if (!linkInstance.IsHidden(doc.ActiveView))
                    {
                        var linkedDoc = linkInstance.GetLinkDocument();
                        if (linkedDoc != null)
                        {
                            try
                            {
                                var pipesLinked = new FilteredElementCollector(linkedDoc)
                                    .OfClass(typeof(Pipe))
                                    .WhereElementIsNotElementType()
                                    .ToList();
                                    
                                var ductsLinked = new FilteredElementCollector(linkedDoc)
                                    .OfClass(typeof(Duct))
                                    .WhereElementIsNotElementType()
                                    .ToList();
                                    
                                var cableTraysLinked = new FilteredElementCollector(linkedDoc)
                                    .OfClass(typeof(CableTray))
                                    .WhereElementIsNotElementType()
                                    .ToList();

                                mepElements.AddRange(pipesLinked);
                                mepElements.AddRange(ductsLinked);
                                mepElements.AddRange(cableTraysLinked);
                                
                                DebugLogger.Debug($"Collected from VISIBLE linked document {linkedDoc.Title}: {pipesLinked.Count} pipes, {ductsLinked.Count} ducts, {cableTraysLinked.Count} cable trays.");
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Error($"Error collecting MEP elements from linked document {linkedDoc?.Title}: {ex.Message}");
                            }
                        }
                    }
                    else
                    {
                        DebugLogger.Debug($"Skipped HIDDEN linked document: {linkInstance.Name}");
                    }
                }

                DebugLogger.Debug($"Total MEP elements collected: {mepElements.Count}.");
                return mepElements;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error in CollectMEPElements: {ex.Message}");
                return new List<Element>();
            }
        }

        /// <summary>
        /// Enhanced geometric validation using solid intersection - matches working command pattern
        /// </summary>
        public static bool ValidateGeometricIntersection(Element mepElement, Element structuralElement)
        {
            try
            {
                // Get MEP element curve
                var locCurve = mepElement.Location as LocationCurve;
                if (locCurve?.Curve == null) return false;
                
                var line = locCurve.Curve as Line;
                if (line == null) return false;

                // Get structural element solid geometry (matches working command pattern)
                Solid structuralSolid = null;
                try 
                {
                    Options geomOptions = new Options { ComputeReferences = true, IncludeNonVisibleObjects = false };
                    GeometryElement geomElem = structuralElement.get_Geometry(geomOptions);
                    foreach (GeometryObject obj in geomElem) 
                    {
                        if (obj is Solid solid && solid.Volume > 0) 
                        {
                            structuralSolid = solid;
                            break;
                        }
                    }
                } 
                catch { return false; }
                
                if (structuralSolid == null) return false;

                // Robust intersection logic (exact copy from working PipeSleeveCommand)
                List<XYZ> intersectionPoints = new List<XYZ>();
                foreach (Face face in structuralSolid.Faces) 
                {
                    IntersectionResultArray ira = null;
                    SetComparisonResult res = face.Intersect(line, out ira);
                    if (res == SetComparisonResult.Overlap && ira != null) 
                    {
                        foreach (Autodesk.Revit.DB.IntersectionResult ir in ira) 
                        {
                            intersectionPoints.Add(ir.XYZPoint);
                        }
                    }
                }

                // Check if endpoints are inside solid (working command pattern)
                bool startInside = false;
                bool endInside = false;
                try
                {
                    XYZ structuralNormal = XYZ.BasisZ; // Default for floors/framing
                    if (structuralElement is Wall wall)
                        structuralNormal = wall.Orientation;
                    
                    startInside = IsPointInsideSolid(structuralSolid, line.GetEndPoint(0), structuralNormal);
                    endInside = IsPointInsideSolid(structuralSolid, line.GetEndPoint(1), structuralNormal);
                }
                catch 
                {
                    startInside = false;
                    endInside = false;
                }

                // Segment sampling fallback if no direct intersections (working command pattern)
                if (intersectionPoints.Count == 0 && !startInside && !endInside) 
                {
                    List<XYZ> altIntersections = new List<XYZ>();
                    int segments = 10;
                    for (int i = 0; i < segments; i++) 
                    {
                        double t1 = (double)i / segments;
                        double t2 = (double)(i + 1) / segments;
                        XYZ pt1 = line.Evaluate(t1, true);
                        XYZ pt2 = line.Evaluate(t2, true);
                        double segmentLength = pt1.DistanceTo(pt2);
                        if (segmentLength < 0.01) continue;
                        
                        Line segment;
                        try { segment = Line.CreateBound(pt1, pt2); } catch { continue; }
                        
                        foreach (Face face in structuralSolid.Faces) 
                        {
                            IntersectionResultArray ira = null;
                            SetComparisonResult res = face.Intersect(segment, out ira);
                            if (res == SetComparisonResult.Overlap && ira != null) 
                            {
                                foreach (Autodesk.Revit.DB.IntersectionResult ir in ira) 
                                {
                                    if (!altIntersections.Any(pt => pt.DistanceTo(ir.XYZPoint) < UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Millimeters))) 
                                    {
                                        altIntersections.Add(ir.XYZPoint);
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

                // Final validation: must have actual intersections or endpoints inside
                bool hasValidIntersection = intersectionPoints.Count > 0 || startInside || endInside;
                
                if (hasValidIntersection && intersectionPoints.Count >= 2)
                {
                    // Check minimum penetration length (working command pattern)
                    var sortedPoints = intersectionPoints.OrderBy(pt => (pt - line.GetEndPoint(0)).GetLength()).ToList();
                    double penetrationLength = sortedPoints.First().DistanceTo(sortedPoints.Last());
                    hasValidIntersection = penetrationLength >= UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters);
                }

                return hasValidIntersection;
            }
            catch (Exception ex)
            {
                DebugLogger.LogDetailed($"Geometric validation failed - Element {mepElement.Id}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Point-in-solid validation method from working PipeSleeveCommand
        /// </summary>
        private static bool IsPointInsideSolid(Solid solid, XYZ point, XYZ wallNormal)
        {
            try
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
            catch
            {
                return false;
            }
        }
    }

    public static class SleevePlacer
    {
        public static void PlaceSleeve(Document doc, IntersectionResult intersection)
        {
            try
            {
                DebugLogger.Debug($"Placing sleeve at intersection point: {intersection.IntersectionPoint}.");

                // Determine the family to use based on the structural element type
                FamilySymbol familySymbol = null;
                if (intersection.StructuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                {
                    // Use slab opening families
                    if (intersection.MEPElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves)
                    {
                        familySymbol = GetFamilySymbol(doc, "PipeOpeningOnSlab");
                    }
                    else if (intersection.MEPElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves)
                    {
                        familySymbol = GetFamilySymbol(doc, "DuctOpeningOnSlab");
                    }
                    else if (intersection.MEPElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
                    {
                        familySymbol = GetFamilySymbol(doc, "PipeOpeningOnSlabRect");
                    }
                }
                else if (intersection.StructuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // Use wall opening families
                    if (intersection.MEPElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves)
                    {
                        familySymbol = GetFamilySymbol(doc, "PipeOpeningOnWall");
                    }
                    else
                    {
                        familySymbol = GetFamilySymbol(doc, "OnWall");
                    }
                }

                if (familySymbol == null)
                {
                    DebugLogger.Warning($"No suitable family found for the intersection between MEP element {intersection.MEPElement.Id} and structural element {intersection.StructuralElement.Id}.");
                    return;
                }

                // Ensure the family symbol is activated
                if (!familySymbol.IsActive)
                {
                    familySymbol.Activate();
                    doc.Regenerate();
                }

                // Place the family instance at the intersection point
                using (Transaction trans = new Transaction(doc, "Place Sleeve"))
                {
                    trans.Start();
                    doc.Create.NewFamilyInstance(intersection.IntersectionPoint, familySymbol, intersection.StructuralElement, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    trans.Commit();
                }

                DebugLogger.Info("Sleeve placed successfully.");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error placing sleeve: {ex.Message}");
            }
        }

        private static FamilySymbol GetFamilySymbol(Document doc, string familyName)
        {
            // Retrieve the family symbol by name
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(fs => fs.Family.Name.Contains(familyName));
        }

        /// <summary>
        /// TASK 2: Place appropriate sleeve family based on structural element type
        /// 
        /// IMPORTANT: DO NOT CHANGE THE FAMILY PLACEMENT LINES BELOW UNLESS YOU FULLY UNDERSTAND THE WORKFLOW.
        /// - We are placing Generic Model, work plane-based sleeve families in the ACTIVE DOCUMENT ONLY.
        /// - These families do NOT require a host or level and are NOT wall-hosted or face-based.
        /// - The placement must use NewFamilyInstance(XYZ, FamilySymbol, StructuralType.NonStructural) ONLY.
        /// - The geometry of linked elements is used for intersection and placement point calculation, but sleeves are always created in the active model.
        /// - This logic is proven and matches the working DuctSleevePlacer and PipeSleevePlacer patterns.
        /// - If you need to support a different family type, create a new placer class. DO NOT edit this one.
        /// </summary>
        public static void PlaceSleeveAtMidpoint(Element mepElement, Element structuralElement, XYZ placementPoint)
        {
            try
            {
                DebugLogger.Info($"[DEBUG] Entered PlaceSleeveAtMidpoint: MEP {mepElement?.Id.IntegerValue}, Structural {structuralElement?.Id.IntegerValue}, Point: {placementPoint}");
                Document doc = mepElement.Document;
                // Determine appropriate family based on MEP element type and structural element type
                FamilySymbol familySymbol = null;
                string familyName = "";
                bool isFloor = structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors;
                bool isFraming = structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming;
                // --- Only place sleeves on structural floors (skip architectural) ---
                if (isFloor)
                {
                    // Check instance parameter only
                    Parameter structuralParam = structuralElement.LookupParameter("Structural");
                    bool? isStructural = null;
                    int? rawStructuralValue = null;
                    if (structuralParam != null && structuralParam.StorageType == StorageType.Integer)
                    {
                        rawStructuralValue = structuralParam.AsInteger();
                        isStructural = rawStructuralValue == 1;
                    }
                    DebugLogger.Info($"[DEBUG] Floor {structuralElement.Id.IntegerValue} | Name: '{structuralElement.Name}' | Category: '{structuralElement.Category.Name}' | 'Structural' param: {(rawStructuralValue.HasValue ? rawStructuralValue.ToString() : "<null>")} | isStructural: {(isStructural.HasValue ? isStructural.ToString() : "<null>")}");
                    // Only skip if parameter exists and is explicitly set to false (0)
                    if (isStructural.HasValue && !isStructural.Value)
                    {
                        DebugLogger.Info($"[SKIP] Floor {structuralElement.Id.IntegerValue} is not structural (instance param). Skipping sleeve placement.");
                        return;
                    }
                    // For floors: Use "OnSlab" families
                    if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves)
                        familyName = "PipeOpeningOnSlab";
                    else if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves)
                        familyName = "DuctOpeningOnSlab";
                    else if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
                        familyName = "CableTrayOpeningOnSlab";
                }
                else if (isFraming)
                {
                    // For structural framing: Use "OnWall" families
                    if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves)
                        familyName = "PipeOpeningOnWall";
                    else if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves)
                        familyName = "DuctOpeningOnWall";
                    else if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
                        familyName = "CableTrayOpeningOnWall";
                }
                DebugLogger.Info($"[DEBUG] PlaceSleeveAtMidpoint: familyName resolved to '{familyName}'");
                if (string.IsNullOrEmpty(familyName))
                {
                    DebugLogger.Warning($"TASK 2: No family mapping found for MEP {mepElement.Category.Name} + {structuralElement.Category.Name}");
                    return;
                }
                // Find the family symbol
                familySymbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(fs => fs.Family.Name.Contains(familyName, StringComparison.OrdinalIgnoreCase));
                DebugLogger.Info($"[DEBUG] PlaceSleeveAtMidpoint: familySymbol {(familySymbol == null ? "NOT FOUND" : "FOUND: " + familySymbol.Family.Name)}");
                if (familySymbol == null)
                {
                    DebugLogger.Warning($"TASK 2: Family '{familyName}' not found in document for MEP {mepElement.Id.IntegerValue}");
                    return;
                }
                // Ensure family symbol is active
                using (var txActivate = new Transaction(doc, "Activate Family Symbol"))
                {
                    txActivate.Start();
                    if (!familySymbol.IsActive)
                        familySymbol.Activate();
                    txActivate.Commit();
                }
                // TASK 4: Duplication suppression - Check for existing sleeves at this location
                double searchRadius = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters); // 25mm radius
                var existingSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Contains("Opening", StringComparison.OrdinalIgnoreCase))
                    .ToList();
                foreach (var existingSleeve in existingSleeves)
                {
                    LocationPoint loc = existingSleeve.Location as LocationPoint;
                    if (loc != null)
                    {
                        double dist = loc.Point.DistanceTo(placementPoint);
                        if (dist < searchRadius)
                        {
                            DebugLogger.Info($"✅ TASK 4: Duplicate sleeve detected at {placementPoint} (existing sleeve ID: {existingSleeve.Id.IntegerValue}, distance: {UnitUtils.ConvertFromInternalUnits(dist, UnitTypeId.Millimeters):F1}mm) - skipping placement for MEP {mepElement.Id.IntegerValue}");
                            return;
                        }
                    }
                }
                DebugLogger.Debug($"🔍 TASK 4: No duplicate sleeves found within {UnitUtils.ConvertFromInternalUnits(searchRadius, UnitTypeId.Millimeters):F1}mm radius at {placementPoint}");
                // Start transaction for sleeve placement
                using (var tx = new Transaction(doc, "Place Structural Sleeve - Task 2"))
                {
                    tx.Start();
                    // Place sleeve family instance (work plane-based, no host/level)
                    FamilyInstance sleeveInstance = doc.Create.NewFamilyInstance(
                        placementPoint,
                        familySymbol,
                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    // Set sleeve dimensions based on MEP element type and existing working command patterns
                    SetSleeveDimensions(sleeveInstance, mepElement, structuralElement);
                    // TASK 3: Set reference level from MEP element to sleeve (following working command patterns)
                    SetSleeveReferenceLevel(sleeveInstance, mepElement, doc);
                    // Log successful placement
                    DebugLogger.Info($"✅ TASK 2 SLEEVE PLACED: {familyName} Id={sleeveInstance.Id.IntegerValue} " +
                                   $"for MEP {mepElement.Category.Name} {mepElement.Id.IntegerValue} at {placementPoint} " +
                                   $"intersecting {structuralElement.Category.Name} Id={structuralElement.Id.IntegerValue}");
                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"TASK 2 SLEEVE PLACEMENT ERROR: MEP {mepElement.Id.IntegerValue} " +
                               $"with {structuralElement.Category.Name} Id={structuralElement.Id.IntegerValue}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Set sleeve dimensions following the exact logic from working commands (PipeSleevePlacer, DuctSleevePlacer, CableTraySleevePlacer)
        /// </summary>
        private static void SetSleeveDimensions(FamilyInstance sleeveInstance, Element mepElement, Element structuralElement)
        {
            try
            {
                Document doc = sleeveInstance.Document;
                
                // Get structural element thickness for depth parameter
                double structuralThickness = GetStructuralElementThickness(structuralElement);
                
                if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_PipeCurves)
                {
                    // PIPE LOGIC: Following PipeSleevePlacer pattern exactly
                    var pipe = mepElement as Pipe;
                    if (pipe != null)
                    {
                        // Get pipe diameter
                        double pipeDiameter = 0.0;
                        var diameterParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                        if (diameterParam != null && diameterParam.HasValue)
                        {
                            pipeDiameter = diameterParam.AsDouble();
                        }
                        
                        // Check for insulation (exact logic from PipeSleeveCommand)
                        double clearancePerSideMM = 50.0; // default: 50mm per side for non-insulated
                        var insulationThicknessParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS);
                        if (insulationThicknessParam != null && insulationThicknessParam.HasValue && insulationThicknessParam.AsDouble() > 0)
                        {
                            clearancePerSideMM = 25.0; // 25mm per side for insulated
                        }
                        
                        // Calculate final diameter with clearance (following working command pattern)
                        double totalClearance = UnitUtils.ConvertToInternalUnits(clearancePerSideMM * 2.0, UnitTypeId.Millimeters);
                        double finalDiameter = pipeDiameter + totalClearance;
                        
                        // Set parameters
                        SetParameterSafely(sleeveInstance, "Diameter", finalDiameter, mepElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Depth", structuralThickness, mepElement.Id.IntegerValue);
                        
                        DebugLogger.Info($"PIPE SLEEVE DIMENSIONS: Pipe {mepElement.Id.IntegerValue} - " +
                                       $"Original Diameter: {UnitUtils.ConvertFromInternalUnits(pipeDiameter, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Clearance: {clearancePerSideMM}mm/side, " +
                                       $"Final Diameter: {UnitUtils.ConvertFromInternalUnits(finalDiameter, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Depth: {UnitUtils.ConvertFromInternalUnits(structuralThickness, UnitTypeId.Millimeters):F0}mm");
                    }
                }
                else if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctCurves)
                {
                    // DUCT LOGIC: Following DuctSleevePlacer pattern exactly
                    var duct = mepElement as Duct;
                    if (duct != null)
                    {
                        // Get duct dimensions
                        double ductWidth = 0.0, ductHeight = 0.0;
                        var widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                        var heightParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                        
                        if (widthParam != null && widthParam.HasValue)
                            ductWidth = widthParam.AsDouble();
                        if (heightParam != null && heightParam.HasValue)
                            ductHeight = heightParam.AsDouble();
                        
                        // Add 50mm clearance on all sides (following working command pattern)
                        double clearanceMM = 50.0;
                        double clearance = UnitUtils.ConvertToInternalUnits(clearanceMM, UnitTypeId.Millimeters);
                        double finalWidth = ductWidth + (clearance * 2);
                        double finalHeight = ductHeight + (clearance * 2);
                        
                        // Set parameters
                        SetParameterSafely(sleeveInstance, "Width", finalWidth, mepElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Height", finalHeight, mepElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Depth", structuralThickness, mepElement.Id.IntegerValue);
                        
                        DebugLogger.Info($"DUCT SLEEVE DIMENSIONS: Duct {mepElement.Id.IntegerValue} - " +
                                       $"Original: {UnitUtils.ConvertFromInternalUnits(ductWidth, UnitTypeId.Millimeters):F0}mm x {UnitUtils.ConvertFromInternalUnits(ductHeight, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Clearance: {clearanceMM}mm/side, " +
                                       $"Final: {UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F0}mm x {UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Depth: {UnitUtils.ConvertFromInternalUnits(structuralThickness, UnitTypeId.Millimeters):F0}mm");
                    }
                }
                else if (mepElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_CableTray)
                {
                    // CABLE TRAY LOGIC: Following CableTraySleevePlacer pattern exactly
                    var cableTray = mepElement as CableTray;
                    if (cableTray != null)
                    {
                        // Get cable tray dimensions
                        double trayWidth = 0.0, trayHeight = 0.0;
                        var widthParam = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
                        var heightParam = cableTray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
                        
                        if (widthParam != null && widthParam.HasValue)
                            trayWidth = widthParam.AsDouble();
                        if (heightParam != null && heightParam.HasValue)
                            trayHeight = heightParam.AsDouble();
                        
                        // Add 50mm clearance on all sides (following working command pattern)
                        double clearanceMM = 50.0;
                        double clearance = UnitUtils.ConvertToInternalUnits(clearanceMM, UnitTypeId.Millimeters);
                        double finalWidth = trayWidth + (clearance * 2);
                        double finalHeight = trayHeight + (clearance * 2);
                        
                        // Set parameters
                        SetParameterSafely(sleeveInstance, "Width", finalWidth, mepElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Height", finalHeight, mepElement.Id.IntegerValue);
                        SetParameterSafely(sleeveInstance, "Depth", structuralThickness, mepElement.Id.IntegerValue);
                        
                        DebugLogger.Info($"CABLE TRAY SLEEVE DIMENSIONS: Cable Tray {mepElement.Id.IntegerValue} - " +
                                       $"Original: {UnitUtils.ConvertFromInternalUnits(trayWidth, UnitTypeId.Millimeters):F0}mm x {UnitUtils.ConvertFromInternalUnits(trayHeight, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Clearance: {clearanceMM}mm/side, " +
                                       $"Final: {UnitUtils.ConvertFromInternalUnits(finalWidth, UnitTypeId.Millimeters):F0}mm x {UnitUtils.ConvertFromInternalUnits(finalHeight, UnitTypeId.Millimeters):F0}mm, " +
                                       $"Depth: {UnitUtils.ConvertFromInternalUnits(structuralThickness, UnitTypeId.Millimeters):F0}mm");
                    }
                }
                
                // Regenerate document after setting parameters
                doc.Regenerate();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error setting sleeve dimensions for MEP {mepElement.Id.IntegerValue}: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Get structural element thickness for depth parameter
        /// </summary>
        private static double GetStructuralElementThickness(Element structuralElement)
        {
            try
            {
                if (structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors)
                {
                    // For floors, get thickness parameter
                    var floor = structuralElement as Floor;
                    if (floor != null)
                    {
                        var thicknessParam = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                        if (thicknessParam != null && thicknessParam.HasValue)
                        {
                            return thicknessParam.AsDouble();
                        }
                    }
                    // Fallback for floors: use default thickness
                    return UnitUtils.ConvertToInternalUnits(300.0, UnitTypeId.Millimeters); // 300mm default
                }
                else if (structuralElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // For structural framing, get the 'b' TYPE parameter from the FamilySymbol
                    var familyInstance = structuralElement as FamilyInstance;
                    if (familyInstance != null)
                    {
                        var symbol = structuralElement.Document.GetElement(familyInstance.GetTypeId()) as FamilySymbol;
                        var bParam = symbol?.LookupParameter("b");
                        if (bParam != null && bParam.StorageType == StorageType.Double)
                        {
                            return bParam.AsDouble();
                        }
                        // If 'b' is not set, inform user and abort
                        TaskDialog.Show("Missing Type Parameter", $"The structural framing type parameter 'b' is not set for type '{symbol?.Name}'. Please set this parameter and try again.");
                        throw new InvalidOperationException($"Missing required type parameter 'b' on structural framing type '{symbol?.Name}'");
                    }
                    // If not a FamilyInstance, inform user and abort
                    TaskDialog.Show("Invalid Structural Element", "The selected structural element is not a FamilyInstance. Cannot determine thickness.");
                    throw new InvalidOperationException("Structural element is not a FamilyInstance");
                }
                // Default fallback
                return UnitUtils.ConvertToInternalUnits(300.0, UnitTypeId.Millimeters); // 300mm default
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Error getting structural element thickness: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Safely set parameter value (following working command pattern)
        /// </summary>
        private static void SetParameterSafely(FamilyInstance instance, string paramName, double value, int elementId)
        {
            try
            {
                var param = instance.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(value);
                    DebugLogger.Debug($"Set {paramName} = {UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters):F1}mm for element {elementId}");
                }
                else
                {
                    DebugLogger.Warning($"Parameter '{paramName}' not found or read-only for element {elementId}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to set parameter '{paramName}' for element {elementId}: {ex.Message}");
            }
        }

        /// <summary>
        /// TASK 3: Set sleeve's "Schedule Level" parameter to match MEP element's reference level
        /// Following exact pattern from PipeSleevePlacer, DuctSleevePlacer, and CableTraySleevePlacer
        /// </summary>
        private static void SetSleeveReferenceLevel(FamilyInstance sleeveInstance, Element mepElement, Document doc)
        {
            try
            {
                // Extract MEP element's reference level using HostLevelHelper
                Level mepReferenceLevel = HostLevelHelper.GetHostReferenceLevel(doc, mepElement);
                
                if (mepReferenceLevel != null)
                {
                    // Set sleeve's Schedule Level parameter to MEP's reference level
                    Parameter schedLevelParam = sleeveInstance.LookupParameter("Schedule Level");
                    if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                    {
                        schedLevelParam.Set(mepReferenceLevel.Id);
                        DebugLogger.Info($"✅ TASK 3: Set Schedule Level to '{mepReferenceLevel.Name}' for sleeve {sleeveInstance.Id.IntegerValue} (MEP: {mepElement.Id.IntegerValue})");
                    }
                    else
                    {
                        DebugLogger.Warning($"⚠️ TASK 3: Schedule Level parameter not found/readonly for sleeve {sleeveInstance.Id.IntegerValue}");
                    }
                }
                else
                {
                    DebugLogger.Warning($"⚠️ TASK 3: No reference level found for MEP element {mepElement.Id.IntegerValue}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"❌ TASK 3: Failed to set reference level for sleeve {sleeveInstance.Id.IntegerValue}: {ex.Message}");
            }
        }

        /// <summary>
        /// Compatibility helper method to get intersection point from IntersectionResult
        /// Handles different property names across Revit versions (2020-2024)
        /// </summary>
        private static XYZ GetIntersectionPoint(Autodesk.Revit.DB.IntersectionResult ir)
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

    /// <summary>
    /// Custom IntersectionResult class for MEP-Structural intersections
    /// This is distinct from Autodesk.Revit.DB.IntersectionResult
    /// </summary>
    public class IntersectionResult
    {
        public Element MEPElement { get; set; }
        public Element StructuralElement { get; set; }
        public XYZ IntersectionPoint { get; set; }
    }
}
