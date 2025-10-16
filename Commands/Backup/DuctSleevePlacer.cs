using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// FINAL PRODUCTION-SAFE DUCT SLEEVE PLACER
    /// 
    /// DO NOT CHANGE THE CORE LOGIC BELOW UNLESS YOU FULLY UNDERSTAND:
    /// - The sleeve family is a Generic Model, Workplane-Based (NOT wall-hosted, NOT face-based).
    /// - The family is placed in the active model, but the wall may be from a linked file (linked wall).
    /// - Host-based placement (passing a wall as host) will NOT work for this family and will break linked wall workflows.
    /// - This logic is proven to work for both through-ducts and stub/capped ducts (ducts that enter but do not exit a wall).
    /// - All placement is done using NewFamilyInstance(XYZ, FamilySymbol, Level, StructuralType.NonStructural) ONLY.
    /// - Parameter setting and logging are robust and production-safe. DO NOT REMOVE or bypass these checks.
    /// - If you need to support a different family type, create a new placer class. DO NOT edit this one.
    /// 
    /// If you are unsure, consult the Revit API docs and the PipeSleevePlacer logic before making changes.
    /// </summary>
    public class DuctSleevePlacer
    {
        private readonly Document _doc;

        public DuctSleevePlacer(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            DebugLogger.Log($"[DuctSleevePlacer] Constructor called");
        }

        /// <summary>
        /// FIXED: Place duct sleeve with explicit level reference (like your reference code)
        /// </summary>
        public void PlaceDuctSleeveOptimized(Duct duct, XYZ placementPoint, double finalWidth, double finalHeight, 
            XYZ mepOrientation, FamilySymbol sleeveSymbol, string structuralElementType, double structuralElementThickness, string levelName, double levelElevation)
        {
            int ductElementId = (int)(duct?.Id?.IntegerValue ?? 0);
            try
            {
                DebugLogger.Log($"[DuctSleevePlacer] FIXED: Placing sleeve for duct {ductElementId}");
                DebugLogger.Log($"[DuctSleevePlacer] FIXED: Placement point: {placementPoint}");
                DebugLogger.Log($"[DuctSleevePlacer] FIXED: Dimensions: {finalWidth} x {finalHeight}");
                DebugLogger.Log($"[DuctSleevePlacer] FIXED: Orientation: {mepOrientation}");
                
                if (duct == null || sleeveSymbol == null)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] ERROR: Null parameters provided for duct {ductElementId}");
                    return;
                }

                // CRITICAL: Find nearest level for explicit placement (like your reference)
                Level nearestLevel = FindNearestLevel(placementPoint);
                if (nearestLevel == null)
                {
                    DebugLogger.Error($"[DuctSleevePlacer] ERROR: Could not find nearest level for placement point {placementPoint}");
                    return;
                }
                DebugLogger.Log($"[DuctSleevePlacer] FIXED: Using nearest level '{nearestLevel.Name}' (elevation: {nearestLevel.Elevation:F2}) for placement at Z={placementPoint.Z:F2}");

                // Activate the family symbol if not already active
                if (!sleeveSymbol.IsActive)
                {
                    sleeveSymbol.Activate();
                }

                // CRITICAL: Place with explicit level reference (like your reference code)
                FamilyInstance sleeveInstance = null;
                
                // Place without host element (freestanding) - using pre-calculated thickness for depth
                sleeveInstance = _doc.Create.NewFamilyInstance(
                    placementPoint,
                    sleeveSymbol,
                    nearestLevel,  // ← EXPLICIT LEVEL
                    Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    
                DebugLogger.Log($"[DuctSleevePlacer] FIXED: Placed freestanding");

                if (sleeveInstance == null)
                {
                    DebugLogger.Error($"[DuctSleevePlacer] ERROR: Failed to create sleeve instance for duct {ductElementId}");
                    return;
                }

                // Set sleeve parameters using pre-calculated dimensions
                // Check if this is a round (circular) sleeve family - uses Diameter parameter
                var diameterParam = sleeveInstance.LookupParameter("Diameter");
                if (diameterParam != null && !diameterParam.IsReadOnly)
                {
                    // Round sleeve - set diameter
                    diameterParam.Set(finalWidth); // finalWidth = diameter for round ducts
                    DebugLogger.Log($"[DuctSleevePlacer] FIXED: Set diameter to {finalWidth:F2} (round sleeve)");
                }
                else
                {
                    // Rectangular sleeve - set width and height
                    var widthParam = sleeveInstance.LookupParameter("Width");
                    if (widthParam != null && !widthParam.IsReadOnly)
                    {
                        widthParam.Set(finalWidth);
                        DebugLogger.Log($"[DuctSleevePlacer] FIXED: Set width to {finalWidth:F2}");
                    }

                    var heightParam = sleeveInstance.LookupParameter("Height");
                    if (heightParam != null && !heightParam.IsReadOnly)
                    {
                        heightParam.Set(finalHeight);
                        DebugLogger.Log($"[DuctSleevePlacer] FIXED: Set height to {finalHeight:F2}");
                    }
                }

                // Set depth parameter using pre-calculated thickness
                DebugLogger.Log($"[DuctSleevePlacer] FIXED: About to set depth for structural element type: {structuralElementType}, thickness: {structuralElementThickness:F3}");
                
                if (structuralElementThickness > 0)
                {
                    var depthParam = sleeveInstance.LookupParameter("Depth") ?? sleeveInstance.LookupParameter("d");
                    if (depthParam != null && !depthParam.IsReadOnly)
                    {
                        depthParam.Set(structuralElementThickness);
                        var depthMm = UnitUtils.ConvertFromInternalUnits(structuralElementThickness, UnitTypeId.Millimeters);
                        DebugLogger.Log($"[DuctSleevePlacer] FIXED: Set depth to {depthMm:F1}mm ({structuralElementThickness:F3} internal) for element type: {structuralElementType}");
                    }
                    else
                    {
                        DebugLogger.Warning($"[DuctSleevePlacer] FIXED: Depth parameter not found or read-only for sleeve instance {sleeveInstance.Id}");
                    }
                }
                else
                {
                    DebugLogger.Warning($"[DuctSleevePlacer] FIXED: Structural element thickness is 0 for type {structuralElementType}");
                }

                // CRITICAL: Apply proper orientation using MEP element direction
                if (mepOrientation != null)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] FIXED: Applying orientation: {mepOrientation}");
                    
                    // Get the sleeve instance's location point for rotation
                    var locationPoint = sleeveInstance.Location as LocationPoint;
                    if (locationPoint != null)
                    {
                        // Calculate rotation angle from MEP orientation
                        double rotationAngle = Math.Atan2(mepOrientation.Y, mepOrientation.X);
                        double rotationAngleDegrees = rotationAngle * 180 / Math.PI;
                        
                        DebugLogger.Log($"[DuctSleevePlacer] FIXED: Rotation angle: {rotationAngleDegrees:F1} degrees");
                        
                        // CRITICAL FIX: Rotate around the placement point, not the sleeve center
                        // This prevents the sleeve from moving away from the wall center
                        var rotationAxis = XYZ.BasisZ; // Rotate around Z-axis
                        var rotationLine = Line.CreateUnbound(placementPoint, rotationAxis);
                        locationPoint.Rotate(rotationLine, rotationAngle);
                        
                        DebugLogger.Log($"[DuctSleevePlacer] FIXED: Sleeve rotated around placement point to match MEP orientation");
                    }
                    else
                    {
                        DebugLogger.Warning($"[DuctSleevePlacer] FIXED: Cannot rotate sleeve - LocationPoint not available");
                    }
                }

                DebugLogger.Log($"[DuctSleevePlacer] FIXED: Successfully placed sleeve {sleeveInstance.Id} for duct {ductElementId}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DuctSleevePlacer] ERROR: Exception placing sleeve for duct {ductElementId}: {ex.Message}");
                DebugLogger.Error($"[DuctSleevePlacer] Stack trace: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// REDUNDANT METHOD - Replaced by PlaceDuctSleeveOptimized
        /// Places a duct sleeve with pre-calculated orientation to avoid timing bugs
        /// </summary>
        /*
        public void PlaceDuctSleeveWithOrientation(Duct duct, XYZ intersection, double width, double height, 
            XYZ ductDirection, XYZ preCalculatedOrientation, FamilySymbol sleeveSymbol, Element hostElement, XYZ? faceNormal = null)
        {
            int ductElementId = (int)(duct?.Id?.IntegerValue ?? 0);
            try
            {
                DebugLogger.Log($"[DuctSleevePlacer] USING PRE-CALCULATED ORIENTATION: ({preCalculatedOrientation.X:F6},{preCalculatedOrientation.Y:F6},{preCalculatedOrientation.Z:F6})");
                
                if (duct == null)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] ERROR: Duct is null, cannot place sleeve");
                    return;
                }
                
                // Call the standard placement method but with the pre-calculated orientation
                PlaceDuctSleeve(duct, intersection, width, height, ductDirection, sleeveSymbol, hostElement, faceNormal, preCalculatedOrientation);
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Exception in PlaceDuctSleeveWithOrientation for duct {ductElementId}: {ex.Message}");
                    DebugLogger.Log($"[DuctSleevePlacer] Stack trace: {ex.StackTrace}");
                    DebugLogger.Error($"[DuctSleevePlacer] FAILURE: Exception during placement for duct {ductElementId}: {ex.Message}");
                    throw; // Re-throw to ensure error is not silently ignored
                }
        }
        */

        /// <summary>
        /// REDUNDANT METHOD - Replaced by PlaceDuctSleeveOptimized
        /// Places a duct sleeve at the intersection point with robust positioning
        /// </summary>
        /*
        public void PlaceDuctSleeve(Duct duct, XYZ intersection, double width, double height, 
            XYZ ductDirection, FamilySymbol sleeveSymbol, Element hostElement, XYZ? faceNormal = null, XYZ? preCalculatedOrientation = null)
        {
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.SetDuctLogFile();
            int ductElementId = (int)(duct?.Id?.IntegerValue ?? 0);
            
            // CRITICAL FIX: Declare transaction variable outside try block so it's accessible in catch
            Transaction? localTransaction = null;
            bool wasInTransaction = false;
            
            try
            {
                DebugLogger.Log($"[DuctSleevePlacer] DIAGNOSTIC: Called for ductId={ductElementId}, hostType={hostElement?.GetType().Name}, direction=({ductDirection.X:F3},{ductDirection.Y:F3},{ductDirection.Z:F3})");
                DebugLogger.Log($"[DuctSleevePlacer] === LOG FILE START: PlaceDuctSleeve called for duct {ductElementId} at intersection {intersection} ===");
                try {
                    var logFilePath = typeof(DebugLogger).GetMethod("GetLogFilePath")?.Invoke(null, null) as string;
                    DebugLogger.Log($"[DuctSleevePlacer] DebugLogger log file path: {logFilePath}");
                } catch (Exception ex) {
                    DebugLogger.Log($"[DuctSleevePlacer] Could not get log file path: {ex.Message}");
                }
                DebugLogger.Log($"[DuctSleevePlacer] PlaceDuctSleeve called for duct {ductElementId} at intersection {intersection}");

                if (duct == null || intersection == null || sleeveSymbol == null || hostElement == null)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Null parameter check failed");
                    DebugLogger.Error($"[DuctSleevePlacer] FAILURE: Null parameters provided for duct {ductElementId}");
                    return;
                }

                // Damper-in-wall filter: skip placement if a damper is found at the intersection (only for walls)
                if (hostElement is Wall hostWall && IsDamperAtIntersection(_doc, intersection, hostWall, ductElementId))
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Damper detected at intersection for duct {ductElementId}, skipping sleeve placement.");
                    DebugLogger.Warning($"[DuctSleevePlacer] FAILURE: Damper present at intersection for duct {ductElementId}");
                    return;
                }

                // Host-specific logic for depth, normal, and placement point
                double sleeveDepth = 0.0;
                XYZ n = XYZ.BasisX;
                XYZ placePoint = intersection;
                if (hostElement is Wall wall)
                {
                    double wallThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                    n = (faceNormal != null) ? faceNormal.Normalize() : GetWallNormal(wall, intersection).Normalize();
                    XYZ wallVector = n.Multiply(-wallThickness);
                    placePoint = intersection + wallVector.Multiply(0.0); // wall face (intersection point)
                    sleeveDepth = wallThickness;
                    // Wall stub filter: skip if duct stops before 1/4 wall depth
                    double ductEndDist = double.MaxValue;
                    var locCurve = duct.Location as LocationCurve;
                    if (locCurve != null)
                        ductEndDist = locCurve.Curve.GetEndPoint(1).DistanceTo(intersection);
                    double quarterWall = wallThickness * 0.25;
                    if (ductEndDist < quarterWall)
                    {
                        DebugLogger.Log($"[DuctSleevePlacer] Duct {ductElementId} ends before 1/4 wall depth (dist={ductEndDist:F3}, 1/4 wall={quarterWall:F3}), skipping sleeve placement.");
                        DebugLogger.Warning($"[DuctSleevePlacer] FAILURE: Duct ends before 1/4 wall depth for duct {ductElementId}");
                        return;
                    }
                }
                // === STRUCTURAL FLOOR LOGIC START ===
                else if (hostElement is Floor floor)
                {
                    // STRUCTURAL FLOOR GUARD: Duct sleeves should ONLY be placed in structural floors
                    Parameter structuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                    bool isStructural = structuralParam?.AsInteger() == 1;
                    
                    DebugLogger.Log($"[DuctSleevePlacer] STRUCTURAL CHECK: FloorId={floor.Id.IntegerValue}, IsStructural={isStructural}");
                    
                    if (!isStructural)
                    {
                        DebugLogger.Log($"[DuctSleevePlacer] ERROR: Non-structural floor {floor.Id.IntegerValue} passed to duct placer! This should have been filtered in the command. Skipping placement for duct {ductElementId}.");
                        DebugLogger.Warning($"[DuctSleevePlacer] FAILURE: Non-structural floor {floor.Id.IntegerValue} for duct {ductElementId}");
                        return;
                    }
                    
                    // Try to get the floor type from the linked document if available
                    ElementId typeId = floor.GetTypeId();
                    Element? floorType = null;
                    Document typeDoc = floor.Document;
                    if (typeDoc.IsLinked)
                    {
                        // If the host is from a linked doc, get the type from the linked doc
                        floorType = typeDoc.GetElement(typeId);
                        DebugLogger.Log($"[DuctSleevePlacer] Host floor is from linked document: {typeDoc.Title}");
                    }
                    else
                    {
                        // If not linked, get from active doc
                        floorType = _doc.GetElement(typeId);
                    }
                    Parameter? thicknessParam = null;
                    if (floorType != null)
                    {
                        thicknessParam = floorType.LookupParameter("Thickness")
                            ?? floorType.LookupParameter("Depth")
                            ?? floorType.LookupParameter("Default Thickness");
                    }
                    if (thicknessParam != null && thicknessParam.StorageType == StorageType.Double)
                    {
                        sleeveDepth = thicknessParam.AsDouble();
                        DebugLogger.Log($"[DuctSleevePlacer] Floor thickness detected (param: {thicknessParam.Definition.Name}): {UnitUtils.ConvertFromInternalUnits(sleeveDepth, UnitTypeId.Millimeters):F1}mm");
                    }
                    else
                    {
                        // Log all available type parameters for debugging
                        if (floorType != null)
                        {
                            DebugLogger.Log($"[DuctSleevePlacer] Floor thickness parameter not found. Listing all type parameters:");
                            foreach (Parameter p in floorType.Parameters)
                            {
                                string val = p.HasValue ? p.AsValueString() : "<no value>";
                                DebugLogger.Log($"  - {p.Definition.Name}: {val} (StorageType={p.StorageType})");
                            }
                        }
                        sleeveDepth = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters); // fallback
                        DebugLogger.Log($"[DuctSleevePlacer] Floor thickness parameter not found, using fallback 500mm");
                    }
                    n = XYZ.BasisZ;
                    placePoint = intersection; // no offset for floor
                    DebugLogger.Log($"[DuctSleevePlacer] Duct direction for floor sleeve: ({ductDirection.X:F3}, {ductDirection.Y:F3}, {ductDirection.Z:F3})");
                }
                else if (hostElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    var framingType = famInst.Symbol;
                    var bParam = framingType.LookupParameter("b");
                    if (bParam != null && bParam.StorageType == StorageType.Double)
                        sleeveDepth = bParam.AsDouble();
                    else
                        sleeveDepth = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters); // fallback
                    var loc = famInst.Location as LocationCurve;
                    n = loc != null && loc.Curve is Line line ? line.Direction.CrossProduct(XYZ.BasisZ).Normalize() : XYZ.BasisY;
                    placePoint = intersection; // no offset for framing
                }
                else
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Unsupported host type for duct {ductElementId}, skipping");
                    DebugLogger.Warning($"[DuctSleevePlacer] FAILURE: Unsupported host type for duct {ductElementId}");
                    return;
                }

                // Ensure symbol is active
                if (!sleeveSymbol.IsActive)
                    sleeveSymbol.Activate();

                // Find the nearest or lowest level for placement (required for workplane-based family)
                Level level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(lvl => Math.Abs(lvl.Elevation - placePoint.Z))
                    .FirstOrDefault();
                if (level == null)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] No level found for placement");
                    DebugLogger.Error($"[DuctSleevePlacer] FAILURE: No level found for placement for duct {ductElementId}");
                    return;
                }

                // CRITICAL: Prevent accidental placement at (0,0,0) or far from intersection due to logic or data errors.
                double mmDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(placePoint), UnitTypeId.Millimeters);
                double depthMM = UnitUtils.ConvertFromInternalUnits(sleeveDepth, UnitTypeId.Millimeters);
                double allowedOffset = (depthMM * 0.5) + 5.0; // Allow up to half depth (mm) plus 5mm
                bool isAtOrigin = (Math.Abs(placePoint.X) < 0.001 && Math.Abs(placePoint.Y) < 0.001 && Math.Abs(placePoint.Z) < 0.001);
                bool isTooFar = mmDistance > allowedOffset;
                DebugLogger.Log($"[DuctSleevePlacer] DEBUG: intersection=({intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}), placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}), mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                if (isAtOrigin || isTooFar)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] ERROR: placePoint invalid for duct {ductElementId}: isAtOrigin={isAtOrigin}, isTooFar={isTooFar}, mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}. Skipping placement. Intersection: [{intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}], placePoint: [{placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}]");
                    DebugLogger.Warning($"[DuctSleevePlacer] FAILURE: Invalid placePoint for duct {ductElementId}: isAtOrigin={isAtOrigin}, isTooFar={isTooFar}, mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                    return;
                }

                DebugLogger.Log($"[DuctSleevePlacer] Intended centerline placement: placePoint=[{placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}]");
                // Log the actual width, height, and direction being used for placement
                DebugLogger.Log($"[DuctSleevePlacer] Placing sleeve: width={width}, height={height}, intersection=({intersection.X},{intersection.Y},{intersection.Z})");
                DebugLogger.Log($"[DuctSleevePlacer] Duct direction: ({ductDirection.X}, {ductDirection.Y}, {ductDirection.Z})");
                // Use conditional logging to avoid hardcoded file path issues
                string compareLogPath = "C:\\JSE_CSharp_Projects\\JSE_RevitAddin_MEP_OPENINGS\\JSE_RevitAddin_MEP_OPENINGS\\Log\\MEP_Sleeve_Placement_Compare.log";
                if (DebugLogger.IsEnabled)
                {
                    try
                    {
                        LoggingConfiguration.ConditionalAppendAllText(compareLogPath, $"[DuctSleevePlacer] Duct direction: ({ductDirection.X}, {ductDirection.Y}, {ductDirection.Z})\n");
                        LoggingConfiguration.ConditionalAppendAllText(compareLogPath, $"[DuctSleevePlacer] Placing sleeve: width={width}, height={height}, intersection=({intersection.X},{intersection.Y},{intersection.Z})\n");
                    }
                    catch (Exception logEx)
                    {
                        DebugLogger.Log($"[DuctSleevePlacer] Could not write to compare log: {logEx.Message}");
                    }
                }
                
                // CRITICAL FIX: Check if we're already in a transaction, if not create one
                wasInTransaction = _doc.IsModifiable;
                
                if (!wasInTransaction)
                {
                    localTransaction = new Transaction(_doc, $"Place Duct Sleeve {ductElementId}");
                    localTransaction.Start();
                    DebugLogger.Log($"[DuctSleevePlacer] Created local transaction for duct {ductElementId}");
                }
                else
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Using existing transaction for duct {ductElementId}");
                }
                
                FamilyInstance instance = _doc.Create.NewFamilyInstance(
                    placePoint,  // CRITICAL FIX: Use placePoint instead of intersection for proper placement
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);

                var settings = ApplicationProfileService.Instance.GetCurrentSettings();

                if (settings.CutOpeningWithHosts)
                {
                    try
                    {
                        InstanceVoidCutUtils.AddInstanceVoidCut(_doc, hostElement, instance);
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"[DuctSleevePlacer] Failed to cut host element: {ex.Message}");
                    }
                }
                // Set HostOrientation parameter for the new sleeve directly for wall/floor/framing
                string hostOrientationToSet = "";
                if (hostElement is Wall)
                {
                    // Use wall normal to determine orientation (X or Y)
                    // If wall normal is closer to X, orientation is X; if closer to Y, orientation is Y
                    double absX = Math.Abs(n.X);
                    double absY = Math.Abs(n.Y);
                    if (absX > absY)
                        hostOrientationToSet = "X";
                    else if (absY > absX)
                        hostOrientationToSet = "Y";
                    else
                        hostOrientationToSet = "Unknown";
                }
                else if (hostElement is Floor)
                {
                    hostOrientationToSet = "FloorHosted";
                }
                else if (hostElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // Use framing direction to determine orientation (X or Y)
                    var locationCurve = famInst.Location as LocationCurve;
                    if (locationCurve != null)
                    {
                        var curve = locationCurve.Curve as Line;
                        if (curve != null)
                        {
                            var direction = curve.Direction;
                            double absX = Math.Abs(direction.X);
                            double absY = Math.Abs(direction.Y);
                            if (absX > absY)
                                hostOrientationToSet = "X";
                            else if (absY > absX)
                                hostOrientationToSet = "Y";
                            else
                                hostOrientationToSet = "Unknown";
                        }
                        else
                        {
                            hostOrientationToSet = "Unknown";
                        }
                    }
                    else
                    {
                        hostOrientationToSet = "Unknown";
                    }
                }
                // Set the parameter if it exists and is writable
                var hostOrientationParam = instance.LookupParameter("HostOrientation");
                if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                {
                    hostOrientationParam.Set(hostOrientationToSet);
                    DebugLogger.Log($"[DuctSleevePlacer] HostOrientation set to '{hostOrientationToSet}' for sleeveId={instance.Id.IntegerValue}");
                }
                else
                {
                    DebugLogger.Log($"[DuctSleevePlacer] HostOrientation parameter not found or read-only for sleeveId={instance.Id.IntegerValue}");
                }
                // Explicitly log HostOrientation value after setting
                string hostOrientationValue = hostOrientationParam != null ? hostOrientationParam.AsString() : "<null>";
                DebugLogger.Log($"[DuctSleevePlacer] HostOrientation after set: '{hostOrientationValue}' for sleeveId={instance.Id.IntegerValue}");
                // Explicitly set the Level parameter for schedule consistency
                Parameter levelParam = instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                if (levelParam != null && !levelParam.IsReadOnly)
                {
                    levelParam.Set(level.Id);
                }
                else
                {
                    var levelByName = instance.LookupParameter("Level");
                    if (levelByName != null && !levelByName.IsReadOnly)
                        levelByName.Set(level.Id);
                }
                
                // Use helper for clearance
                // Use width and height as passed in (already includes clearance from command)
                double widthMM = UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters);
                double heightMM = UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters);
                DebugLogger.Log($"[DuctSleevePlacer] About to set Width: {width} (internal), {widthMM}mm; Height: {height} (internal), {heightMM}mm for duct {ductElementId}");

                // Set parameters with validation before any rotation
                var (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(width, height);
                SetParameterSafely(instance, "Width", roundedWidth, ductElementId);
                SetParameterSafely(instance, "Height", roundedHeight, ductElementId);
                SetParameterSafely(instance, "Depth", sleeveDepth, ductElementId); // from type param
                
                // Single regenerate after all parameters are set
                try
                {
                    _doc.Regenerate();
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Warning during regenerate: {ex.Message}");
                }
                
                DebugLogger.Log($"[DuctSleevePlacer] PLACED: ductId={ductElementId}, sleeveId={instance.Id.IntegerValue}, at {intersection}");

                // Only rotate for Y-axis ducts
                // Align sleeve to duct direction for ALL host types
                DebugLogger.Log($"[DuctSleevePlacer] hostElement type: {hostElement?.GetType().FullName}, category: {hostElement?.Category?.Name}, id: {hostElement?.Id}, family: {(hostElement as FamilyInstance)?.Symbol?.FamilyName}");
                bool isFloorHost = hostElement is Floor
                    || (hostElement is FamilyInstance fi && fi.Category != null && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors);
                
                try
                {
                    LocationPoint? loc = instance.Location as LocationPoint;
                    if (loc == null)
                    {
                        DebugLogger.Log("[DuctSleevePlacer] ERROR: instance.Location is not a LocationPoint - cannot rotate.");
                        return;
                    }

                    // For floor-hosted sleeves: always rotate using preCalculatedOrientation (command already filtered)
                    if (isFloorHost && preCalculatedOrientation != null)
                    {
                        DebugLogger.Log($"[DuctSleevePlacer] FLOOR: Rotating using pre-calculated orientation: ({preCalculatedOrientation.X:F6},{preCalculatedOrientation.Y:F6},{preCalculatedOrientation.Z:F6})");
                        double sleeveAngle = Math.Atan2(preCalculatedOrientation.Y, preCalculatedOrientation.X);
                        double sleeveAngleDegrees = sleeveAngle * 180 / Math.PI;
                        DebugLogger.Log($"[DuctSleevePlacer] FLOOR: Rotation angle: {sleeveAngleDegrees:F1} degrees");
                        Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, sleeveAngle);
                        DebugLogger.Log($"[DuctSleevePlacer] FLOOR: Applied rotation of {sleeveAngleDegrees:F1} degrees");
                    }
                    // For floor hosts with no preCalculatedOrientation: no rotation (command determined X-oriented)
                    else if (isFloorHost)
                    {
                        DebugLogger.Log("[DuctSleevePlacer] FLOOR: No rotation - command determined X-oriented duct");
                    }
                    // For walls/structural framing: check if Y-axis orientation
                    else
                    {
                        bool isYAxisDuct = Math.Abs(ductDirection.Y) > Math.Abs(ductDirection.X);
                        DebugLogger.Log($"[DuctSleevePlacer] WALL/FRAMING: isYAxisDuct={isYAxisDuct}, direction=({ductDirection.X:F3},{ductDirection.Y:F3},{ductDirection.Z:F3})");
                        
                        if (isYAxisDuct)
                        {
                            double rotationAngle = Math.PI / 2; // 90 degrees
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, rotationAngle);
                            DebugLogger.Log("[DuctSleevePlacer] WALL/FRAMING: Rotated Y-axis duct sleeve 90 degrees");
                        }
                        else
                        {
                            DebugLogger.Log("[DuctSleevePlacer] WALL/FRAMING: X-axis duct - no rotation needed");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Error during sleeve alignment: {ex.Message}");
                    DebugLogger.Log($"[DuctSleevePlacer] Stack trace: {ex.StackTrace}");
                }

                // Validate final position and log offset from centerline
                LocationPoint? locationPoint = instance.Location as LocationPoint;
                XYZ finalPosition;
                if (locationPoint == null || locationPoint.Point == null || 
                    (Math.Abs(locationPoint.Point.X) < 0.001 && Math.Abs(locationPoint.Point.Y) < 0.001 && Math.Abs(locationPoint.Point.Z) < 0.001))
                {
                    // If the location point is null or at origin, use the placement point as fallback
                    finalPosition = placePoint;
                    DebugLogger.Log($"[DuctSleevePlacer] Warning: Location retrieval returned null or origin. Using placement point as fallback.");
                }
                else
                {
                    finalPosition = locationPoint.Point;
                }

                // Calculate and log the offset from the intended centerline
                double offsetX = Math.Abs(finalPosition.X - placePoint.X);
                double offsetY = Math.Abs(finalPosition.Y - placePoint.Y);
                double offsetZ = Math.Abs(finalPosition.Z - placePoint.Z);
                double totalOffset = Math.Sqrt(offsetX * offsetX + offsetY * offsetY + offsetZ * offsetZ);

                DebugLogger.Log($"[DuctSleevePlacer] Placement validation:");
                DebugLogger.Log($"  - Intended centerline: [{placePoint.X:F3}, {placePoint.Y:F3}, {placePoint.Z:F3}]");
                DebugLogger.Log($"  - Actual placement: [{finalPosition.X:F3}, {finalPosition.Y:F3}, {finalPosition.Z:F3}]");
                DebugLogger.Log($"  - Offset: X={offsetX:F3}, Y={offsetY:F3}, Z={offsetZ:F3}, Total={totalOffset:F3}");

                if (totalOffset > 0.001)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] WARNING: Sleeve placement is not at the centerline. Offset detected.");
                }
                else
                {
                    DebugLogger.Log($"[DuctSleevePlacer] SUCCESS: Sleeve placement is at the centerline.");
                }
                
                double finalWidth = GetParameterValue(instance, "Width");
                double finalHeight = GetParameterValue(instance, "Height");
                double finalDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(finalPosition), UnitTypeId.Millimeters);
                DebugLogger.Log($"Duct {ductElementId} - Intersection vs Sleeve position distance: {finalDistance:F1}mm");
                DebugLogger.Log($"  - Intersection: [{intersection.X:F3}, {intersection.Y:F3}, {intersection.Z:F3}]");
                DebugLogger.Log($"  - Sleeve pos: [{finalPosition.X:F3}, {finalPosition.Y:F3}, {finalPosition.Z:F3}]");
                DebugLogger.Info($"[DuctSleevePlacer] SUCCESS: Placed sleeve for duct {ductElementId} -> sleeveId={(int)instance.Id.IntegerValue}, width={finalWidth:F1}mm, height={finalHeight:F1}mm, pos=({finalPosition.X:F3},{finalPosition.Y:F3},{finalPosition.Z:F3})");
                
                // Get reference level from host element (duct)
                Level? refLevel = HostLevelHelper.GetHostReferenceLevel(_doc, duct);
                if (refLevel != null)
                {
                    Parameter schedLevelParam = instance.LookupParameter("Schedule Level");
                    if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                    {
                        schedLevelParam.Set(refLevel.Id);
                        DebugLogger.Log($"[DuctSleevePlacer] Set Schedule Level to {refLevel.Name} for duct {ductElementId}");
                    }
                }
                
                // CRITICAL FIX: Commit local transaction if we created one
                if (localTransaction != null)
                {
                    localTransaction.Commit();
                    DebugLogger.Log($"[DuctSleevePlacer] Committed local transaction for duct {ductElementId}");
                }
            }
            catch (Exception ex)
            {
                // CRITICAL FIX: Rollback local transaction if we created one
                if (localTransaction != null)
                {
                    localTransaction.RollBack();
                    DebugLogger.Log($"[DuctSleevePlacer] Rolled back local transaction for duct {ductElementId} due to error");
                }
                
                DebugLogger.Log($"[DuctSleevePlacer] Exception in PlaceDuctSleeve for duct {ductElementId}: {ex.Message}");
                DebugLogger.Log($"[DuctSleevePlacer] Stack trace: {ex.StackTrace}");
                DebugLogger.Error($"[DuctSleevePlacer] FAILURE: Exception during placement for duct {ductElementId}: {ex.Message}");
                throw; // Re-throw to ensure error is not silently ignored
            }
        }
        */

        /// <summary>
        /// REDUNDANT METHOD - Replaced by PlaceDuctSleeveOptimized
        /// Static helper method for compatibility
        /// DO NOT change this signature or logic unless you are updating ALL callers.
        /// </summary>
        /*
        public static void PlaceDuctSleeveStatic(Document doc, Duct duct, XYZ intersection, double width, double height, 
            XYZ ductDirection, FamilySymbol sleeveSymbol, Wall hostWall, XYZ? faceNormal = null)
        {
            var placer = new DuctSleevePlacer(doc);
            placer.PlaceDuctSleeve(duct, intersection, width, height, ductDirection, sleeveSymbol, hostWall, faceNormal);
        }
        */

        private void SetParameterSafely(FamilyInstance instance, string paramName, double value, int ductElementId)
        {
            // DO NOT REMOVE: This ensures robust parameter setting and logging for all placements.
            try
            {
                Parameter param = instance.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                {
                    // Validate value is reasonable (not zero, negative, or extremely large)
                    if (value <= 0.0)
                    {
                        DebugLogger.Log($"[DuctSleevePlacer] WARNING: Invalid {paramName} value {value} for duct {ductElementId} - skipping");
                        return;
                    }
                    
                    double valueInMm = UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters);
                    if (valueInMm > 10000.0) // Sanity check: nothing should be larger than 10 meters
                    {
                        DebugLogger.Log($"[DuctSleevePlacer] WARNING: Extremely large {paramName} value {valueInMm:F1}mm for duct {ductElementId} - skipping");
                        return;
                    }
                    
                    DebugLogger.Log($"[DuctSleevePlacer] Setting {paramName} to {valueInMm:F1}mm (internal: {value:F6}) for duct {ductElementId}");

                    // Value is already in internal units (feet), set directly
                    param.Set(value);

                    // Verify the set value
                    double actualValue = param.AsDouble();
                    double actualValueInMm = UnitUtils.ConvertFromInternalUnits(actualValue, UnitTypeId.Millimeters);
                    DebugLogger.Log($"[DuctSleevePlacer] Verified {paramName} set to {actualValueInMm:F1}mm for duct {ductElementId}");
                }
                else
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Parameter {paramName} not found or read-only for duct {ductElementId}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DuctSleevePlacer] Failed to set {paramName} for duct {ductElementId}: {ex.Message}");
                // Don't throw - continue with placement even if parameter setting fails
            }
        }

        private double GetParameterValue(FamilyInstance instance, string paramName)
        {
            // DO NOT REMOVE: Used for robust logging and validation.
            try
            {
                Parameter param = instance.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    return UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DuctSleevePlacer] Failed to get {paramName}: {ex.Message}");
            }
            return 0.0;
        }

        private double GetParameterValueInternalUnits(FamilyInstance instance, string paramName)
        {
            try
            {
                Parameter param = instance.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    return param.AsDouble(); // Return in internal units (feet)
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DuctSleevePlacer] Failed to get {paramName} in internal units: {ex.Message}");
            }
            return 0.0;
        }

        private XYZ GetWallNormal(Wall wall, XYZ point)
        {
            // Used for orientation only, not for hosting. Do not use wall as host.
            try
            {
                // Get the wall's location curve
                LocationCurve? locationCurve = wall.Location as LocationCurve;
                if (locationCurve != null && locationCurve.Curve is Line line)
                {
                    XYZ direction = line.Direction.Normalize();
                    XYZ normal = new XYZ(-direction.Y, direction.X, 0).Normalize();
                    return normal;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DuctSleevePlacer] Failed to get wall normal: {ex.Message}");
            }
            
            // Default to X-axis normal
            return new XYZ(1, 0, 0);
        }

        /// <summary>
        /// REDUNDANT METHOD - Only used by old PlaceDuctSleeve method
        /// Checks if a damper is present at the intersection point (active document only - no linked document access)
        /// </summary>
        /*
        private bool IsDamperAtIntersection(Document doc, XYZ intersection, Wall linkedWall, int ductElementId)
        {
            double searchRadius = 0.2; // 200mm for broader catch
            // OPTIMIZATION: Only search in active document, not linked documents
            var collector = new FilteredElementCollector(_doc) // Use _doc (active document) instead of doc parameter
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_DuctAccessory);

            foreach (FamilyInstance fi in collector)
            {
                LocationPoint? loc = fi.Location as LocationPoint;
                if (loc == null) continue;

                XYZ? locPoint = loc.Point;
                if (locPoint == null) continue;

                double dist = locPoint.DistanceTo(intersection);

                if (dist < searchRadius)
                {
                    string famName = fi.Symbol.FamilyName.ToLower();
                    string typeName = fi.Symbol.Name.ToLower();
                    DebugLogger.Log($"[DuctSleevePlacer][DamperCheck] Duct {ductElementId}: Found damper {fi.Id.IntegerValue} at dist={dist:F3} (within radius), family={famName}, type={typeName}. Skipping sleeve placement.");
                    return true;
                }
            }
            
            return false;
        }
        */

    /// <summary>
    /// Find nearest level for explicit placement (like your reference code)
    /// </summary>
    private Level FindNearestLevel(XYZ point)
    {
        try
        {
            // First, log all available levels for debugging
            var allLevels = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();
            
            DebugLogger.Log($"[DuctSleevePlacer] Host document has {allLevels.Count} levels:");
            foreach (var level in allLevels)
            {
                DebugLogger.Log($"  - {level.Name} at elevation {level.Elevation:F2}");
            }
            
            DebugLogger.Log($"[DuctSleevePlacer] Placement point Z = {point.Z:F2}");
            
            if (allLevels.Count == 0)
            {
                throw new InvalidOperationException(
                    "No levels found in host document. Cannot place sleeve without level reference.");
            }

            // Find the nearest level by elevation
            var nearestLevel = allLevels
                .OrderBy(lvl => Math.Abs(lvl.Elevation - point.Z))
                .FirstOrDefault();
            
            if (nearestLevel == null)
            {
                throw new InvalidOperationException(
                    "Could not determine nearest level for sleeve placement.");
            }
            
            DebugLogger.Log($"[DuctSleevePlacer] Using nearest level: {nearestLevel.Name} (elevation: {nearestLevel.Elevation:F2})");
            return nearestLevel;
        }
        catch (Exception ex)
        {
            DebugLogger.Error($"[DuctSleevePlacer] Error finding nearest level: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Get or create the appropriate level using LevelMonitoringService for proper level association
    /// </summary>
    private Level GetOrCreateLevelInActiveDocument(string levelName, double levelElevation)
    {
        try
        {
            // Use LevelMonitoringService to get or create the appropriate level
            var levelMonitoringService = new LevelMonitoringService(_doc);
            
            // Try to get level by elevation first (this will find closest or create new if needed)
            var level = levelMonitoringService.GetLevelForElevation(levelElevation, levelName);
            
            if (level != null)
            {
                DebugLogger.Log($"[DuctSleevePlacer] Using level '{level.Name}' (elevation: {level.Elevation:F2}) for placement at Z={levelElevation:F2}");
                return level;
            }

            // Fallback: find closest existing level
            var allLevels = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            if (allLevels.Count > 0)
            {
                var closestLevel = allLevels.OrderBy(l => Math.Abs(l.Elevation - levelElevation)).First();
                DebugLogger.Warning($"[DuctSleevePlacer] LevelMonitoringService failed, using closest level '{closestLevel.Name}' (elevation: {closestLevel.Elevation:F2}) for placement at Z={levelElevation:F2}");
                return closestLevel;
            }

            DebugLogger.Error($"[DuctSleevePlacer] No levels found in active document and LevelMonitoringService failed");
            return null;
        }
        catch (Exception ex)
        {
            DebugLogger.Error($"[DuctSleevePlacer] Error getting level for placement Z={levelElevation:F2}: {ex.Message}");
            return null;
        }
    }

        /// <summary>
        /// Calculate element depth based on structural element type
        /// </summary>
        private double CalculateElementDepth(Element structuralElement)
        {
            try
            {
                if (structuralElement == null)
                {
                    DebugLogger.Warning($"[DuctSleevePlacer] CalculateElementDepth: structuralElement is null");
                    return 0.0;
                }

                DebugLogger.Log($"[DuctSleevePlacer] CalculateElementDepth: Processing element {structuralElement.Id.IntegerValue} of type {structuralElement.GetType().Name}");

                if (structuralElement is Wall wall)
                {
                    // For walls: Get wall thickness parameter
                    var thicknessParam = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                    var thickness = thicknessParam?.AsDouble() ?? 0.0;
                    
                    if (thickness > 0)
                    {
                        var thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                        DebugLogger.Log($"[DuctSleevePlacer] CalculateElementDepth: Wall {wall.Id.IntegerValue} thickness = {thicknessMm:F1}mm");
                        return thickness;
                    }
                    else
                    {
                        DebugLogger.Warning($"[DuctSleevePlacer] CalculateElementDepth: Wall {wall.Id.IntegerValue} thickness parameter not found or zero");
                        return 0.0;
                    }
                }
                else if (structuralElement is Floor floor)
                {
                    // For floors: Get floor thickness parameter
                    var floorType = floor.FloorType;
                    var thicknessParam = floorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM);
                    var thickness = thicknessParam?.AsDouble() ?? 0.0;
                    
                    if (thickness > 0)
                    {
                        var thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                        DebugLogger.Log($"[DuctSleevePlacer] CalculateElementDepth: Floor {floor.Id.IntegerValue} thickness = {thicknessMm:F1}mm");
                        return thickness;
                    }
                    else
                    {
                        DebugLogger.Warning($"[DuctSleevePlacer] CalculateElementDepth: Floor {floor.Id.IntegerValue} thickness parameter not found or zero");
                        return 0.0;
                    }
                }
                else if (structuralElement is FamilyInstance famInst && 
                         famInst.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // For structural framing: Try common parameter names
                    var widthParam = famInst.LookupParameter("Width") ?? 
                                   famInst.LookupParameter("b") ?? 
                                   famInst.LookupParameter("Depth") ??
                                   famInst.LookupParameter("Thickness");
                    
                    if (widthParam != null && widthParam.AsDouble() > 0)
                    {
                        var thickness = widthParam.AsDouble();
                        var thicknessMm = UnitUtils.ConvertFromInternalUnits(thickness, UnitTypeId.Millimeters);
                        DebugLogger.Log($"[DuctSleevePlacer] CalculateElementDepth: Structural framing {famInst.Id.IntegerValue} thickness = {thicknessMm:F1}mm");
                        return thickness;
                    }
                        
                    DebugLogger.Warning($"[DuctSleevePlacer] CalculateElementDepth: Cannot determine structural framing depth: no suitable parameter found for element ID {famInst.Id}");
                    return 0.0;
                }
                
                DebugLogger.Warning($"[DuctSleevePlacer] CalculateElementDepth: Unsupported structural element type: {structuralElement.GetType().Name}");
                return 0.0;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[DuctSleevePlacer] CalculateElementDepth: Error calculating element depth: {ex.Message}");
                return 0.0;
            }
        }
    }
}
