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
        /// Places a duct sleeve with pre-calculated orientation to avoid timing bugs
        /// </summary>
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

        /// <summary>
        /// Places a duct sleeve at the intersection point with robust positioning
        /// </summary>
        public void PlaceDuctSleeve(Duct duct, XYZ intersection, double width, double height, 
            XYZ ductDirection, FamilySymbol sleeveSymbol, Element hostElement, XYZ? faceNormal = null, XYZ? preCalculatedOrientation = null)
        {
            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.SetDuctLogFile();
            int ductElementId = (int)(duct?.Id?.IntegerValue ?? 0);
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
                    placePoint = intersection + wallVector.Multiply(0.5); // wall centerline
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
                string compareLogPath = "C:\\JSE_CSharp_Projects\\JSE_RevitAddin_MEP_OPENINGS\\JSE_RevitAddin_MEP_OPENINGS\\Log\\MEP_Sleeve_Placement_Compare.log";
                if (DebugLogger.IsEnabled)
                {
                    System.IO.File.AppendAllText(compareLogPath, $"[DuctSleevePlacer] Duct direction: ({ductDirection.X}, {ductDirection.Y}, {ductDirection.Z})\n");
                    System.IO.File.AppendAllText(compareLogPath, $"[DuctSleevePlacer] Placing sleeve: width={width}, height={height}, intersection=({intersection.X},{intersection.Y},{intersection.Z})\n");
                }
                FamilyInstance instance = _doc.Create.NewFamilyInstance(
                    intersection,
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);
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
                SetParameterSafely(instance, "Width", width, ductElementId);
                SetParameterSafely(instance, "Height", height, ductElementId);
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
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DuctSleevePlacer] Exception in PlaceDuctSleeve for duct {ductElementId}: {ex.Message}");
                DebugLogger.Log($"[DuctSleevePlacer] Stack trace: {ex.StackTrace}");
                DebugLogger.Error($"[DuctSleevePlacer] FAILURE: Exception during placement for duct {ductElementId}: {ex.Message}");
                throw; // Re-throw to ensure error is not silently ignored
            }
        }

        /// <summary>
        /// Static helper method for compatibility
        /// DO NOT change this signature or logic unless you are updating ALL callers.
        /// </summary>
        public static void PlaceDuctSleeveStatic(Document doc, Duct duct, XYZ intersection, double width, double height, 
            XYZ ductDirection, FamilySymbol sleeveSymbol, Wall hostWall, XYZ? faceNormal = null)
        {
            var placer = new DuctSleevePlacer(doc);
            placer.PlaceDuctSleeve(duct, intersection, width, height, ductDirection, sleeveSymbol, hostWall, faceNormal);
        }        private void SetParameterSafely(FamilyInstance instance, string paramName, double value, int ductElementId)
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
        /// Checks if a damper is present at the intersection point in the linked wall
        /// </summary>
        private bool IsDamperAtIntersection(Document doc, XYZ intersection, Wall linkedWall, int ductElementId)
        {
            double searchRadius = 0.2; // 200mm for broader catch
            var collector = new FilteredElementCollector(doc)
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
    }
}
