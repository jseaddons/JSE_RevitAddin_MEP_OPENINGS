using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;  // for CableTray
using Autodesk.Revit.DB.Structure;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Places rectangular sleeves for Cable Tray elements with 50mm clearance all around.
    /// </summary>
    public class CableTraySleevePlacer
    {
        private readonly Document _doc;

        public CableTraySleevePlacer(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// Places a cable tray sleeve with pre-calculated orientation to avoid timing bugs
        /// </summary>
        public FamilyInstance PlaceCableTraySleeveWithOrientation(CableTray tray, XYZ intersection, double width, double height, 
            XYZ direction, XYZ? preCalculatedOrientation, FamilySymbol sleeveSymbol, Element hostElement)
        {
            int cableTrayId = (int)(tray?.Id?.Value ?? 0);
            try
            {
                DebugLogger.Log($"[CableTraySleevePlacer] USING PRE-CALCULATED ORIENTATION: {(preCalculatedOrientation != null ? $"({preCalculatedOrientation.X:F6},{preCalculatedOrientation.Y:F6},{preCalculatedOrientation.Z:F6})" : "null")}");
                
                if (tray == null)
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] ERROR: CableTray is null, cannot place sleeve");
                    return null;
                }
                
                // Call the standard placement method but with the pre-calculated orientation
                return PlaceCableTraySleeve(tray, intersection, width, height, direction, sleeveSymbol, hostElement, preCalculatedOrientation);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CableTraySleevePlacer] Exception in PlaceCableTraySleeveWithOrientation for cable tray {cableTrayId}: {ex.Message}");
                DebugLogger.Log($"[CableTraySleevePlacer] Stack trace: {ex.StackTrace}");
                throw; // Re-throw to ensure error is not silently ignored
            }
        }

        public FamilyInstance PlaceCableTraySleeve(
            CableTray tray,
            XYZ intersection,
            double width,
            double height,
            XYZ direction,
            FamilySymbol sleeveSymbol,
            Element hostElement,
            XYZ? preCalculatedOrientation = null)
        {
            DebugLogger.Log($"[CableTraySleevePlacer] ENTRY: PlaceCableTraySleeve called with trayId={(tray != null ? tray.Id.IntegerValue.ToString() : "null")}");
            int cableTrayId = (tray != null) ? (int)tray.Id.Value : 0;
            try
            {

            DebugLogger.Log($"[CableTraySleevePlacer] === LOG FILE START: PlaceCableTraySleeve called for cable tray {cableTrayId} at intersection {intersection} ===");
            DebugLogger.Log($"[CableTraySleevePlacer] PlaceCableTraySleeve called for cable tray {cableTrayId} at intersection {intersection}");

            if (intersection == null || sleeveSymbol == null || hostElement == null)
            {
                DebugLogger.Log($"[CableTraySleevePlacer] Null parameter check failed");
                return null;
            }

            // --- NEW: Prevent placement in invisible/hidden linked files ---
            // If hostElement is from a linked file, check if the link is visible in the active view
            Document hostDoc = hostElement.Document;
            if (hostDoc.IsLinked)
            {
                // Find the RevitLinkInstance in the main doc that corresponds to this linked doc
                var linkInstance = new FilteredElementCollector(_doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .FirstOrDefault(link => link.GetLinkDocument() == hostDoc);
                if (linkInstance != null)
                {
                    var activeView = _doc.ActiveView;
                    bool isHidden = activeView.GetCategoryHidden(linkInstance.Category.Id) || linkInstance.IsHidden(activeView);
                    if (isHidden)
                    {
                        DebugLogger.Log($"[CableTraySleevePlacer] SKIP: Host element is from a hidden/invisible linked file. No sleeve will be placed.");
                        return null;
                    }
                }
            }

                // Host-specific logic for depth, normal, and placement point
                double sleeveDepth = 0.0;
                XYZ n = XYZ.BasisX;
                XYZ placePoint = intersection;
                if (hostElement is Wall wall)
                {
                    double wallThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                    n = GetWallNormal(wall, intersection).Normalize();
                    // Use robust centerline projection for XY, keep Z from intersection
                    XYZ centerlinePoint = GetWallCenterlinePoint(wall, intersection);
                    placePoint = new XYZ(centerlinePoint.X, centerlinePoint.Y, intersection.Z);
                    sleeveDepth = wallThickness;
                }
                else if (hostElement is Floor floor)
                {
                    ElementId typeId = floor.GetTypeId();
                    Element floorType = null;
                    Document typeDoc = floor.Document;
                    if (typeDoc.IsLinked)
                        floorType = typeDoc.GetElement(typeId);
                    else
                        floorType = _doc.GetElement(typeId);
                    Parameter thicknessParam = null;
                    if (floorType != null)
                    {
                        thicknessParam = floorType.LookupParameter("Thickness")
                            ?? floorType.LookupParameter("Depth")
                            ?? floorType.LookupParameter("Default Thickness");
                    }
                    if (thicknessParam != null && thicknessParam.StorageType == StorageType.Double)
                    {
                        sleeveDepth = thicknessParam.AsDouble();
                        DebugLogger.Log($"[CableTraySleevePlacer] Floor thickness detected (param: {thicknessParam.Definition.Name}): {UnitUtils.ConvertFromInternalUnits(sleeveDepth, UnitTypeId.Millimeters):F1}mm");
                    }
                    else
                    {
                        if (floorType != null)
                        {
                            DebugLogger.Log($"[CableTraySleevePlacer] Floor thickness parameter not found. Listing all type parameters:");
                            foreach (Parameter p in floorType.Parameters)
                            {
                                string val = p.HasValue ? p.AsValueString() : "<no value>";
                                DebugLogger.Log($"  - {p.Definition.Name}: {val} (StorageType={p.StorageType})");
                            }
                        }
                        sleeveDepth = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters); // fallback
                        DebugLogger.Log($"[CableTraySleevePlacer] Floor thickness parameter not found, using fallback 500mm");
                    }
                    n = XYZ.BasisZ;
                    placePoint = intersection; // no offset for floor
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
                    DebugLogger.Log($"[CableTraySleevePlacer] Unsupported host type for cable tray {cableTrayId}, skipping");
                    return null;
                }

                // Ensure symbol is active
                if (!sleeveSymbol.IsActive)
                    sleeveSymbol.Activate();

                // Find the nearest or lowest level for placement
                Level level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(lvl => Math.Abs(lvl.Elevation - placePoint.Z))
                    .FirstOrDefault();
                if (level == null)
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] No level found for placement");
                    return null;
                }

                // Validate placement point
                double mmDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(placePoint), UnitTypeId.Millimeters);
                double depthMM = UnitUtils.ConvertFromInternalUnits(sleeveDepth, UnitTypeId.Millimeters);
                double allowedOffset = (depthMM * 0.5) + 5.0;
                bool isAtOrigin = (Math.Abs(placePoint.X) < 0.001 && Math.Abs(placePoint.Y) < 0.001 && Math.Abs(placePoint.Z) < 0.001);
                bool isTooFar = mmDistance > allowedOffset;
                DebugLogger.Log($"[CableTraySleevePlacer] DEBUG: intersection=({intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}), placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}), mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                if (isAtOrigin || isTooFar)
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] ERROR: placePoint invalid for cable tray {cableTrayId}: isAtOrigin={isAtOrigin}, isTooFar={isTooFar}, mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}. Skipping placement.");
                    return null;
                }

                DebugLogger.Log($"[CableTraySleevePlacer] Intended centerline placement: placePoint=[{placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}]");
                DebugLogger.Log($"[CableTraySleevePlacer] Placing sleeve: width={width}, height={height}, intersection=({intersection.X},{intersection.Y},{intersection.Z})");
                DebugLogger.Log($"[CableTraySleevePlacer] Cable tray direction: ({direction.X}, {direction.Y}, {direction.Z})");
                string compareLogPath = "C:\\JSE_CSharp_Projects\\JSE_RevitAddin_MEP_OPENINGS\\JSE_RevitAddin_MEP_OPENINGS\\Log\\MEP_Sleeve_Placement_Compare.log";
                if (DebugLogger.IsEnabled)
                {
                    System.IO.File.AppendAllText(compareLogPath, $"[CableTraySleevePlacer] Cable tray direction: ({direction.X}, {direction.Y}, {direction.Z})\n");
                    System.IO.File.AppendAllText(compareLogPath, $"[CableTraySleevePlacer] Placing sleeve: width={width}, height={height}, intersection=({intersection.X},{intersection.Y},{intersection.Z})\n");
                }

                FamilyInstance instance = _doc.Create.NewFamilyInstance(
                    placePoint,
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);

                // Set HostOrientation parameter for the new sleeve directly for wall/floor/framing
                string hostOrientationToSet = "";
                if (hostElement is Wall)
                {
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
                else if (hostElement is FamilyInstance famInst2 && famInst2.Category != null && famInst2.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    var locationCurve = famInst2.Location as LocationCurve;
                    if (locationCurve != null)
                    {
                        var curve = locationCurve.Curve as Line;
                        if (curve != null)
                        {
                            var framingDir = curve.Direction;
                            double absX = Math.Abs(framingDir.X);
                            double absY = Math.Abs(framingDir.Y);
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
                var hostOrientationParam = instance.LookupParameter("HostOrientation");
                if (hostOrientationParam != null && !hostOrientationParam.IsReadOnly)
                {
                    hostOrientationParam.Set(hostOrientationToSet);
                    DebugLogger.Log($"[CableTraySleevePlacer] HostOrientation set to '{hostOrientationToSet}' for sleeveId={instance.Id.Value}");
                }
                else
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] HostOrientation parameter not found or read-only for sleeveId={instance.Id.Value}");
                }
                string hostOrientationValue = hostOrientationParam != null ? hostOrientationParam.AsString() : "<null>";
                DebugLogger.Log($"[CableTraySleevePlacer] HostOrientation after set: '{hostOrientationValue}' for sleeveId={instance.Id.Value}");

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

                DebugLogger.Log($"[CableTraySleevePlacer] hostElement type: {hostElement?.GetType().FullName}, category: {hostElement?.Category?.Name}, id: {hostElement?.Id}, family: {(hostElement as FamilyInstance)?.Symbol?.FamilyName}");
                bool isFloorHost = hostElement is Floor
                    || (hostElement is FamilyInstance fi && fi.Category != null && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Floors);
                // Use helper for clearance
                double clearance = JSE_RevitAddin_MEP_OPENINGS.Helpers.SleeveClearanceHelper.GetClearance(tray);
                double widthWithClearance = width + 2 * clearance;
                double heightWithClearance = height + 2 * clearance;
                // Set parameters with validation before any rotation
                SetParameterSafelyWithLog(instance, "Width", widthWithClearance, cableTrayId);
                SetParameterSafelyWithLog(instance, "Height", heightWithClearance, cableTrayId);
                SetParameterSafelyWithLog(instance, "Depth", sleeveDepth, cableTrayId);

                try
                {
                    _doc.Regenerate();
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] Warning during regenerate: {ex.Message}");
                }

                // Only apply rotation for floors (when preCalculatedOrientation is set)
                try
                {
                    if (hostElement is Floor && preCalculatedOrientation != null)
                    {
                        LocationPoint loc = instance.Location as LocationPoint;
                        if (loc == null)
                        {
                            DebugLogger.Log("[CableTraySleevePlacer] ERROR: instance.Location is not a LocationPoint - cannot rotate.");
                            return instance;
                        }
                        DebugLogger.Log($"[CableTraySleevePlacer] ROTATING: Using pre-calculated orientation from command: ({preCalculatedOrientation.X:F6},{preCalculatedOrientation.Y:F6},{preCalculatedOrientation.Z:F6})");
                        double sleeveAngle = Math.Atan2(preCalculatedOrientation.Y, preCalculatedOrientation.X);
                        double sleeveAngleDegrees = sleeveAngle * 180 / Math.PI;
                        DebugLogger.Log($"[CableTraySleevePlacer] Rotation angle: {sleeveAngleDegrees:F1} degrees for host type: {hostElement?.GetType().Name ?? "Unknown"}");
                        Line rotationAxis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, sleeveAngle);
                        DebugLogger.Log($"[CableTraySleevePlacer] Applied rotation of {sleeveAngleDegrees:F1} degrees (command determined Y-oriented)");
                    }
                    else
                    {
                        DebugLogger.Log($"[CableTraySleevePlacer] NO ROTATION: Wall/framing host or no preCalculatedOrientation. Sleeve placed as-is.");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] Error during sleeve alignment: {ex.Message}");
                }

                // Validate final position and log offset from centerline
                LocationPoint locationPoint = instance.Location as LocationPoint;
                XYZ finalPosition = locationPoint?.Point;
                if (finalPosition == null || (Math.Abs(finalPosition.X) < 0.001 && Math.Abs(finalPosition.Y) < 0.001 && Math.Abs(finalPosition.Z) < 0.001))
                {
                    finalPosition = placePoint;
                    DebugLogger.Log($"[CableTraySleevePlacer] Warning: Location retrieval returned null or origin. Using placement point as fallback.");
                }
                double offsetX = Math.Abs(finalPosition.X - placePoint.X);
                double offsetY = Math.Abs(finalPosition.Y - placePoint.Y);
                double offsetZ = Math.Abs(finalPosition.Z - placePoint.Z);
                double totalOffset = Math.Sqrt(offsetX * offsetX + offsetY * offsetY + offsetZ * offsetZ);
                DebugLogger.Log($"[CableTraySleevePlacer] Placement validation:");
                DebugLogger.Log($"  - Intended centerline: [{placePoint.X:F3}, {placePoint.Y:F3}, {placePoint.Z:F3}]");
                DebugLogger.Log($"  - Actual placement: [{finalPosition.X:F3}, {finalPosition.Y:F3}, {finalPosition.Z:F3}]");
                DebugLogger.Log($"  - Offset: X={offsetX:F3}, Y={offsetY:F3}, Z={offsetZ:F3}, Total={totalOffset:F3}");
                if (totalOffset > 0.001)
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] WARNING: Sleeve placement is not at the centerline. Offset detected.");
                }
                else
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] SUCCESS: Sleeve placement is at the centerline.");
                }

                double finalWidth = GetParameterValue(instance, "Width");
                double finalHeight = GetParameterValue(instance, "Height");
                double finalDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(finalPosition), UnitTypeId.Millimeters);
                DebugLogger.Log($"CableTray {cableTrayId} - Intersection vs Sleeve position distance: {finalDistance:F1}mm");
                DebugLogger.Log($"  - Intersection: [{intersection.X:F3}, {intersection.Y:F3}, {intersection.Z:F3}]");
                DebugLogger.Log($"  - Sleeve pos: [{finalPosition.X:F3}, {finalPosition.Y:F3}, {finalPosition.Z:F3}]");

                // Get reference level from host element (cable tray)
                Level refLevel = JSE_RevitAddin_MEP_OPENINGS.Helpers.HostLevelHelper.GetHostReferenceLevel(_doc, tray);
                if (refLevel != null)
                {
                    Parameter schedLevelParam = instance.LookupParameter("Schedule Level");
                    if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                    {
                        schedLevelParam.Set(refLevel.Id);
                        DebugLogger.Log($"[CableTraySleevePlacer] Set Schedule Level to {refLevel.Name} for cable tray {cableTrayId}");
                    }
                }

                return instance;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CableTraySleevePlacer] Exception in PlaceCableTraySleeve for cable tray {cableTrayId}: {ex.Message}");
                DebugLogger.Log($"[CableTraySleevePlacer] Stack trace: {ex.StackTrace}");
                return null;
            }
        }

        private void SetParameterSafelyWithLog(FamilyInstance instance, string paramName, double value, int elementId)
        {
            try
            {
                Parameter param = instance.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                {
                    if (value <= 0.0)
                    {
                        DebugLogger.Log($"[CableTraySleevePlacer] WARNING: Invalid {paramName} value {value} for cable tray {elementId} - skipping");
                        return;
                    }
                    double valueInMm = UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters);
                    if (valueInMm > 10000.0)
                    {
                        DebugLogger.Log($"[CableTraySleevePlacer] WARNING: Extremely large {paramName} value {valueInMm:F1}mm for cable tray {elementId} - skipping");
                        return;
                    }
                    DebugLogger.Log($"[CableTraySleevePlacer] Setting {paramName} to {valueInMm:F1}mm (internal: {value:F6}) for cable tray {elementId}");
                    param.Set(value);
                    double actualValue = param.AsDouble();
                    double actualValueInMm = UnitUtils.ConvertFromInternalUnits(actualValue, UnitTypeId.Millimeters);
                    DebugLogger.Log($"[CableTraySleevePlacer] Verified {paramName} set to {actualValueInMm:F1}mm for cable tray {elementId}");
                }
                else
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] Parameter {paramName} not found or read-only for cable tray {elementId}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CableTraySleevePlacer] Failed to set {paramName} for cable tray {elementId}: {ex.Message}");
            }
        }

        private double GetParameterValue(FamilyInstance instance, string paramName)
        {
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
                DebugLogger.Log($"[CableTraySleevePlacer] Failed to get {paramName}: {ex.Message}");
            }
            return 0.0;
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
            catch (Exception)
            {
            }

            // Default to X-axis normal
            return new XYZ(1, 0, 0);
        }

        private XYZ GetWallCenterlinePoint(Wall wall, XYZ intersectionPoint)
        {
            // Project intersection point onto wall centerline using wall's location curve
            try
            {
                if (wall == null)
                {
                    return intersectionPoint;
                }

                LocationCurve locationCurve = wall.Location as LocationCurve;
                if (locationCurve?.Curve is Line wallCenterline)
                {
                    // Project the intersection point onto the wall's centerline
                    IntersectionResult projectionResult = wallCenterline.Project(intersectionPoint);
                    XYZ centerlinePoint = projectionResult.XYZPoint;

                    // Preserve the Z coordinate from the original intersection
                    XYZ finalPoint = new XYZ(centerlinePoint.X, centerlinePoint.Y, intersectionPoint.Z);

                    double distanceFromOriginal = intersectionPoint.DistanceTo(finalPoint);

                    return finalPoint;
                }
                else
                {
                    return intersectionPoint;
                }
            }
            catch (Exception)
            {
                return intersectionPoint;
            }
        }

        /// <summary>
        /// Checks if a cable tray fitting is present at the intersection point.
        /// Enhanced: matches more fitting types, supports LocationCurve, and logs all suppression decisions.
        /// </summary>
        private bool IsFittingAtIntersection(Document doc, XYZ intersection, CableTray tray, int cableTrayId)
        {
            double searchRadius = UnitUtils.ConvertToInternalUnits(100.0, UnitTypeId.Millimeters); // 100mm like duct logic
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType();

            // Keywords for fittings (case-insensitive)
            string[] fittingKeywords = new[] { "fitting", "elbow", "tee", "cross", "junction", "bend", "transition", "reducer", "wye" };
            bool foundSuppression = false;
            foreach (FamilyInstance fi in collector)
            {
                var famName = fi.Symbol?.Family?.Name?.ToLowerInvariant() ?? "";
                var typeName = fi.Symbol?.Name?.ToLowerInvariant() ?? "";
                bool isFitting = fittingKeywords.Any(kw => famName.Contains(kw) || typeName.Contains(kw));
                if (!isFitting)
                {
                    continue;
                }

                double dist = double.MaxValue;
                string locationType = "unknown";
                if (fi.Location is LocationPoint locPt)
                {
                    dist = locPt.Point.DistanceTo(intersection);
                    locationType = "LocationPoint";
                }
                else if (fi.Location is LocationCurve locCurve)
                {
                    var curve = locCurve.Curve;
                    if (curve != null)
                    {
                        var proj = curve.Project(intersection);
                        if (proj != null)
                        {
                            dist = proj.XYZPoint.DistanceTo(intersection);
                            locationType = "LocationCurve";
                        }
                    }
                }
                else
                {
                    continue;
                }

                if (dist < searchRadius)
                {
                    foundSuppression = true;
                }
            }
            return foundSuppression;
        }
    }
}
