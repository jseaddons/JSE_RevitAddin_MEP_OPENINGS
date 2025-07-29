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

        public void PlaceCableTraySleeve(
            CableTray tray,
            XYZ intersection,
            double width,
            double height,
            XYZ direction,
            FamilySymbol sleeveSymbol,
            Wall hostWall)
        {
            // Debug start of placement
            DebugLogger.Log($"CableTrayPlacer: start placement for ID={tray?.Id.IntegerValue}, intersection={intersection}, width={width}, height={height}, direction={direction}");
            // Get the element ID for logging
            int cableTrayId = (tray != null) ? tray.Id.IntegerValue : 0;

            // Track the wall intersection in the logs
            SleeveLogManager.LogCableTrayWallIntersection(cableTrayId, intersection, width, height);

            // Log selected symbol ID before create
            DebugLogger.Log($"CableTrayPlacer: using symbol ID={sleeveSymbol.Id.IntegerValue}");

            // Check if a cable tray sleeve already exists at this location
            double tolerance = UnitUtils.ConvertToInternalUnits(2.0, UnitTypeId.Millimeters); // Use 2mm like other placers
            bool exists = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Family.Name.Contains("CableTrayOpeningOnWall"))
                .Any(fi => fi.GetTransform().Origin.DistanceTo(intersection) < tolerance);

            if (exists)
            {
                SleeveLogManager.LogCableTraySleeveFailure(cableTrayId, "Sleeve already exists at this location");
                return;
            }

            // Auto-select CT# type from CableTrayOpeningOnWall family
            var allSymbols = new FilteredElementCollector(_doc)
                                .OfClass(typeof(FamilySymbol))
                                .Cast<FamilySymbol>();
            var ctSymbol = allSymbols.FirstOrDefault(sym =>
                string.Equals(sym.Family.Name, "CableTrayOpeningOnWall", StringComparison.OrdinalIgnoreCase)
                && sym.Name.StartsWith("CT#", StringComparison.OrdinalIgnoreCase));
            if (ctSymbol == null)
            {
                SleeveLogManager.LogCableTraySleeveFailure(cableTrayId, "No CableTrayOpeningOnWall family with CT# type found");
                return;
            }
            sleeveSymbol = ctSymbol;

            // Determine wall thickness for sleeve depth only
            double wallThickness = 0.0;
            if (hostWall != null)
            {
                var thicknessParam = hostWall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                wallThickness = thicknessParam != null ? thicknessParam.AsDouble() : hostWall.Width;
            }

            // *** PROJECT INTERSECTION POINT TO WALL CENTERLINE ***
            XYZ placePoint = GetWallCenterlinePoint(hostWall, intersection);
            double sleeveDepth = wallThickness;

            DebugLogger.Log($"CableTrayPlacer: PROJECTION METHOD - original intersection={intersection}, projected centerline placePoint={placePoint}, sleeveDepth={sleeveDepth:F3}");

            // Add validation - check that placePoint is not at origin
            bool isAtOrigin = (Math.Abs(placePoint.X) < 0.001 && Math.Abs(placePoint.Y) < 0.001 && Math.Abs(placePoint.Z) < 0.001);

            DebugLogger.Log($"CableTrayPlacer: DEBUG: placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}), isAtOrigin={isAtOrigin}");
            if (isAtOrigin)
            {
                DebugLogger.Log($"CableTrayPlacer: ERROR: placePoint invalid for cable tray {cableTrayId}: isAtOrigin={isAtOrigin}. Skipping placement. Intersection: [{intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}]");
                SleeveLogManager.LogCableTraySleeveFailure(cableTrayId, $"Invalid placePoint: isAtOrigin={isAtOrigin}");
                return;
            }
            // Activate the symbol
            if (!sleeveSymbol.IsActive)
                sleeveSymbol.Activate();
            // Find a valid level for unhosted placement
            Level level = new FilteredElementCollector(_doc)
                              .OfClass(typeof(Level))
                              .Cast<Level>()
                              .FirstOrDefault();
            if (level == null)
            {
                SleeveLogManager.LogCableTraySleeveFailure(cableTrayId, "No level found for placement");
                return;
            }
            // Create family instance at intersection point (simplest approach)
            DebugLogger.Log($"CableTrayPlacer: creating family instance at intersection placePoint={placePoint}");

            FamilyInstance instance = _doc.Create.NewFamilyInstance(
                placePoint,
                sleeveSymbol,
                level,
                StructuralType.NonStructural);

            if (instance == null)
            {
                DebugLogger.Log($"CableTrayPlacer: NewFamilyInstance returned null for ID={cableTrayId}");
                SleeveLogManager.LogCableTraySleeveFailure(cableTrayId, "Failed to create instance");
                return;
            }
            DebugLogger.Log($"CableTrayPlacer: instance created with ID={instance.Id.IntegerValue}");

            // Add clearance to width and height (50mm each side)
            double clearanceMM = 50.0;
            double clearance = UnitUtils.ConvertToInternalUnits(clearanceMM, UnitTypeId.Millimeters);
            double finalWidth = width + (clearance * 2);
            double finalHeight = height + (clearance * 2);

            // Set parameters with validation
            SetParameterSafely(instance, "Width", finalWidth, cableTrayId);
            SetParameterSafely(instance, "Height", finalHeight, cableTrayId);
            SetParameterSafely(instance, "Depth", wallThickness, cableTrayId);

            // Add regenerate after setting parameters like duct sleeve
            try
            {
                _doc.Regenerate();
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"CableTrayPlacer: Warning during regenerate: {ex.Message}");
            }

            // Use original working rotation logic based on cable tray direction
            // Rotate if needed - Only rotate for Y-axis cable trays (when direction is primarily along Y)
            double absXRot = Math.Abs(direction.X);
            double absYRot = Math.Abs(direction.Y);

            DebugLogger.Log($"CableTrayPlacer: Cable tray direction={direction}, absX={absXRot:F3}, absY={absYRot:F3}");

            if (absYRot > absXRot)
            {
                // Cable tray runs along Y-axis, needs rotation
                LocationPoint loc = instance.Location as LocationPoint;
                if (loc != null)
                {
                    XYZ rotationPoint = loc.Point;
                    double angle = Math.PI / 2; // Always 90 degrees
                    Line rotAxis = Line.CreateBound(rotationPoint, rotationPoint.Add(XYZ.BasisZ));
                    try
                    {
                        ElementTransformUtils.RotateElement(_doc, instance.Id, rotAxis, angle);
                        DebugLogger.Log($"CableTrayPlacer: rotated instance by {angle * 180 / Math.PI} degrees for Y-axis cable tray");
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"CableTrayPlacer: rotation failed: {ex.Message}");
                    }
                }
            }
            else
            {
                DebugLogger.Log("CableTrayPlacer: no rotation applied (X-axis cable tray)");
            }
            // Log success without extra read tx
            SleeveLogManager.LogCableTraySleeveSuccess(cableTrayId, instance.Id.IntegerValue, finalWidth, finalHeight, instance.GetTransform().Origin);
        }

        private void SetParameterSafely(FamilyInstance instance, string paramName, double value, int elementId)
        {
            try
            {
                Parameter param = instance.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                {
                    // Validate value is reasonable
                    if (value <= 0.0)
                    {
                        DebugLogger.Log($"[CableTraySleevePlacer] WARNING: Invalid {paramName} value {value} for element {elementId} - skipping");
                        return;
                    }

                    double valueInMm = UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters);
                    if (valueInMm > 10000.0) // Sanity check
                    {
                        DebugLogger.Log($"[CableTraySleevePlacer] WARNING: Extremely large {paramName} value {valueInMm:F1}mm for element {elementId} - skipping");
                        return;
                    }

                    param.Set(value);
                    DebugLogger.Log($"[CableTraySleevePlacer] Set {paramName} to {valueInMm:F1}mm for element {elementId}");
                }
                else
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] Parameter {paramName} not found or read-only for element {elementId}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CableTraySleevePlacer] Failed to set {paramName} for element {elementId}: {ex.Message}");
            }
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
                DebugLogger.Log($"[CableTraySleevePlacer] Failed to get wall normal: {ex.Message}");
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
                    DebugLogger.Log($"[CableTraySleevePlacer] Wall is null, using intersection point as-is");
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
                    DebugLogger.Log($"[CableTraySleevePlacer] Projected to centerline: original={intersectionPoint}, centerline={finalPoint}, distance={distanceFromOriginal:F3}");

                    return finalPoint;
                }
                else
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] Wall location curve is not a line, using intersection point as-is");
                    return intersectionPoint;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CableTraySleevePlacer] Failed to project to wall centerline: {ex.Message}, using intersection point as-is");
                return intersectionPoint;
            }
        }
    }
}
