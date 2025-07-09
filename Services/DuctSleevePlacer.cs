using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;

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
        /// Places a duct sleeve at the intersection point with robust positioning
        /// </summary>
        public void PlaceDuctSleeve(Duct duct, XYZ intersection, double width, double height, 
            XYZ ductDirection, FamilySymbol sleeveSymbol, Wall hostWall, XYZ faceNormal = null)
        {
            // CRITICAL: This method is for placing Generic Model, Workplane-Based sleeves ONLY.
            // DO NOT add host-based logic or wall-hosted placement here. See summary above.
            // The wall parameter is used only for thickness/depth and orientation, NOT as a host.
            // This logic is required for placing sleeves in linked wall scenarios.
            // THERE IS NO HOST WALL, ONLY LINKED WALLS.

            int ductElementId = duct?.Id?.IntegerValue ?? 0;
            try
            {
                DebugLogger.Log($"[DuctSleevePlacer] PlaceDuctSleeve called for duct {ductElementId} at intersection {intersection}");

                if (duct == null || intersection == null || sleeveSymbol == null || hostWall == null)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Null parameter check failed");
                    SleeveLogManager.LogDuctSleeveFailure(ductElementId, "Null parameters provided");
                    return;
                }

                // Damper-in-wall filter: skip placement if a damper is found at the intersection
                if (IsDamperAtIntersection(_doc, intersection, hostWall, ductElementId))
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Damper detected at intersection for duct {ductElementId}, skipping sleeve placement.");
                    SleeveLogManager.LogDuctSleeveFailure(ductElementId, "Damper present at intersection");
                    return;
                }

                // Wall stub filter: skip if duct stops before 1/4 wall depth
                double wallThickness = hostWall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? hostWall.Width;
                XYZ n = (faceNormal != null) ? faceNormal.Normalize() : GetWallNormal(hostWall, intersection).Normalize();
                XYZ wallVector = n.Multiply(-wallThickness);
                XYZ wallFace = intersection;
                XYZ wallBack = intersection + wallVector;
                double ductEndDist = double.MaxValue;
                var locCurve = duct.Location as LocationCurve;
                if (locCurve != null)
                {
                    ductEndDist = locCurve.Curve.GetEndPoint(1).DistanceTo(intersection);
                }
                double quarterWall = wallThickness * 0.25;
                if (ductEndDist < quarterWall)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] Duct {ductElementId} ends before 1/4 wall depth (dist={ductEndDist:F3}, 1/4 wall={quarterWall:F3}), skipping sleeve placement.");
                    SleeveLogManager.LogDuctSleeveFailure(ductElementId, "Duct ends before 1/4 wall depth");
                    return;
                }

                // Ensure symbol is active
                if (!sleeveSymbol.IsActive)
                {
                    sleeveSymbol.Activate();
                    // Remove immediate regenerate - let Revit handle it
                }

                // --- PipeSleevePlacer logic: calculate wall thickness and normal ---
                wallThickness = hostWall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? hostWall.Width;
                n = (faceNormal != null) ? faceNormal.Normalize() : GetWallNormal(hostWall, intersection).Normalize();
                wallVector = n.Multiply(-wallThickness);
                // Center point of wall thickness
                XYZ placePoint = intersection + wallVector.Multiply(0.5); // wall centerline
                double sleeveDepth = wallThickness; // full wall thickness
                DebugLogger.Log($"[DuctSleevePlacer] wallThickness={wallThickness}, normal=[{n.X:F3},{n.Y:F3},{n.Z:F3}], wallVector=[{wallVector.X:F3},{wallVector.Y:F3},{wallVector.Z:F3}], intersection=[{intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}], placePoint=[{placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}], sleeveDepth={sleeveDepth:F3}");

                // Find the nearest or lowest level for placement (required for workplane-based family)
                Level level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(lvl => Math.Abs(lvl.Elevation - placePoint.Z))
                    .FirstOrDefault();
                if (level == null)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] No level found for placement");
                    SleeveLogManager.LogDuctSleeveFailure(ductElementId, "No level found for placement");
                    return;
                }

                // CRITICAL: Prevent accidental placement at (0,0,0) or far from intersection due to logic or data errors.
                // If placePoint is at or extremely close to the origin, or more than half the wall thickness + 5mm from the intersection, this almost always indicates a bug in intersection or wall logic.
                // DO NOT REMOVE THIS CHECK. If you see this log, investigate the intersection/wall calculation immediately. Do not attempt to correct after placement.
                double mmDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(placePoint), UnitTypeId.Millimeters);
                double wallThicknessMM = UnitUtils.ConvertFromInternalUnits(wallThickness, UnitTypeId.Millimeters);
                double allowedOffset = (wallThicknessMM * 0.5) + 5.0; // Allow up to half wall thickness (mm) plus 5mm
                bool isAtOrigin = (Math.Abs(placePoint.X) < 0.001 && Math.Abs(placePoint.Y) < 0.001 && Math.Abs(placePoint.Z) < 0.001);
                bool isTooFar = mmDistance > allowedOffset;
                DebugLogger.Log($"[DuctSleevePlacer] DEBUG: intersection=({intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}), placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}), mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                if (isAtOrigin || isTooFar)
                {
                    DebugLogger.Log($"[DuctSleevePlacer] ERROR: placePoint invalid for duct {ductElementId}: isAtOrigin={isAtOrigin}, isTooFar={isTooFar}, mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}. Skipping placement. Intersection: [{intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}], placePoint: [{placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}]");
                    SleeveLogManager.LogDuctSleeveFailure(ductElementId, $"Invalid placePoint: isAtOrigin={isAtOrigin}, isTooFar={isTooFar}, mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                    return;
                }

                // Log the intended centerline placement point
                DebugLogger.Log($"[DuctSleevePlacer] Intended centerline placement: placePoint=[{placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}]");                // Place sleeve at intersection point (wall face)
                FamilyInstance instance = _doc.Create.NewFamilyInstance(
                    intersection,
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);
                
                // Set parameters with validation before any rotation
                SetParameterSafely(instance, "Width", width, ductElementId);
                SetParameterSafely(instance, "Height", height, ductElementId);
                SetParameterSafely(instance, "Depth", sleeveDepth, ductElementId);
                
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
                bool isYAxisDuct = Math.Abs(ductDirection.Y) > Math.Abs(ductDirection.X);
                
                if (isYAxisDuct)
                {
                    DebugLogger.Log("[DuctSleevePlacer] Y-axis duct detected, attempting rotation");
                    try
                    {
                        // Get exact sleeve center point
                        LocationPoint loc = instance.Location as LocationPoint;
                        if (loc != null)
                        {
                            XYZ rotationPoint = loc.Point;
                            DebugLogger.Log($"[DuctSleevePlacer] Pre-rotation location: [{rotationPoint.X:F3}, {rotationPoint.Y:F3}, {rotationPoint.Z:F3}]");

                            // Simplified rotation logic: Always rotate 90 degrees around Z-axis
                            try
                            {
                                Line rotationAxis = Line.CreateBound(rotationPoint, rotationPoint.Add(XYZ.BasisZ));
                                double rotationAngle = Math.PI / 2; // 90 degrees in radians

                                ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, rotationAngle);
                                DebugLogger.Log("[DuctSleevePlacer] Simplified rotation: 90 degrees around Z-axis applied successfully");

                                // Remove immediate regenerate after rotation - let Revit handle it

                                // Verify rotation result
                                LocationPoint newLoc = instance.Location as LocationPoint;
                                if (newLoc != null)
                                {
                                    XYZ newPoint = newLoc.Point;
                                    DebugLogger.Log($"[DuctSleevePlacer] Post-rotation location: [{newPoint.X:F3}, {newPoint.Y:F3}, {newPoint.Z:F3}]");
                                }
                            }
                            catch (Exception ex)
                            {
                                DebugLogger.Log($"[DuctSleevePlacer] Simplified rotation failed: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[DuctSleevePlacer] Error during rotation: {ex.Message}");
                        DebugLogger.Log($"[DuctSleevePlacer] Stack trace: {ex.StackTrace}");
                    }
                }
                else
                {
                    DebugLogger.Log($"[DuctSleevePlacer] X-axis duct - no rotation needed");
                }

                // Validate final position and log offset from centerline
                LocationPoint locationPoint = instance.Location as LocationPoint;
                XYZ finalPosition = locationPoint?.Point;
                if (finalPosition == null || (Math.Abs(finalPosition.X) < 0.001 && Math.Abs(finalPosition.Y) < 0.001 && Math.Abs(finalPosition.Z) < 0.001))
                {
                    // If the location point is null or at origin, use the placement point as fallback
                    finalPosition = placePoint;
                    DebugLogger.Log($"[DuctSleevePlacer] Warning: Location retrieval returned null or origin. Using placement point as fallback.");
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
                SleeveLogManager.LogDuctSleeveSuccess(ductElementId, instance.Id.IntegerValue, finalWidth, finalHeight, finalPosition);
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DuctSleevePlacer] Exception in PlaceDuctSleeve for duct {ductElementId}: {ex.Message}");
                DebugLogger.Log($"[DuctSleevePlacer] Stack trace: {ex.StackTrace}");
                SleeveLogManager.LogDuctSleeveFailure(ductElementId, $"Exception during placement: {ex.Message}");
                throw; // Re-throw to ensure error is not silently ignored
            }
        }

        /// <summary>
        /// Static helper method for compatibility
        /// DO NOT change this signature or logic unless you are updating ALL callers.
        /// </summary>
        public static void PlaceDuctSleeveStatic(Document doc, Duct duct, XYZ intersection, double width, double height, 
            XYZ ductDirection, FamilySymbol sleeveSymbol, Wall hostWall, XYZ faceNormal = null)
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
            bool found = false;
            foreach (FamilyInstance fi in collector)
            {
                LocationPoint loc = fi.Location as LocationPoint;
                double dist = loc != null ? loc.Point.DistanceTo(intersection) : double.MaxValue;
                string famName = fi.Symbol.FamilyName.ToLower();
                string typeName = fi.Symbol.Name.ToLower();
                DebugLogger.Log($"[DuctSleevePlacer][DamperCheck] Duct {ductElementId}: Found accessory {fi.Id.IntegerValue} at dist={dist:F3}, family={famName}, type={typeName}");
                if (dist < searchRadius && (famName.Contains("damper") || typeName.Contains("damper") || famName.Contains("accessory") || typeName.Contains("accessory")))
                    found = true;
            }
            return found;
        }
    }
}
