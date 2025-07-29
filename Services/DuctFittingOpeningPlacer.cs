using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class DuctFittingOpeningPlacer
    {
        private readonly Document _doc;

        public DuctFittingOpeningPlacer(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        public void PlaceDuctFittingOpening(
            FamilyInstance fitting,
            XYZ intersection,
            double width,
            double height,
            XYZ direction,
            FamilySymbol openingSymbol,
            Wall wall)
        {
            try
            {
                if (!openingSymbol.IsActive)
                {
                    openingSymbol.Activate();
                    DebugLogger.Log("Activated opening symbol: " + openingSymbol.Name);
                }

                double wallThickness = wall.Width;

                // Get wall normal for rotation
                XYZ wallNormal = GetWallNormal(wall);

                // *** USE PROJECTION METHOD LIKE WORKING CABLE TRAY PLACER ***
                XYZ placePoint = GetWallCenterlinePoint(wall, intersection);

                DebugLogger.Log($"DuctFittingPlacer: PROJECTION METHOD - original intersection={intersection}, projected centerline placePoint={placePoint}, wallThickness={wallThickness:F3}");

                Level level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(lvl => Math.Abs(lvl.Elevation - placePoint.Z))
                    .FirstOrDefault();

                if (level == null)
                {
                    throw new Exception("No level found near " + placePoint.Z);
                }

                DebugLogger.Log("Using level: " + level.Name);

                FamilyInstance instance = _doc.Create.NewFamilyInstance(
                    placePoint,
                    openingSymbol,
                    level,
                    StructuralType.NonStructural);

                if (instance == null)
                {
                    throw new Exception("Failed to create family instance at " + placePoint);
                }

                DebugLogger.Log("Created instance ID: " + instance.Id.IntegerValue);

                SetParameterSafely(instance, "Width", width);
                SetParameterSafely(instance, "Height", height);
                SetParameterSafely(instance, "Depth", wallThickness);

                RotateOpening(instance, direction, wallNormal, intersection);

                DebugLogger.Log("Opening placement successful for fitting " + fitting.Id);
            }
            catch (Exception ex)
            {
                DebugLogger.Log("Error placing opening for fitting " + fitting.Id + ": " + ex.Message);
                throw;
            }
        }

        private void RotateOpening(FamilyInstance instance, XYZ direction, XYZ wallNormal, XYZ intersection)
        {
            try
            {
                XYZ instanceXAxis = instance.GetTransform().BasisX;
                XYZ projectedDirection = direction - wallNormal.Multiply(direction.DotProduct(wallNormal));
                XYZ projectedXAxis = instanceXAxis - wallNormal.Multiply(instanceXAxis.DotProduct(wallNormal));

                if (projectedDirection.IsZeroLength() || projectedXAxis.IsZeroLength())
                {
                    DebugLogger.Log("Cannot rotate: zero-length projected vector");
                    return;
                }

                projectedDirection = projectedDirection.Normalize();
                projectedXAxis = projectedXAxis.Normalize();

                double cosAngle = projectedXAxis.DotProduct(projectedDirection);
                double crossZ = projectedXAxis.CrossProduct(projectedDirection).DotProduct(wallNormal);
                double angle = Math.Acos(cosAngle);

                if (crossZ < 0)
                    angle = -angle;

                Line axis = Line.CreateBound(intersection, intersection + wallNormal);
                ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);
                DebugLogger.Log("Rotated opening by " + (angle * 180.0 / Math.PI) + " degrees");
            }
            catch (Exception ex)
            {
                DebugLogger.Log("Rotation failed: " + ex.Message);
            }
        }

        private void SetParameterSafely(FamilyInstance instance, string paramName, double value)
        {
            try
            {
                Parameter param = instance.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(value);
                    DebugLogger.Log("Set parameter " + paramName + " to " + value);
                }
                else
                {
                    DebugLogger.Log("Parameter " + paramName + " not found or read-only. Cannot set value.");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log("Failed to set parameter " + paramName + ": " + ex.Message);
            }
        }

        private XYZ GetWallCenterlinePoint(Wall wall, XYZ intersectionPoint)
        {
            // Project intersection point onto wall centerline using wall's location curve
            // Same method as the working CableTraySleevePlacer
            try
            {
                if (wall == null)
                {
                    DebugLogger.Log($"[DuctFittingOpeningPlacer] Wall is null, using intersection point as-is");
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
                    DebugLogger.Log($"[DuctFittingOpeningPlacer] Projected to centerline: original={intersectionPoint}, centerline={finalPoint}, distance={distanceFromOriginal:F3}");

                    return finalPoint;
                }
                else
                {
                    DebugLogger.Log($"[DuctFittingOpeningPlacer] Wall location curve is not a line, using intersection point as-is");
                    return intersectionPoint;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DuctFittingOpeningPlacer] Failed to project to wall centerline: {ex.Message}, using intersection point as-is");
                return intersectionPoint;
            }
        }

        /// <summary>
        /// Calculate the normal vector of a wall
        /// </summary>
        private XYZ GetWallNormal(Wall wall)
        {
            try
            {
                LocationCurve locationCurve = wall.Location as LocationCurve;
                if (locationCurve?.Curve is Line wallLine)
                {
                    XYZ wallDirection = wallLine.Direction.Normalize();
                    // Cross product with Z-axis to get normal vector
                    XYZ normal = wallDirection.CrossProduct(XYZ.BasisZ).Normalize();
                    DebugLogger.Log($"[DuctFittingOpeningPlacer] Wall normal calculated: {normal}");
                    return normal;
                }
                else
                {
                    DebugLogger.Log("[DuctFittingOpeningPlacer] Wall location is not a line, using default normal");
                    return XYZ.BasisX; // Default normal
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[DuctFittingOpeningPlacer] Error calculating wall normal: {ex.Message}, using default");
                return XYZ.BasisX; // Default normal
            }
        }
    }
}
