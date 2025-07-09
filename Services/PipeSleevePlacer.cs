// THERE IS NO HOST WALL, ONLY LINKED WALLS. Always place sleeves at the wall centerline, not the face.
// USING REVIT 2024 API
// LOGIC OVERVIEW:
// - This placer only deals with linked walls (not host walls).
// - The intersection point provided is always the wall centerline (computed in the command, not here).
// - No wall face or orientation adjustment is performed here; all placement is at the true wall centerline.
// - After placement, for Y-pipes only, a single 90-degree rotation about the Z axis is applied at the placement point (to match duct logic).
// - No duplicate or orientation-matching rotation is performed.
//
// See comments in PlaceSleeve for details.

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class PipeSleevePlacer
    {
        private readonly Document _doc;

        public PipeSleevePlacer(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// Places a round opening (pipe sleeve) at the specified location and aligns it with the pipe direction on the wall face.
        /// </summary>
        public void PlaceSleeve(
            XYZ intersection,
            double pipeDiameter,
            XYZ pipeDirection,
            FamilySymbol sleeveSymbol,
            Wall? linkedWall,
            bool shouldRotate, // NEW PARAMETER
            int pipeElementId = 0)
        {
            // --- LOGIC: Only linked walls are handled; intersection is always wall centerline ---
            // The intersection point is precomputed as the wall centerline in the calling command logic.
            // No host wall or wall face logic is used here.
            DebugLogger.Log($"[PipeSleevePlacer] ENTER: pipeElementId={pipeElementId}, intersection={intersection}, pipeDiameter={pipeDiameter}, pipeDirection={pipeDirection}, sleeveSymbol={sleeveSymbol?.Id.IntegerValue}, linkedWall={(linkedWall != null ? linkedWall.Id.IntegerValue.ToString() : "null")}, shouldRotate={shouldRotate}");
            Pipe intersectingPipe = null;
            if (pipeElementId == 0)
            {
                intersectingPipe = FindIntersectingPipe(intersection);
                if (intersectingPipe != null)
                {
                    pipeElementId = intersectingPipe.Id.IntegerValue;
                    DebugLogger.Log($"[PipeSleevePlacer] Determined pipeElementId {pipeElementId} from intersection");
                }
                else
                {
                    DebugLogger.Log($"[PipeSleevePlacer] WARNING: Could not find intersecting pipe for intersection {intersection}");
                }
            }
            SleeveLogManager.LogPipeWallIntersection(pipeElementId, intersection, pipeDiameter);
            if (sleeveSymbol == null)
            {
                DebugLogger.Log($"[PipeSleevePlacer] ERROR: sleeveSymbol is null for pipe {pipeElementId}");
                SleeveLogManager.LogPipeSleeveFailure(pipeElementId, "Sleeve symbol is null");
                return;
            }
            double tolerance = UnitUtils.ConvertToInternalUnits(1.0, UnitTypeId.Inches);
            var existingInstances = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol.Id == sleeveSymbol.Id)
                .ToList();
            bool exists = existingInstances.Any(fi => fi.GetTransform().Origin.DistanceTo(intersection) < tolerance);
            if (exists)
            {
                DebugLogger.Log($"[PipeSleevePlacer] Skipping pipe {pipeElementId}, existing sleeve detected within {tolerance} units");
                SleeveLogManager.LogPipeSleeveFailure(pipeElementId, "Sleeve already exists at this location");
                return;
            }
            // Remove duplicate variable declarations by reusing variables outside the if blocks
            Level level = null;
            FamilyInstance instance = null;
            double insulationThickness = 0.0;
            if (linkedWall == null)
            {
                DebugLogger.Log($"[PipeSleevePlacer] linkedWall is null for pipe {pipeElementId} -- placing sleeve exactly at intersection point with NO adjustment");
                level = new FilteredElementCollector(_doc)
                          .OfClass(typeof(Level))
                          .Cast<Level>()
                          .FirstOrDefault();
                if (level == null)
                {
                    DebugLogger.Log($"[PipeSleevePlacer] ERROR: No level found for placement for pipe {pipeElementId}");
                    SleeveLogManager.LogPipeSleeveFailure(pipeElementId, "No level found for placement");
                    return;
                }
                try
                {
                    instance = _doc.Create.NewFamilyInstance(
                        intersection,
                        sleeveSymbol,
                        level,
                        StructuralType.NonStructural);
                    DebugLogger.Log($"[PipeSleevePlacer] NewFamilyInstance created at {intersection} for pipe {pipeElementId}");
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[PipeSleevePlacer] EXCEPTION: Failed to create NewFamilyInstance for pipe {pipeElementId}: {ex.Message}");
                    SleeveLogManager.LogPipeSleeveFailure(pipeElementId, $"Exception during instance creation: {ex.Message}");
                    return;
                }
                if (instance == null)
                {
                    DebugLogger.Log($"[PipeSleevePlacer] ERROR: NewFamilyInstance returned null for pipe {pipeElementId}");
                    SleeveLogManager.LogPipeSleeveFailure(pipeElementId, "Failed to create instance");
                    return;
                }
                if (intersectingPipe == null)
                {
                    intersectingPipe = FindIntersectingPipe(intersection);
                    if (intersectingPipe == null)
                    {
                        DebugLogger.Log($"[PipeSleevePlacer] WARNING: No intersecting pipe found for parameter set for pipe {pipeElementId}");
                    }
                }
                if (intersectingPipe != null)
                {
                    var insulation = intersectingPipe.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS);
                    if (insulation != null)
                    {
                        insulationThickness = insulation.AsDouble();
                        DebugLogger.Log($"[PipeSleevePlacer] Insulation thickness: {insulationThickness}");
                    }
                }
                // Set parameters with debug
                try
                {
                    SetParameterSafely(instance, "Diameter", pipeDiameter + insulationThickness * 2, pipeElementId);
                    // Depth: use pipeDiameter as a fallback if wall thickness is unknown
                    SetParameterSafely(instance, "Depth", pipeDiameter, pipeElementId);
                    DebugLogger.Log($"[PipeSleevePlacer] Set parameters for pipe {pipeElementId}: Diameter={pipeDiameter + insulationThickness * 2}, Depth={pipeDiameter}");
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[PipeSleevePlacer] EXCEPTION: Failed to set parameters for pipe {pipeElementId}: {ex.Message}");
                }
                try
                {
                    _doc.Regenerate();
                    DebugLogger.Log($"[PipeSleevePlacer] Document regenerated after placement for pipe {pipeElementId}");
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[PipeSleevePlacer] EXCEPTION: Regenerate failed for pipe {pipeElementId}: {ex.Message}");
                }
                DebugLogger.Log($"[PipeSleevePlacer] SUCCESS: Sleeve placed for pipe {pipeElementId} at {intersection}");
                // --- ROTATION: Only for Y-pipes, only once, after placement ---
                // For Y-pipes (abs(pipeDirection.Y) > abs(pipeDirection.X)), rotate the sleeve 90 degrees about Z at the placement point.
                // This matches the proven duct logic. No duplicate or orientation-matching rotation is performed.
                try
                {
                    if (instance != null && pipeDirection != null)
                    {
                        bool isYPipe = Math.Abs(pipeDirection.Y) > Math.Abs(pipeDirection.X);
                        if (isYPipe)
                        {
                            DebugLogger.Log("[PipeSleevePlacer] Y-pipe detected, applying 90-degree rotation about Z axis");
                            var locPoint = instance.Location as LocationPoint;
                            if (locPoint != null)
                            {
                                Line rotAxis = Line.CreateBound(locPoint.Point, locPoint.Point + XYZ.BasisZ);
                                double angle = Math.PI / 2.0; // 90 degrees
                                ElementTransformUtils.RotateElement(_doc, instance.Id, rotAxis, angle);
                                DebugLogger.Log($"[PipeSleevePlacer] Rotated sleeve instance 90 degrees for pipe {pipeElementId}");
                            }
                        }
                        else
                        {
                            DebugLogger.Log("[PipeSleevePlacer] X-pipe detected, no rotation needed");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[PipeSleevePlacer] ERROR: Y-pipe rotation failed: {ex.Message}");
                }
                // --- END ROTATION ---
                // --- BEGIN: Pipe/Sleeve Orientation Table Logging ---
                try
                {
                    string pipeOrientationStr = pipeDirection != null ? $"({pipeDirection.X:F6}, {pipeDirection.Y:F6}, {pipeDirection.Z:F6})" : "null";
                    string sleeveOrientationStr = "null";
                    if (instance != null)
                    {
                        // Try FacingOrientation, else fallback to transform
                        try
                        {
                            var facingOrientation = (instance as FamilyInstance)?.FacingOrientation;
                            if (facingOrientation != null)
                            {
                                sleeveOrientationStr = $"({facingOrientation.X:F6}, {facingOrientation.Y:F6}, {facingOrientation.Z:F6})";
                            }
                            else
                            {
                                var basisX = instance.GetTransform().BasisX;
                                sleeveOrientationStr = $"({basisX.X:F6}, {basisX.Y:F6}, {basisX.Z:F6})";
                            }
                        }
                        catch { }
                    }
                    DebugLogger.Log($"[PipeSleevePlacer] PIPE/SLEEVE ORIENTATION TABLE:\nPipeId | SleeveId\n{pipeElementId} | {instance?.Id.IntegerValue}\nPipeOrientation | SleeveOrientation\n{pipeOrientationStr} | {sleeveOrientationStr}");
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[PipeSleevePlacer] ERROR logging orientation table: {ex.Message}");
                }
                // --- END: Pipe/Sleeve Orientation Table Logging ---
                return;
            }
            // --- SIMPLIFIED LOGIC: Always place at provided intersection (wall centerline), do not adjust for wall face or orientation ---
            // Only use linkedWall to get wall thickness for parameter setting
            level = new FilteredElementCollector(_doc)
                      .OfClass(typeof(Level))
                      .Cast<Level>()
                      .FirstOrDefault();
            if (level == null)
            {
                DebugLogger.Log($"[PipeSleevePlacer] ERROR: No level found for placement for pipe {pipeElementId}");
                SleeveLogManager.LogPipeSleeveFailure(pipeElementId, "No level found for placement");
                return;
            }
            try
            {
                instance = _doc.Create.NewFamilyInstance(
                    intersection, // Always use the provided intersection (centerline)
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);
                DebugLogger.Log($"[PipeSleevePlacer] NewFamilyInstance created at {intersection} for pipe {pipeElementId}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[PipeSleevePlacer] EXCEPTION: Failed to create NewFamilyInstance for pipe {pipeElementId}: {ex.Message}");
                SleeveLogManager.LogPipeSleeveFailure(pipeElementId, $"Exception during instance creation: {ex.Message}");
                return;
            }
            if (instance == null)
            {
                DebugLogger.Log($"[PipeSleevePlacer] ERROR: NewFamilyInstance returned null for pipe {pipeElementId}");
                SleeveLogManager.LogPipeSleeveFailure(pipeElementId, "Failed to create instance");
                return;
            }
            if (intersectingPipe == null)
            {
                intersectingPipe = FindIntersectingPipe(intersection);
                if (intersectingPipe == null)
                {
                    DebugLogger.Log($"[PipeSleevePlacer] WARNING: No intersecting pipe found for parameter set for pipe {pipeElementId}");
                }
            }
            if (intersectingPipe != null)
            {
                var insulation = intersectingPipe.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS);
                if (insulation != null)
                {
                    insulationThickness = insulation.AsDouble();
                    DebugLogger.Log($"[PipeSleevePlacer] Insulation thickness: {insulationThickness}");
                }
            }
            // Set parameters with debug
            try
            {
                SetParameterSafely(instance, "Diameter", pipeDiameter + insulationThickness * 2, pipeElementId);
                // Depth: use wall thickness if available, otherwise fallback to pipeDiameter
                double wallThickness = linkedWall != null ? linkedWall.Width : pipeDiameter;
                SetParameterSafely(instance, "Depth", wallThickness, pipeElementId);
                DebugLogger.Log($"[PipeSleevePlacer] Set parameters for pipe {pipeElementId}: Diameter={pipeDiameter + insulationThickness * 2}, Depth={wallThickness}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[PipeSleevePlacer] EXCEPTION: Failed to set parameters for pipe {pipeElementId}: {ex.Message}");
            }
            try
            {
                _doc.Regenerate();
                DebugLogger.Log($"[PipeSleevePlacer] Document regenerated after placement for pipe {pipeElementId}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[PipeSleevePlacer] EXCEPTION: Regenerate failed for pipe {pipeElementId}: {ex.Message}");
            }
            DebugLogger.Log($"[PipeSleevePlacer] SUCCESS: Sleeve placed for pipe {pipeElementId} at {intersection}");
            // --- ROTATION: Only for Y-pipes, only once, after placement ---
            // For Y-pipes (abs(pipeDirection.Y) > abs(pipeDirection.X)), rotate the sleeve 90 degrees about Z at the placement point.
            // This matches the proven duct logic. No duplicate or orientation-matching rotation is performed.
            try
            {
                if (instance != null && pipeDirection != null)
                {
                    bool isYPipe = Math.Abs(pipeDirection.Y) > Math.Abs(pipeDirection.X);
                    if (isYPipe)
                    {
                        DebugLogger.Log("[PipeSleevePlacer] Y-pipe detected, applying 90-degree rotation about Z axis");
                        var locPoint = instance.Location as LocationPoint;
                        if (locPoint != null)
                        {
                            Line rotAxis = Line.CreateBound(locPoint.Point, locPoint.Point + XYZ.BasisZ);
                            double angle = Math.PI / 2.0; // 90 degrees
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotAxis, angle);
                            DebugLogger.Log($"[PipeSleevePlacer] Rotated sleeve instance 90 degrees for pipe {pipeElementId}");
                        }
                    }
                    else
                    {
                        DebugLogger.Log("[PipeSleevePlacer] X-pipe detected, no rotation needed");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[PipeSleevePlacer] ERROR: Y-pipe rotation failed: {ex.Message}");
            }
            // --- END ROTATION ---
            // --- BEGIN: Pipe/Sleeve Orientation Table Logging ---
            try
            {
                string pipeOrientationStr = pipeDirection != null ? $"({pipeDirection.X:F6}, {pipeDirection.Y:F6}, {pipeDirection.Z:F6})" : "null";
                string sleeveOrientationStr = "null";
                if (instance != null)
                {
                    // Try FacingOrientation, else fallback to transform
                    try
                    {
                        var facingOrientation = (instance as FamilyInstance)?.FacingOrientation;
                        if (facingOrientation != null)
                        {
                            sleeveOrientationStr = $"({facingOrientation.X:F6}, {facingOrientation.Y:F6}, {facingOrientation.Z:F6})";
                        }
                        else
                        {
                            var basisX = instance.GetTransform().BasisX;
                            sleeveOrientationStr = $"({basisX.X:F6}, {basisX.Y:F6}, {basisX.Z:F6})";
                        }
                    }
                    catch { }
                }
                DebugLogger.Log($"[PipeSleevePlacer] PIPE/SLEEVE ORIENTATION TABLE:\nPipeId | SleeveId\n{pipeElementId} | {instance?.Id.IntegerValue}\nPipeOrientation | SleeveOrientation\n{pipeOrientationStr} | {sleeveOrientationStr}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[PipeSleevePlacer] ERROR logging orientation table: {ex.Message}");
            }
            // --- END: Pipe/Sleeve Orientation Table Logging ---
        }

        // Helper method to get wall face normal at a point
        private XYZ GetWallFaceNormal(Wall wall, XYZ point)
        {
            try
            {
                Options options = new Options { ComputeReferences = true };
                GeometryElement geomElem = wall.get_Geometry(options);
                Face closestFace = null;
                double minDistance = double.MaxValue;
                foreach (GeometryObject geomObj in geomElem)
                {
                    if (geomObj is Solid solid)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            XYZ closestPoint = face.Project(point).XYZPoint;
                            double distance = point.DistanceTo(closestPoint);
                            if (distance < minDistance)
                            {
                                minDistance = distance;
                                closestFace = face;
                            }
                        }
                    }
                }
                if (closestFace != null)
                {
                    UV uv = closestFace.Project(point).UVPoint;
                    return closestFace.ComputeNormal(uv);
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[PipeSleevePlacer] WARNING: GetWallFaceNormal failed: {ex.Message}");
            }
            return XYZ.BasisZ; // Default to Z-axis if calculation fails
        }
        
        /// <summary>
        /// Find the pipe that most closely intersects with the given point
        /// </summary>
        private Pipe FindIntersectingPipe(XYZ intersection)
        {
            Pipe result = null;
            double minDistance = double.MaxValue;
            
            var pipeCollector = new FilteredElementCollector(_doc)
                .OfClass(typeof(Pipe))
                .WhereElementIsNotElementType()
                .Cast<Pipe>();
            
            foreach (Pipe pipe in pipeCollector)
            {
                // Get location curve
                LocationCurve locCurve = pipe.Location as LocationCurve;
                if (locCurve == null)
                    continue;

                Line line = locCurve.Curve as Line;
                if (line == null)
                    continue;

                // Calculate the precise distance from the intersection to the pipe centerline
                double dist = line.Distance(intersection);

                // Only consider horizontal pipes (not vertical)
                XYZ direction = line.Direction.Normalize();
                bool isVertical = Math.Abs(direction.Z) > 0.9;
                if (isVertical)
                    continue;

                // Check if intersection is within pipe segment
                XYZ p0 = line.GetEndPoint(0);
                XYZ p1 = line.GetEndPoint(1);
                XYZ projection = line.Project(intersection).XYZPoint;
                double param = line.GetEndParameter(0) +
                    (line.GetEndParameter(1) - line.GetEndParameter(0)) *
                    p0.DistanceTo(projection) / p0.DistanceTo(p1);

                bool isWithinSegment = param >= line.GetEndParameter(0) && param <= line.GetEndParameter(1);

                // If pipe is horizontal, close enough to intersection, and intersection is within pipe segment
                double tolerance = UnitUtils.ConvertToInternalUnits(5.0, UnitTypeId.Millimeters); // 5mm tolerance
                if (!isVertical && isWithinSegment && dist < minDistance && dist < tolerance)
                {
                    result = pipe;
                    minDistance = dist;
                }
            }

            return result;
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
                        DebugLogger.Log($"[PipeSleevePlacer] WARNING: Invalid {paramName} value {value} for element {elementId} - skipping");
                        return;
                    }
                    double valueInMm = UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters);
                    if (valueInMm > 10000.0) // Sanity check
                    {
                        DebugLogger.Log($"[PipeSleevePlacer] WARNING: Extremely large {paramName} value {valueInMm:F1}mm for element {elementId} - skipping");
                        return;
                    }
                    param.Set(value);
                    DebugLogger.Log($"[PipeSleevePlacer] Set {paramName} to {valueInMm:F1}mm for element {elementId}");
                }
                else
                {
                    DebugLogger.Log($"[PipeSleevePlacer] Parameter {paramName} not found or read-only for element {elementId}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[PipeSleevePlacer] Failed to set {paramName} for element {elementId}: {ex.Message}");
            }
        }

        // Helper: Check if a point is within a bounding box
        private static bool IsPointInBounds(XYZ pt, XYZ min, XYZ max)
        {
            return pt.X >= min.X && pt.X <= max.X &&
                   pt.Y >= min.Y && pt.Y <= max.Y &&
                   pt.Z >= min.Z && pt.Z <= max.Z;
        }
    }
}
