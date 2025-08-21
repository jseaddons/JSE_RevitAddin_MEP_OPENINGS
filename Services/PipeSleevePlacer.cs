
using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class PipeSleevePlacer
    {
        private readonly Document _doc;

        public PipeSleevePlacer(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }
    public bool PlaceSleeve(
            Pipe? pipe,
            XYZ intersection,
            XYZ pipeDirection,
            FamilySymbol sleeveSymbol,
            Element hostElement)
        {
            int pipeElementId = pipe != null ? (int)pipe.Id.IntegerValue : 0;
            try
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DIAGNOSTIC: Called for pipeId={pipeElementId}, hostType={hostElement?.GetType().Name}, direction=({pipeDirection.X:F3},{pipeDirection.Y:F3},{pipeDirection.Z:F3})");

                if (pipe == null || intersection == null || sleeveSymbol == null || hostElement == null)
                {
                    DebugLogger.Error($"[PipeSleevePlacer] FAILURE: Null parameters provided for pipe {pipeElementId}");
                    return false;
                }

                double sleeveDepth = 0.0;
                XYZ n = XYZ.BasisX;
                XYZ placePoint = intersection;

                if (hostElement is Wall hostWall)
                {
                    // The incoming 'intersection' is already a point on the wall centerline (we project earlier).
                    // To center the sleeve through the wall thickness, use the intersection directly as the
                    // placement anchor and set sleeveDepth to the wall thickness. A later bounding-box
                    // recenter step will ensure the family geometry is centered on this point.
                    double wallThickness = hostWall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? hostWall.Width;
                    n = GetWallNormal(hostWall, intersection).Normalize();
                    placePoint = intersection; // center on wall centerline
                    sleeveDepth = wallThickness;
                }
                else if (hostElement is Floor floor)
                {
                    // STRUCTURAL FLOOR GUARD: follow duct behavior and only place in structural floors
                    Parameter structuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                    bool isStructural = structuralParam?.AsInteger() == 1;
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] STRUCTURAL CHECK: FloorId={floor.Id.IntegerValue}, IsStructural={isStructural}");
                    if (!isStructural)
                    {
                        DebugLogger.Warning($"[PipeSleevePlacer] FAILURE: Non-structural floor {floor.Id.IntegerValue} - skipped for pipe {pipeElementId}");
                        return false;
                    }

                    ElementId typeId = floor.GetTypeId();
                    Element? floorType = floor.Document.IsLinked ? floor.Document.GetElement(typeId) : _doc.GetElement(typeId);
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
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Floor thickness detected: {UnitUtils.ConvertFromInternalUnits(sleeveDepth, UnitTypeId.Millimeters):F1}mm");
                    }
                    else
                    {
                        sleeveDepth = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters);
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Floor thickness param not found, using fallback 500mm");
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
                        sleeveDepth = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters);

                    var loc = famInst.Location as LocationCurve;
                    n = loc != null && loc.Curve is Line line ? line.Direction.CrossProduct(XYZ.BasisZ).Normalize() : XYZ.BasisY;
                    placePoint = intersection; // no offset for framing
                }
                else
                {
                    DebugLogger.Warning($"[PipeSleevePlacer] FAILURE: Unsupported host type for pipe {pipeElementId}");
                    return false;
                }

                if (!sleeveSymbol.IsActive)
                    sleeveSymbol.Activate();

                // Get the pipe's actual level from the linked model, not the active document level
                Level? pipeLevel = null;
                if (pipe != null)
                {
                    // DEBUG: Check what parameters are available on the linked pipe
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: Checking parameters for linked pipe {pipeElementId}");
                    
                    // Try to get the pipe's reference level using HostLevelHelper
                    pipeLevel = HostLevelHelper.GetHostReferenceLevel(_doc, pipe);
                    if (pipeLevel != null)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: HostLevelHelper returned level: '{pipeLevel.Name}' (ID: {pipeLevel.Id.IntegerValue}) for pipe {pipeElementId}");
                    }
                    else
                    {
                        // DEBUG: Log what parameters are available on the linked pipe
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: HostLevelHelper returned NULL for pipe {pipeElementId}");
                        
                        // Check what level parameters are available on the linked pipe
                        var refLevelParam = pipe.LookupParameter("Reference Level");
                        var levelParam = pipe.LookupParameter("Level");
                        var hostLevelParam = pipe.LookupParameter("Host Level");
                        
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: Pipe {pipeElementId} parameters - Reference Level: {refLevelParam?.StorageType}, Level: {levelParam?.StorageType}, Host Level: {hostLevelParam?.StorageType}");
                        
                        if (refLevelParam != null)
                            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: Reference Level value: {refLevelParam.AsString()}");
                        if (levelParam != null)
                            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: Level value: {levelParam.AsString()}");
                        if (hostLevelParam != null)
                            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: Host Level value: {hostLevelParam.AsString()}");
                    }
                }
                
                // Fallback: if we can't get the pipe's level, use the closest level in active document
                if (pipeLevel == null)
                {
                    pipeLevel = new FilteredElementCollector(_doc)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .OrderBy(lvl => Math.Abs(lvl.Elevation - placePoint.Z))
                        .FirstOrDefault();
                    
                    if (pipeLevel != null)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Fallback: Using closest active document level: {pipeLevel.Name} for pipe {pipeElementId}");
                    }
                }
                
                if (pipeLevel == null)
                {
                    DebugLogger.Error($"[PipeSleevePlacer] FAILURE: No level found for placement for pipe {pipeElementId}");
                    return false;
                }

                double mmDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(placePoint), UnitTypeId.Millimeters);
                double depthMM = UnitUtils.ConvertFromInternalUnits(sleeveDepth, UnitTypeId.Millimeters);
                double allowedOffset = (depthMM * 0.5) + 5.0;
                bool isAtOrigin = (Math.Abs(placePoint.X) < 0.001 && Math.Abs(placePoint.Y) < 0.001 && Math.Abs(placePoint.Z) < 0.001);
                bool isTooFar = mmDistance > allowedOffset;
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: intersection=({intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}), placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}), mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                if (isAtOrigin || isTooFar)
                {
                    DebugLogger.Warning($"[PipeSleevePlacer] FAILURE: Invalid placePoint for pipe {pipeElementId}: isAtOrigin={isAtOrigin}, isTooFar={isTooFar}, mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                    return false;
                }

                // Compute diameter/clearance early so it is available for final logging
                double pipeDiameter = 0;
                double insulationThickness = 0;
                if (pipe != null)
                {
                    pipeDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsDouble() ?? 0;
                    insulationThickness = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS)?.AsDouble() ?? 0;
                }
                double clearance = insulationThickness > 0.0
                    ? UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters)
                    : UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                double totalDiameter = pipeDiameter + 2 * insulationThickness + 2 * clearance;

                // Create the instance and perform manipulations assuming the caller has an active Transaction.
                // Avoid starting a nested Transaction here; starting transactions inside an already active transaction
                // causes the "Starting a new transaction is not permitted" exception. The caller (command) must
                // create and manage the outer transaction and, if needed, install a failures preprocessor.
                FamilyInstance? instance = null;
                try
                {
                                         instance = _doc.Create.NewFamilyInstance(
                         placePoint,
                         sleeveSymbol,
                         pipeLevel,
                         StructuralType.NonStructural);

                    // Set parameters after creation
                    SetParameterSafely(instance, "Diameter", totalDiameter, pipeElementId);
                    SetParameterSafely(instance, "Depth", sleeveDepth, pipeElementId);

                    try { _doc.Regenerate(); } catch { }

                    // Alignment and rotation (same logic as before)
                    if (hostElement is FamilyInstance fi && fi.Category != null && fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                    {
                        LocationCurve? framingCurve = fi.Location as LocationCurve;
                        if (framingCurve?.Curve is Line framingLine)
                        {
                            XYZ framingDir = framingLine.Direction.Normalize();
                            XYZ framingNormal = new XYZ(-framingDir.Y, framingDir.X, 0).Normalize();
                            double angle = Math.Atan2(framingNormal.Y, framingNormal.X);
                            if (instance.Location is LocationPoint loc)
                            {
                                Line rotationAxis = Line.CreateBound(loc.Point, loc.Point.Add(XYZ.BasisZ));
                                ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, angle);
                            }
                        }
                    }
                    else if (hostElement is Wall hw)
                    {
                        LocationCurve? locationCurve = hw.Location as LocationCurve;
                        if (locationCurve?.Curve is Line wallLine)
                        {
                            XYZ wallDir = wallLine.Direction.Normalize();
                            XYZ wallNormal = new XYZ(-wallDir.Y, wallDir.X, 0).Normalize();
                            double angle = Math.Atan2(wallNormal.Y, wallNormal.X);
                            if (instance.Location is LocationPoint loc)
                            {
                                Line rotationAxis = Line.CreateBound(loc.Point, loc.Point.Add(XYZ.BasisZ));
                                ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, angle);
                            }
                        }
                    }
                    else if (hostElement is Floor)
                    {
                        if (instance.Location is LocationPoint loc)
                        {
                            XYZ pipeDirXY = new XYZ(pipeDirection.X, pipeDirection.Y, 0).Normalize();
                            double angle = Math.Atan2(pipeDirXY.Y, pipeDirXY.X);
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point.Add(XYZ.BasisZ));
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, angle);
                        }
                    }
                    else
                    {
                        if (instance.Location is LocationPoint loc)
                        {
                            XYZ pipeDirXY = new XYZ(pipeDirection.X, pipeDirection.Y, 0).Normalize();
                            double angle = Math.Atan2(pipeDirXY.Y, pipeDirXY.X);
                            Line rotationAxis = Line.CreateBound(loc.Point, loc.Point.Add(XYZ.BasisZ));
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, angle);
                        }
                    }

                    // Recentre instance so its bbox center matches placePoint
                    try
                    {
                        try { _doc.Regenerate(); } catch { }
                        var bb = instance.get_BoundingBox(null);
                        if (bb != null)
                        {
                            var instCenter = (bb.Min + bb.Max) * 0.5;
                            var translation = placePoint - instCenter;
                            double moveMM = UnitUtils.ConvertFromInternalUnits(translation.GetLength(), UnitTypeId.Millimeters);
                            if (moveMM > 0.1) // only move if significant (>0.1mm)
                            {
                                ElementTransformUtils.MoveElement(_doc, instance.Id, translation);
                                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Adjusted instance by {translation} ({moveMM:F1}mm) to center through host");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Warning: failed to recentre instance: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Error($"[PipeSleevePlacer] FAILURE: Exception during placement for pipe {pipeElementId}: {ex.Message}");
                    return false;
                }

                // Set both Level and Schedule Level parameters using the pipe's actual level
                if (pipeLevel != null)
                {
                    // Set the Level parameter (Revit's internal level reference)
                    Parameter levelParam = instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                    bool levelSet = false;
                    if (levelParam != null && !levelParam.IsReadOnly)
                    {
                        levelParam.Set(pipeLevel.Id);
                        levelSet = true;
                    }
                    else
                    {
                        var levelByName = instance.LookupParameter("Level");
                        if (levelByName != null && !levelByName.IsReadOnly)
                        {
                            levelByName.Set(pipeLevel.Id);
                            levelSet = true;
                        }
                    }
                    
                    if (levelSet)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Set Level parameter to {pipeLevel.Name} for pipe {pipeElementId}");
                    }
                    else
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] WARNING: Could not set Level parameter for pipe {pipeElementId}. No writable level parameter found.");
                    }
                    
                    // Set the Schedule Level parameter
                    Parameter schedLevelParam = instance.LookupParameter("Schedule Level");
                    if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                    {
                        schedLevelParam.Set(pipeLevel.Id);
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Set Schedule Level to {pipeLevel.Name} for pipe {pipeElementId}");
                    }
                    else
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] WARNING: Could not set Schedule Level parameter for pipe {pipeElementId}. No writable Schedule Level parameter found.");
                    }
                }
                else
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] WARNING: No level available to set parameters for pipe {pipeElementId}");
                }

                // Final validation and logging
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Info($"[PipeSleevePlacer] Created sleeve instance Id={instance.Id.IntegerValue}");
                // calculate final offset
                LocationPoint? locationPoint = instance.Location as LocationPoint;
                XYZ finalPosition = (locationPoint == null || locationPoint.Point == null) ? placePoint : locationPoint.Point;
                double finalDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(finalPosition), UnitTypeId.Millimeters);
                DebugLogger.Info($"[PipeSleevePlacer] SUCCESS: Placed sleeve for pipe {pipeElementId} -> sleeveId={(int)instance.Id.IntegerValue}, diameter={totalDiameter:F1} (internal), pos=({finalPosition.X:F3},{finalPosition.Y:F3},{finalPosition.Z:F3})");
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[PipeSleevePlacer] FAILURE: Exception during placement for pipe {pipeElementId}: {ex.Message}");
                return false;
            }
        }

        private void SetParameterSafely(FamilyInstance instance, string paramName, double value, int elementId)
        {
            try
            {
                Parameter? param = instance.LookupParameter(paramName);
                if (param != null && !param.IsReadOnly)
                {
                    if (value <= 0.0)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] WARNING: Invalid {paramName} value {value} for pipe {elementId} - skipping");
                        return;
                    }

                    double valueInMm = UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters);
                    if (valueInMm > 10000.0)
                    {
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] WARNING: Extremely large {paramName} value {valueInMm:F1}mm for pipe {elementId} - skipping");
                        return;
                    }

                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Setting {paramName} to {valueInMm:F1}mm (internal: {value:F6}) for pipe {elementId}");
                    param.Set(value);
                    double actualValueInMm = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters);
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Verified {paramName} set to {actualValueInMm:F1}mm for pipe {elementId}");
                }
                else
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Parameter {paramName} not found or read-only for pipe {elementId}");
                }
            }
            catch (Exception ex)
            {
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Failed to set {paramName} for pipe {elementId}: {ex.Message}");
            }
        }

        private XYZ GetWallNormal(Wall wall, XYZ point)
        {
            try
            {
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
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Failed to get wall normal: {ex.Message}");
            }
            return new XYZ(1, 0, 0);
        }
    }
}

