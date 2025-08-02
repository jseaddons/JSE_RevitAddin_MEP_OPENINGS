
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

        public void PlaceSleeve(
            Pipe pipe,
            XYZ intersection,
            XYZ pipeDirection,
            FamilySymbol sleeveSymbol,
            Element hostElement)
        {
            int pipeElementId = pipe != null ? (int)pipe.Id.Value : 0;
            try
            {
                if (pipe == null || intersection == null || sleeveSymbol == null || hostElement == null)
                {
                    return;
                }

                double sleeveDepth = 0.0;
                XYZ placePoint = intersection;

                // Always use the pipe's "middle elevation" parameter for Z when host is Wall or Framing
                // Use the 'Middle Elevation' parameter for Z placement (robust for all Revit versions)
                // DuctSleevePlacer-mimic: always use intersection for floors/framing, offset to wall centerline for walls, no Middle Elevation logic
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Pipe Id: {pipe?.Id?.Value}, Intersection: ({intersection.X:F3}, {intersection.Y:F3}, {intersection.Z:F3})");

                if (hostElement is Wall wall)
                {
                    double wallThickness = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)?.AsDouble() ?? wall.Width;
                    sleeveDepth = wallThickness;
                    LocationCurve? locationCurve = wall.Location as LocationCurve;
                    if (locationCurve?.Curve is Line wallLine)
                    {
                        // Project intersection onto wall centerline (XY), but use intersection's Z (duct logic mimic)
                        double param = wallLine.Project(intersection).Parameter;
                        XYZ centerlinePoint = wallLine.Evaluate(param, false);
                        placePoint = new XYZ(centerlinePoint.X, centerlinePoint.Y, intersection.Z);
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Wall host: Projected to centerline (duct logic mimic). placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}), wallThickness={wallThickness:F3}");
                    }
                }
                else if (hostElement is Floor floor)
                {
                    // Place sleeves in both architectural and structural floors (filter removed per user request)
                    // Get the floor type from the correct document (linked or host)
                    ElementId typeId = floor.GetTypeId();
                    Element? floorType = null;
                    Document typeDoc = floor.Document;
                    if (typeDoc.IsLinked)
                    {
                        floorType = typeDoc.GetElement(typeId);
                    }
                    else
                    {
                        floorType = _doc.GetElement(typeId);
                    }
                    Parameter? thicknessParam = null;
                    if (floorType != null)
                    {
                        // Only use 'Default Thickness' for slab depth, as in DuctSleevePlacer
                        thicknessParam = floorType.LookupParameter("Default Thickness");
                    }
                    if (thicknessParam != null && thicknessParam.StorageType == StorageType.Double)
                    {
                        sleeveDepth = thicknessParam.AsDouble();
                    }
                    else
                    {
                        // Log all available type parameters for debugging
                        if (floorType != null)
                        {
                            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Floor thickness parameter not found. Listing all type parameters:");
                            foreach (Parameter p in floorType.Parameters)
                            {
                                string val = p.HasValue ? p.AsValueString() : "<no value>";
                                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"  - {p.Definition.Name}: {val} (StorageType={p.StorageType})");
                            }
                        }
                        sleeveDepth = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters); // fallback
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Floor thickness parameter not found, using fallback 500mm");
                    }
                    // For floors, use intersection directly (no offset)
                    placePoint = intersection;
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Floor host: Using intersection. placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3})");
                }
                else if (hostElement is FamilyInstance innerfamInst && innerfamInst.Category != null && innerfamInst.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    var framingType = innerfamInst.Symbol;
                    var bParam = framingType.LookupParameter("b");
                    double framingWidth = 0.0;
                    if (bParam != null && bParam.StorageType == StorageType.Double)
                        framingWidth = bParam.AsDouble();
                    else
                        framingWidth = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters); // fallback
                    sleeveDepth = framingWidth;

                    XYZ intersectionForPlacement = intersection;
                    bool isFittingAtEnd = false;
                    bool forceCenter = false;
                    if (pipe != null && pipe.Location is LocationCurve pipeCurve)
                    {
                        double distToStart = pipeCurve.Curve.GetEndPoint(0).DistanceTo(intersection);
                        double distToEnd = pipeCurve.Curve.GetEndPoint(1).DistanceTo(intersection);
                        double fittingThreshold = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters); // 25mm
                        if (distToStart < fittingThreshold || distToEnd < fittingThreshold)
                        {
                            isFittingAtEnd = true;
                        }
                        // --- Robust ≥90% inside host rule ---
                        BoundingBoxXYZ hostBox = hostElement.get_BoundingBox(null);
                        if (hostBox != null)
                        {
                            XYZ hostMin = hostBox.Min;
                            XYZ hostMax = hostBox.Max;
                            Curve pipeCrv = pipeCurve.Curve;
                            XYZ pStart = pipeCrv.GetEndPoint(0);
                            XYZ pEnd = pipeCrv.GetEndPoint(1);
                            XYZ dir = (pEnd - pStart).Normalize();
                            double t0 = (hostMin - pStart).DotProduct(dir);
                            double t1 = (hostMax - pStart).DotProduct(dir);
                            double overlap = Math.Max(0, Math.Min(t1, pipeCrv.Length) - Math.Max(t0, 0));
                            double ratio = overlap / pipeCrv.Length;
                            forceCenter = ratio >= 0.90;
                            isFittingAtEnd &= !forceCenter;
                            JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Framing host: overlap={overlap:F3}, ratio={ratio:F3}, forceCenter={forceCenter}");
                        }
                    }
                    LocationCurve? framingCurve = innerfamInst.Location as LocationCurve;
                    if (framingCurve?.Curve is Line framingLine)
                    {
                        XYZ framingDir = framingLine.Direction.Normalize();
                        XYZ framingNormal = new XYZ(-framingDir.Y, framingDir.X, 0).Normalize();
                        // Project intersection to framing axis for fitting-at-end, else use intersection
                        XYZ basePoint = (forceCenter || !isFittingAtEnd)
                            ? intersection
                            : framingLine.Evaluate(framingLine.Project(intersection).Parameter, false);
                        // Find closest host face (hostMin or hostMax) along the normal
                        BoundingBoxXYZ hostBox = innerfamInst.get_BoundingBox(null);
                        XYZ hostMin = hostBox.Min;
                        XYZ hostMax = hostBox.Max;
                        // Compute two face points on the framing axis
                        XYZ axisOrigin = framingLine.Origin;
                        XYZ axisDir = framingDir;
                        // Project hostMin and hostMax onto the normal plane
                        double minProj = (hostMin - basePoint).DotProduct(framingNormal);
                        double maxProj = (hostMax - basePoint).DotProduct(framingNormal);
                        // Choose the face closest to the intersection
                        double useProj = Math.Abs(minProj) < Math.Abs(maxProj) ? minProj : maxProj;
                        XYZ facePoint = basePoint + framingNormal.Multiply(useProj);
                        // Offset from the face into the host by +0.5 * framingWidth
                        placePoint = facePoint + framingNormal.Multiply(framingWidth * 0.5 * Math.Sign(-useProj));
                        JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] Framing host: Offset from closest face. placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3})");
                    }
                }
                // Robust validation: prevent placement at origin or far from intersection (mimic DuctSleevePlacer)
                double mmDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(placePoint), UnitTypeId.Millimeters);
                double depthMM = UnitUtils.ConvertFromInternalUnits(sleeveDepth, UnitTypeId.Millimeters);
                double allowedOffset = (depthMM * 0.5) + 5.0;
                bool isAtOrigin = (Math.Abs(placePoint.X) < 0.001 && Math.Abs(placePoint.Y) < 0.001 && Math.Abs(placePoint.Z) < 0.001);
                bool isTooFar = mmDistance > allowedOffset;
                JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] DEBUG: intersection=({intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}), placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}), mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                if (isAtOrigin || isTooFar)
                {
                    JSE_RevitAddin_MEP_OPENINGS.Services.DebugLogger.Log($"[PipeSleevePlacer] ERROR: placePoint invalid: isAtOrigin={isAtOrigin}, isTooFar={isTooFar}, mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}. Skipping placement. Intersection: [{intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}], placePoint: [{placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}]");
                    return;
                }

                if (!sleeveSymbol.IsActive)
                    sleeveSymbol.Activate();

                Level? level = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(lvl => Math.Abs(lvl.Elevation - placePoint.Z))
                    .FirstOrDefault();
                if (level == null)
                {
                    return;
                }

                FamilyInstance instance = _doc.Create.NewFamilyInstance(
                    placePoint,
                    sleeveSymbol,
                    level,
                    StructuralType.NonStructural);


                double pipeDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)?.AsDouble() ?? 0;
                double insulationThickness = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INSULATION_THICKNESS)?.AsDouble() ?? 0;
                double clearance;
                if (insulationThickness > 0.0)
                {
                    // Insulated pipe: 25mm clearance
                    clearance = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters);
                }
                else
                {
                    // Non-insulated pipe: 50mm clearance
                    clearance = UnitUtils.ConvertToInternalUnits(50.0, UnitTypeId.Millimeters);
                }
                double totalDiameter = pipeDiameter + 2 * insulationThickness + 2 * clearance;



                // All host types: set Diameter (PipeOpeningOnSlab uses Diameter, not Width/Height)
                SetParameterSafely(instance, "Diameter", totalDiameter, pipeElementId);
                SetParameterSafely(instance, "Depth", sleeveDepth, pipeElementId);

                _doc.Regenerate();

                if (hostElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // Align sleeve with framing normal (perpendicular to framing axis)
                    LocationCurve? framingCurve = famInst.Location as LocationCurve;
                    if (framingCurve?.Curve is Line framingLine)
                    {
                        XYZ framingDir = framingLine.Direction.Normalize();
                        XYZ framingNormal = new XYZ(-framingDir.Y, framingDir.X, 0).Normalize();
                        double angle = Math.Atan2(framingNormal.Y, framingNormal.X);

                        if (instance.Location is LocationPoint loc)
                        {
                            XYZ rotationPoint = loc.Point;
                            Line rotationAxis = Line.CreateBound(rotationPoint, rotationPoint.Add(XYZ.BasisZ));
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, angle);
                        }
                    }
                }
                else if (hostElement is Wall hostWall)
                {
                    // Align sleeve with wall normal (perpendicular to wall axis)
                    LocationCurve? locationCurve = hostWall.Location as LocationCurve;
                    if (locationCurve?.Curve is Line wallLine)
                    {
                        XYZ wallDir = wallLine.Direction.Normalize();
                        XYZ wallNormal = new XYZ(-wallDir.Y, wallDir.X, 0).Normalize();
                        double angle = Math.Atan2(wallNormal.Y, wallNormal.X);
                        if (instance.Location is LocationPoint loc)
                        {
                            XYZ rotationPoint = loc.Point;
                            Line rotationAxis = Line.CreateBound(rotationPoint, rotationPoint.Add(XYZ.BasisZ));
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, angle);
                        }
                    }
                }
                else if (hostElement is Floor floor)
                {
                    // Align sleeve with floor normal
                    LocationCurve? locationCurve = floor.Location as LocationCurve;
                    if (locationCurve?.Curve is Line floorLine)
                    {
                        XYZ floorDir = floorLine.Direction.Normalize();
                        XYZ floorNormal = new XYZ(-floorDir.Y, floorDir.X, 0).Normalize();
                        double angle = Math.Atan2(floorNormal.Y, floorNormal.X);
                        if (instance.Location is LocationPoint loc)
                        {
                            XYZ rotationPoint = loc.Point;
                            Line rotationAxis = Line.CreateBound(rotationPoint, rotationPoint.Add(XYZ.BasisZ));
                            ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, angle);
                        }
                    }
                }
                else
                {
                    // Always align the sleeve X axis with the pipe direction in the XY plane
                    if (instance.Location is LocationPoint loc && loc != null)
                    {
                        XYZ pipeDirXY = new XYZ(pipeDirection.X, pipeDirection.Y, 0).Normalize();
                        double angle = Math.Atan2(pipeDirXY.Y, pipeDirXY.X); // Angle from X axis to pipe direction
                        XYZ rotationPoint = loc.Point;
                        Line rotationAxis = Line.CreateBound(rotationPoint, rotationPoint.Add(XYZ.BasisZ));
                        ElementTransformUtils.RotateElement(_doc, instance.Id, rotationAxis, angle);
                }
            }

                Level refLevel = HostLevelHelper.GetHostReferenceLevel(_doc, pipe);
                if (refLevel != null)
                {
                    Parameter schedLevelParam = instance.LookupParameter("Schedule Level");
                    if (schedLevelParam != null && !schedLevelParam.IsReadOnly)
                    {
                        schedLevelParam.Set(refLevel.Id);
                    }
                }
            }
            catch (Exception)
            {
                // Exception intentionally ignored for robustness
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
                        return;
                    }
                    param.Set(value);
                }
            }
            catch (Exception)
            {
                // Exception intentionally ignored for robustness
            }
        }
    }
}

