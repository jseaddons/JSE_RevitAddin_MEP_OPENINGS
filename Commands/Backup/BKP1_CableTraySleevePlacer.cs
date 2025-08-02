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

        public FamilyInstance PlaceCableTraySleeve(
            CableTray tray,
            XYZ intersection,
            double width,
            double height,
            XYZ direction,
            FamilySymbol sleeveSymbol,
            Element hostElement)
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
                // Visibility filter for linked elements: skip if not visible
                if (hostElement.Document.IsLinked)
                {
                    // Find the RevitLinkInstance in the active doc that matches this linked doc
                    var linkInstances = new FilteredElementCollector(_doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>();
                    var linkInstance = linkInstances.FirstOrDefault(li => li.GetLinkDocument() == hostElement.Document);
                    if (linkInstance != null && (_doc.ActiveView.GetCategoryHidden(linkInstance.Category.Id) || linkInstance.IsHidden(_doc.ActiveView)))
                    {
                        DebugLogger.Log($"[CableTraySleevePlacer] Linked host is not visible in the active view, skipping sleeve placement.");
                        return null;
                    }
                }

                // Host-specific logic for depth, normal, and placement point
                double sleeveDepth = 0.0;
                XYZ n = XYZ.BasisX;
                XYZ placePoint = intersection;
                if (hostElement is Wall wall)
                {
                    // Robust wall thickness retrieval (match duct logic)
                    double wallThickness = 0.0;
                    var widthParam = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                    if (widthParam != null && widthParam.StorageType == StorageType.Double)
                        wallThickness = widthParam.AsDouble();
                    if (wallThickness <= 0)
                        wallThickness = wall.Width;
                    if (wallThickness <= 0)
                        wallThickness = wall.WallType.Width;
                    if (wallThickness <= 0)
                        wallThickness = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters); // fallback
                    sleeveDepth = wallThickness;

                    // --- Duct logic mimic: robust centering and offset ---
                    LocationCurve locationCurve = wall.Location as LocationCurve;
                    bool forceCenter = false;
                    bool isFittingAtEnd = false;
                    if (tray != null && tray.Location is LocationCurve trayLocCurve)
                    {
                        // Check if intersection is near tray end
                        double distToStart = trayLocCurve.Curve.GetEndPoint(0).DistanceTo(intersection);
                        double distToEnd = trayLocCurve.Curve.GetEndPoint(1).DistanceTo(intersection);
                        double fittingThreshold = UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters); // 25mm
                        if (distToStart < fittingThreshold || distToEnd < fittingThreshold)
                        {
                            isFittingAtEnd = true;
                        }
                        // Robust ≥90% inside host rule
                        BoundingBoxXYZ hostBox = wall.get_BoundingBox(null);
                        if (hostBox != null)
                        {
                            XYZ hostMin = hostBox.Min;
                            XYZ hostMax = hostBox.Max;
                            Curve trayCrv = trayLocCurve.Curve;
                            XYZ pStart = trayCrv.GetEndPoint(0);
                            XYZ pEnd = trayCrv.GetEndPoint(1);
                            XYZ dir = (pEnd - pStart).Normalize();
                            double t0 = (hostMin - pStart).DotProduct(dir);
                            double t1 = (hostMax - pStart).DotProduct(dir);
                            double overlap = Math.Max(0, Math.Min(t1, trayCrv.Length) - Math.Max(t0, 0));
                            double ratio = overlap / trayCrv.Length;
                            forceCenter = ratio >= 0.90;
                            isFittingAtEnd &= !forceCenter;
                            DebugLogger.Log($"[CableTraySleevePlacer] Wall host: overlap={overlap:F3}, ratio={ratio:F3}, forceCenter={forceCenter}");
                        }
                    }
                    if (locationCurve?.Curve is Line wallLine)
                    {
                        XYZ wallDirection = wallLine.Direction.Normalize();
                        n = new XYZ(-wallDirection.Y, wallDirection.X, 0).Normalize();
                        if (forceCenter || !isFittingAtEnd)
                        {
                            // Centered in wall
                            double param = wallLine.Project(intersection).Parameter;
                            XYZ centerlinePoint = wallLine.Evaluate(param, false);
                            placePoint = centerlinePoint;
                            DebugLogger.Log($"[CableTraySleevePlacer] Wall host: Centered (forceCenter or not fitting-at-end). placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3})");
                            DebugLogger.Log($"[CableTraySleevePlacer] Wall centerline calculation: intersection=({intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}), param={param:F3}, centerlinePoint=({centerlinePoint.X:F3},{centerlinePoint.Y:F3},{centerlinePoint.Z:F3})");
                        }
                        else
                        {
                            // Fitting at end, project to wall axis and shift fully inside
                            double param = wallLine.Project(intersection).Parameter;
                            XYZ intersectionForPlacement = wallLine.Evaluate(param, false);
                            placePoint = intersectionForPlacement;
                            DebugLogger.Log($"[CableTraySleevePlacer] Wall host: Fitting at end, projecting to wall axis and shifting fully inside. placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3})");
                        }
                    }
                    else
                    {
                        n = GetWallNormal(wall, intersection).Normalize();
                    }
                }
                else if (hostElement is Floor floor)
                {
            // Only place sleeves in structural floors
            var isStructuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
            bool isStructural = isStructuralParam != null && isStructuralParam.AsInteger() == 1;
            if (!isStructural)
            {
                DebugLogger.Log($"[CableTraySleevePlacer] Floor host is not structural, skipping sleeve placement.");
                return null;
            }
            // Try to get the floor type from the linked document if available
            ElementId typeId = floor.GetTypeId();
            Element floorType = null;
            Document typeDoc = floor.Document;
            if (typeDoc.IsLinked)
            {
                // If the host is from a linked doc, get the type from the linked doc
                floorType = typeDoc.GetElement(typeId);
                DebugLogger.Log($"[CableTraySleevePlacer] Host floor is from linked document: {typeDoc.Title}");
            }
            else
            {
                // If not linked, get from active doc
                floorType = _doc.GetElement(typeId);
            }
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
                // Log all available type parameters for debugging
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
            DebugLogger.Log($"[CableTraySleevePlacer] Cable tray direction for floor sleeve: ({direction.X:F3}, {direction.Y:F3}, {direction.Z:F3})");
                }
                else if (hostElement is FamilyInstance famInst && famInst.Category != null && famInst.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                {
            var framingType = famInst.Symbol;
            var bParam = framingType.LookupParameter("b");
            if (bParam != null && bParam.StorageType == StorageType.Double)
                sleeveDepth = bParam.AsDouble();
            else
            {
                sleeveDepth = UnitUtils.ConvertToInternalUnits(500.0, UnitTypeId.Millimeters); // fallback
                DebugLogger.Log($"[CableTraySleevePlacer] Framing thickness parameter not found, using fallback 500mm");
            }
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

                // Validate placement point (relaxed: only skip if at origin)
                double mmDistance = UnitUtils.ConvertFromInternalUnits(intersection.DistanceTo(placePoint), UnitTypeId.Millimeters);
                double depthMM = UnitUtils.ConvertFromInternalUnits(sleeveDepth, UnitTypeId.Millimeters);
                double allowedOffset = (depthMM * 0.5) + 5.0;
                bool isAtOrigin = (Math.Abs(placePoint.X) < 0.001 && Math.Abs(placePoint.Y) < 0.001 && Math.Abs(placePoint.Z) < 0.001);
                DebugLogger.Log($"[CableTraySleevePlacer] DEBUG: intersection=({intersection.X:F3},{intersection.Y:F3},{intersection.Z:F3}), placePoint=({placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}), mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}");
                if (isAtOrigin)
                {
                    DebugLogger.Log($"[CableTraySleevePlacer] ERROR: placePoint invalid for cable tray {cableTrayId}: isAtOrigin={isAtOrigin}, mmDistance={mmDistance:F1}, allowedOffset={allowedOffset:F1}. Skipping placement.");
                    return null;
                }

                DebugLogger.Log($"[CableTraySleevePlacer] Intended centerline placement: placePoint=[{placePoint.X:F3},{placePoint.Y:F3},{placePoint.Z:F3}]");
                DebugLogger.Log($"[CableTraySleevePlacer] Placing sleeve: width={width}, height={height}, intersection=({intersection.X},{intersection.Y},{intersection.Z})");
                DebugLogger.Log($"[CableTraySleevePlacer] Cable tray direction: ({direction.X}, {direction.Y}, {direction.Z})");
                DebugLogger.Log($"[CableTraySleevePlacer] FIXED: Using placePoint instead of intersection for placement");
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
                else if (hostElement is FamilyInstance famInst2 && famInst2.Category != null && famInst2.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
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
                    || (hostElement is FamilyInstance fi && fi.Category != null && fi.Category.Id.Value == (int)BuiltInCategory.OST_Floors);
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

                // Only rotate for Y-axis trays/framing
                if (isFloorHost)
                {
                    try
                    {
                        LocationPoint loc = instance.Location as LocationPoint;
                        if (loc != null)
                        {
                            bool isVertical = Math.Abs(direction.Z) > 0.99;
                            DebugLogger.Log($"[CableTraySleevePlacer] [RISER-DEBUG] Floor intersection: trayId={cableTrayId}, direction=({direction.X:F3},{direction.Y:F3},{direction.Z:F3}), isVertical={isVertical}");
                            BoundingBoxXYZ bbox = tray.get_BoundingBox(null);
                            bool shouldRotate = false;
                            if (isVertical)
                            {
                                shouldRotate = JSE_RevitAddin_MEP_OPENINGS.Helpers.SleeveRiserOrientationHelper.ShouldRotateRiserSleeve(tray, loc.Point, width, height);
                            }
                            JSE_RevitAddin_MEP_OPENINGS.Helpers.SleeveRiserOrientationHelper.LogRiserDebugInfo(
                                "CableTray", (int)tray.Id.Value, bbox, width, height, loc.Point, hostElement, direction, shouldRotate);
                            if (isVertical && shouldRotate)
                            {
                                Line axis = Line.CreateBound(loc.Point, loc.Point + XYZ.BasisZ);
                                double angle = Math.PI / 2.0;
                                ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);
                                DebugLogger.Log("[CableTraySleevePlacer] Rotated vertical cable tray sleeve 90 degrees (riser logic).");
                            }
                            else if (isVertical)
                            {
                                DebugLogger.Log("[CableTraySleevePlacer] No rotation for vertical cable tray sleeve (riser logic).");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"[CableTraySleevePlacer] Error in riser/floor logic: {ex.Message}");
                    }
                }
                // Wall-normal based rotation for wall hosts - COMMENTED OUT FOR DEBUGGING
                // // ------------- MASTER FIX -------------
                
                // Wall w = hostElement as Wall;
                // if (w != null)
                // {
                //     LocationCurve lc = w.Location as LocationCurve;
                //     if (lc?.Curve is Line wallLn)
                //     {
                //         XYZ wallNormal = wallLn.Direction.CrossProduct(XYZ.BasisZ).Normalize();

                //         LocationPoint sleeveLoc = instance.Location as LocationPoint;
                //         if (sleeveLoc != null)
                //         {
                //             XYZ currentFacing = instance.FacingOrientation;
                //             double angle = currentFacing.AngleOnPlaneTo(wallNormal, XYZ.BasisZ);

                //             if (!double.IsNaN(angle) && Math.Abs(angle) > 0.001)
                //             {
                //                 Line axis = Line.CreateBound(sleeveLoc.Point, sleeveLoc.Point + XYZ.BasisZ);
                //                 ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);
                //                 DebugLogger.Log($"[CableTraySleevePlacer] Aligned to wall normal, rotated {angle * 180 / Math.PI:F1}°");
                //             }
                //         }
                //     }
                // }
// ------------- END MASTER FIX -------------
                // Framing-normal based rotation for framing hosts - COMMENTED OUT FOR DEBUGGING
                // if (hostElement is FamilyInstance framingInst && framingInst.Category != null && framingInst.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                // {
                //     LocationCurve lc = framingInst.Location as LocationCurve;
                //     if (lc?.Curve is Line framingLn)
                //     {
                //         XYZ framingNormal = framingLn.Direction.CrossProduct(XYZ.BasisZ).Normalize();

                //         LocationPoint sleeveLoc = instance.Location as LocationPoint;
                //         if (sleeveLoc != null)
                //         {
                //             XYZ currentFacing = instance.FacingOrientation;
                //             double angle = currentFacing.AngleOnPlaneTo(framingNormal, XYZ.BasisZ);

                //             if (!double.IsNaN(angle) && Math.Abs(angle) > 0.001)
                //             {
                //                 Line axis = Line.CreateBound(sleeveLoc.Point, sleeveLoc.Point + XYZ.BasisZ);
                //                 ElementTransformUtils.RotateElement(_doc, instance.Id, axis, angle);
                //                 DebugLogger.Log($"[CableTraySleevePlacer] Aligned to framing normal, rotated {angle * 180 / Math.PI:F1}°");
                //             }
                //         }
                //     }
                // }
// ------------- END MASTER FIX -------------


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
                DebugLogger.Log($"[CableTraySleevePlacer] CENTERLINE FIX: Now using placePoint instead of intersection for placement");
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
        /// Checks if a cable tray sleeve of the same MEP type is present at the intersection point (for wall/floor hosts).
        /// Only suppresses if an existing cable tray sleeve is found at the location (not duct or pipe sleeves).
        /// </summary>
        private bool IsCableTraySleeveAtIntersection(Document doc, XYZ intersection, double searchRadiusMM = 100.0)
        {
            double searchRadius = UnitUtils.ConvertToInternalUnits(searchRadiusMM, UnitTypeId.Millimeters);
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .WhereElementIsNotElementType();

            // Only consider cable tray sleeves (category OST_CableTraySleeves)
            // Revit does not have OST_CableTraySleeves; most sleeve families are in OST_GenericModel
            var cableTraySleeveCatId = new ElementId((long)BuiltInCategory.OST_GenericModel);
            bool foundSuppression = false;
            foreach (FamilyInstance fi in collector)
            {
                if (fi.Category == null || fi.Category.Id != cableTraySleeveCatId)
                    continue;

                // Stricter: Only match if family or symbol name contains 'cabletray', 'ct#', or 'sleeve' (case-insensitive)
                string famName = fi.Symbol?.Family?.Name?.ToLowerInvariant() ?? "";
                string symName = fi.Symbol?.Name?.ToLowerInvariant() ?? "";
                if (!(famName.Contains("cabletray") || famName.Contains("ct#") || famName.Contains("sleeve") ||
                      symName.Contains("cabletray") || symName.Contains("ct#") || symName.Contains("sleeve")))
                    continue;

                double dist = double.MaxValue;
                if (fi.Location is LocationPoint locPt)
                {
                    dist = locPt.Point.DistanceTo(intersection);
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
                    break;
                }
            }
            return foundSuppression;
        }
                // ...existing methods...
        
        // Place this method here, before the final closing brace of the class
                public static void LogAllCableTrayOpeningsOnWall(Document doc)
        {
            try
            {
                string logPath = @"C:\JSE_CSharp_Projects\JSE_RevitAddin_MEP_OPENINGS\JSE_RevitAddin_MEP_OPENINGS\Logs\CableTraySleeveAudit.log";
                using (var writer = new System.IO.StreamWriter(logPath, false))
                {
                    writer.WriteLine($"[CableTraySleevePlacer] Audit of all CableTrayOpeningOnWall family instances in model '{doc.Title}' at {DateTime.Now}");
                    var collector = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilyInstance))
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(fi =>
                            fi.Symbol != null &&
                            fi.Symbol.Family != null &&
                            fi.Symbol.Family.Name.IndexOf("CableTrayOpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0);
        
                    int count = 0;
                    foreach (var fi in collector)
                    {
                        string famName = fi.Symbol.Family.Name;
                        string symName = fi.Symbol.Name;
                        string id = fi.Id.IntegerValue.ToString();
                        string loc = "<no location>";
                        if (fi.Location is LocationPoint lp)
                            loc = $"({lp.Point.X:F3}, {lp.Point.Y:F3}, {lp.Point.Z:F3})";
                        double width = fi.LookupParameter("Width")?.AsDouble() ?? 0;
                        double height = fi.LookupParameter("Height")?.AsDouble() ?? 0;
                        double depth = fi.LookupParameter("Depth")?.AsDouble() ?? 0;
                        writer.WriteLine($"ID: {id}, Family: {famName}, Symbol: {symName}, Location: {loc}, Width: {UnitUtils.ConvertFromInternalUnits(width, UnitTypeId.Millimeters):F1}mm, Height: {UnitUtils.ConvertFromInternalUnits(height, UnitTypeId.Millimeters):F1}mm, Depth: {UnitUtils.ConvertFromInternalUnits(depth, UnitTypeId.Millimeters):F1}mm");
                        count++;
                    }
                    writer.WriteLine($"Total CableTrayOpeningOnWall instances found: {count}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("CableTraySleevePlacer.LogAllCableTrayOpeningsOnWall ERROR: " + ex.Message);
            }
        
        }     
               
    }
    
}
