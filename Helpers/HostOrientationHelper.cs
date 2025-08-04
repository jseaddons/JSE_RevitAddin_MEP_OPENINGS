using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
public static class HostOrientationHelper
{
    // Sets HostOrientation parameter for a sleeve instance
    public static void SetHostOrientationParameter(FamilyInstance sleeve, Document doc)
    {
        DebugLogger.Log($"SetHostOrientationParameter: Starting for sleeve {sleeve.Id}");
        try
        {
            var hostOrientationParam = sleeve.LookupParameter("HostOrientation");
            if (hostOrientationParam == null)
            {
                DebugLogger.Log($"SetHostOrientationParameter: HostOrientation parameter not found on sleeve {sleeve.Id}");
                return;
            }
            if (hostOrientationParam.IsReadOnly)
            {
                DebugLogger.Log($"SetHostOrientationParameter: HostOrientation parameter is read-only on sleeve {sleeve.Id}");
                return;
            }
            string orientationToSet = "";
            if (sleeve.Host is Wall wallHost)
            {
                // Use wall normal to determine orientation (X or Y)
                var normal = wallHost.Orientation;
                double absX = Math.Abs(normal.X);
                double absY = Math.Abs(normal.Y);
                if (absX > absY)
                    orientationToSet = "X";
                else if (absY > absX)
                    orientationToSet = "Y";
                else
                    orientationToSet = "Unknown";
                DebugLogger.Log($"SetHostOrientationParameter: Wall host detected, orientation={orientationToSet}");
            }
            else if (sleeve.Host is Floor)
            {
                orientationToSet = "FloorHosted";
                DebugLogger.Log($"SetHostOrientationParameter: Floor host detected, orientation=FloorHosted");
            }
            else if (sleeve.Host != null && sleeve.Host.Category != null && sleeve.Host.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
            {
                // Use framing direction to determine orientation (X or Y), just like wall
                var framing = sleeve.Host as FamilyInstance;
                if (framing != null)
                {
                    var locationCurve = framing.Location as LocationCurve;
                    if (locationCurve != null)
                    {
                        var curve = locationCurve.Curve as Line;
                        if (curve != null)
                        {
                            var direction = curve.Direction;
                            double absX = Math.Abs(direction.X);
                            double absY = Math.Abs(direction.Y);
                            if (absX > absY)
                                orientationToSet = "X";
                            else if (absY > absX)
                                orientationToSet = "Y";
                            else
                                orientationToSet = "Unknown";
                            DebugLogger.Log($"SetHostOrientationParameter: Framing host detected, orientation={orientationToSet}");
                        }
                        else
                        {
                            orientationToSet = "Unknown";
                            DebugLogger.Log($"SetHostOrientationParameter: Framing host detected, but no valid Line");
                        }
                    }
                    else
                    {
                        orientationToSet = "Unknown";
                        DebugLogger.Log($"SetHostOrientationParameter: Framing host detected, but no valid LocationCurve");
                    }
                }
                else
                {
                    orientationToSet = "Unknown";
                    DebugLogger.Log($"SetHostOrientationParameter: Framing host detected, but FamilyInstance is null");
                }
            }
            else
            {
                orientationToSet = "Unknown";
                DebugLogger.Log($"SetHostOrientationParameter: Unknown host type for sleeve {sleeve.Id}");
            }
            DebugLogger.Log($"SetHostOrientationParameter: Setting HostOrientation to '{orientationToSet}' for sleeve {sleeve.Id}");
            hostOrientationParam.Set(orientationToSet);
            DebugLogger.Log($"SetHostOrientationParameter: Completed for sleeve {sleeve.Id} with value '{orientationToSet}'");
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"SetHostOrientationParameter: Exception for sleeve {sleeve.Id}: {ex.Message}\n{ex.StackTrace}");
        }
    }
        // Debug version with extensive logging
        public static (Element? hostElement, string hostType, string orientation) GetIntersectedHostTypeAndOrientation(FamilyInstance sleeve, Document doc)
        {
            DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Starting geometric intersection analysis for sleeve {sleeve.Id}");
            
            // 1. Get the sleeve's solid geometry with detailed logging
            Solid? sleeveSolid = null;
            try
            {
                var sleeveGeom = sleeve.get_Geometry(new Options { ComputeReferences = true });
                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: Got geometry, checking for solids...");
                foreach (GeometryObject geomObj in sleeveGeom)
                {
                    DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: Found geometry object of type {geomObj.GetType().Name}");
                    if (geomObj is Solid s && s.Volume > 0) 
                    { 
                        sleeveSolid = s; 
                        DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: Found direct solid with volume {s.Volume}");
                        break; 
                    }
                    if (geomObj is GeometryInstance gi)
                    {
                        DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: Found GeometryInstance, checking instance geometry...");
                        foreach (GeometryObject instObj in gi.GetInstanceGeometry())
                        {
                            DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: Instance geometry object type {instObj.GetType().Name}");
                            if (instObj is Solid s2 && s2.Volume > 0) 
                            { 
                                sleeveSolid = s2; 
                                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: Found instance solid with volume {s2.Volume}");
                                break; 
                            }
                        }
                        if (sleeveSolid != null) break;
                    }
                }
                if (sleeveSolid == null || sleeveSolid.Volume < 1e-6)
                {
                    DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: No valid solid geometry found (sleeveSolid is null or volume too small)");
                    return (null, "Unknown", ""); // Cannot proceed without sleeve geometry
                }
                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: Valid solid found with volume {sleeveSolid.Volume}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Exception extracting sleeve geometry for {sleeve.Id}: {ex.Message}\n{ex.StackTrace}");
                return (null, "Unknown", "");
            }

            // 2. Collect all structural elements with counts
            var elements = new List<(Element, Transform?)>();
            try
            {
                // Host model
                var hostElements = new FilteredElementCollector(doc)
                    .WherePasses(new ElementMulticategoryFilter(new[] {
                        BuiltInCategory.OST_Walls,
                        BuiltInCategory.OST_StructuralFraming,
                        BuiltInCategory.OST_Floors
                    }))
                    .WhereElementIsNotElementType()
                    .ToElements();
                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Found {hostElements.Count} host model structural elements");
                foreach (Element e in hostElements) elements.Add((e, null));
                // Linked models (only visible links)
                var linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>();
                int linkCount = 0;
                foreach (var linkInstance in linkInstances)
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc == null || doc.ActiveView.GetCategoryHidden(linkInstance.Category.Id))
                    {
                        DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Link instance is null or hidden in active view, skipping");
                        continue;
                    }
                    var linkTransform = linkInstance.GetTotalTransform();
                    var linkedElements = new FilteredElementCollector(linkDoc)
                        .WherePasses(new ElementMulticategoryFilter(new[] {
                            BuiltInCategory.OST_Walls,
                            BuiltInCategory.OST_StructuralFraming,
                            BuiltInCategory.OST_Floors
                        }))
                        .WhereElementIsNotElementType()
                        .ToElements();
                    DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Found {linkedElements.Count} elements in visible linked model");
                    foreach (Element e in linkedElements) elements.Add((e, linkTransform));
                    linkCount++;
                }
                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Total elements to check: {elements.Count} (from {linkCount} linked models)");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Exception collecting host elements: {ex.Message}\n{ex.StackTrace}");
                return (null, "Unknown", "");
            }

            // 3. Check intersection with each element
            int checkedCount = 0;
            try
            {
                foreach (var tuple in elements)
                {
                    Element structuralElement = tuple.Item1;
                    Transform? linkTransform = tuple.Item2;
                    checkedCount++;
                    Solid? hostSolid = null;
                    try
                    {
                        var options = new Options { ComputeReferences = true };
                        var geometry = structuralElement.get_Geometry(options);
                        foreach (GeometryObject geomObj in geometry)
                        {
                            if (geomObj is Solid s && s.Volume > 0) { hostSolid = s; break; }
                            else if (geomObj is GeometryInstance gi)
                            {
                                foreach (GeometryObject instObj in gi.GetInstanceGeometry())
                                {
                                    if (instObj is Solid s2 && s2.Volume > 0) { hostSolid = s2; break; }
                                }
                                if (hostSolid != null) break;
                            }
                        }
                        if (hostSolid == null || hostSolid.Volume < 1e-6)
                        {
                            DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Host element {structuralElement.Id} has no valid solid geometry");
                            continue;
                        }
                        if (linkTransform != null && !linkTransform.IsIdentity)
                        {
                            hostSolid = SolidUtils.CreateTransformed(hostSolid, linkTransform);
                        }
                        var intersectedSolid = BooleanOperationsUtils.ExecuteBooleanOperation(sleeveSolid, hostSolid, BooleanOperationsType.Intersect);
                        if (intersectedSolid != null && intersectedSolid.Volume > 1e-6)
                        {
                            var hostInfo = GetHostTypeAndOrientation(structuralElement);
                            DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: FOUND INTERSECTION with {structuralElement.GetType().Name} ID:{structuralElement.Id} - Type:{hostInfo.hostType}, Orientation:{hostInfo.orientation}");
                            return (structuralElement, hostInfo.hostType, hostInfo.orientation);
                        }
                    }
                    catch (Autodesk.Revit.Exceptions.ApplicationException ex)
                    {
                        DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: ApplicationException for host element {structuralElement.Id}: {ex.Message}");
                        continue;
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Exception for host element {structuralElement.Id}: {ex.Message}\n{ex.StackTrace}");
                        continue;
                    }
                }
                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Sleeve {sleeve.Id}: No intersection found after checking {checkedCount} elements");
                return (null, "Unknown", "");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"GetIntersectedHostTypeAndOrientation: Exception during intersection loop: {ex.Message}\n{ex.StackTrace}");
                return (null, "Unknown", "");
            }
        }

        // Alternative method using bounding box proximity if geometry fails
        public static (Element? hostElement, string hostType, string orientation) GetNearestHostByProximity(FamilyInstance sleeve, Document doc)
        {
            DebugLogger.Log($"Using proximity-based host detection for sleeve {sleeve.Id}");
            
            var sleeveLocation = (sleeve.Location as LocationPoint)?.Point ?? sleeve.GetTransform().Origin;
            DebugLogger.Log($"Sleeve location: {sleeveLocation.X:F2}, {sleeveLocation.Y:F2}, {sleeveLocation.Z:F2}");
            
            Element? nearestHost = null;
            double nearestDistance = double.MaxValue;
            
            // Collect potential hosts from host model
            var hostElements = new FilteredElementCollector(doc)
                .WherePasses(new ElementMulticategoryFilter(new[] {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_Floors
                }))
                .WhereElementIsNotElementType()
                .ToElements();
            foreach (Element element in hostElements)
            {
                try
                {
                    var bbox = element.get_BoundingBox(null);
                    if (bbox == null) continue;
                    var center = (bbox.Min + bbox.Max) / 2;
                    double distance = sleeveLocation.DistanceTo(center);
                    if (distance < nearestDistance)
                    {
                        nearestDistance = distance;
                        nearestHost = element;
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"Error getting bounding box for element {element.Id}: {ex.Message}");
                }
            }
            // Also check visible linked models
            var linkInstances = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>();
            foreach (var linkInstance in linkInstances)
            {
                var linkDoc = linkInstance.GetLinkDocument();
                if (linkDoc == null || doc.ActiveView.GetCategoryHidden(linkInstance.Category.Id)) continue;
                var linkTransform = linkInstance.GetTotalTransform();
                var linkedElements = new FilteredElementCollector(linkDoc)
                    .WherePasses(new ElementMulticategoryFilter(new[] {
                        BuiltInCategory.OST_Walls,
                        BuiltInCategory.OST_StructuralFraming,
                        BuiltInCategory.OST_Floors
                    }))
                    .WhereElementIsNotElementType()
                    .ToElements();
                foreach (Element element in linkedElements)
                {
                    try
                    {
                        var bbox = element.get_BoundingBox(null);
                        if (bbox == null) continue;
                        var center = linkTransform.OfPoint((bbox.Min + bbox.Max) / 2);
                        double distance = sleeveLocation.DistanceTo(center);
                        if (distance < nearestDistance)
                        {
                            nearestDistance = distance;
                            nearestHost = element;
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Log($"Error getting bounding box for linked element {element.Id}: {ex.Message}");
                    }
                }
            }
            
            if (nearestHost != null)
            {
                var hostInfo = GetHostTypeAndOrientation(nearestHost);
                double distanceMm = UnitUtils.ConvertFromInternalUnits(nearestDistance, UnitTypeId.Millimeters);
                DebugLogger.Log($"Sleeve {sleeve.Id}: Nearest host is {nearestHost.GetType().Name} ID:{nearestHost.Id} at {distanceMm:F0}mm - Type:{hostInfo.hostType}");
                return (nearestHost, hostInfo.hostType, hostInfo.orientation);
            }
            
            DebugLogger.Log($"Sleeve {sleeve.Id}: No nearby host found");
            return (null, "Unknown", "");
        }

        // Rest of the methods remain the same
        public static string GetWallOrientation(Wall wall)
        {
            var normal = wall.Orientation;
            if (System.Math.Abs(normal.X) > System.Math.Abs(normal.Y))
                return "X";
            else
                return "Y";
        }

        public static string GetFramingOrientation(FamilyInstance framing)
        {
            var locationCurve = framing.Location as LocationCurve;
            if (locationCurve != null)
            {
                var curve = locationCurve.Curve as Line;
                if (curve != null)
                {
                    var direction = curve.Direction;
                    string orientation = System.Math.Abs(direction.X) > System.Math.Abs(direction.Y) ? "X" : "Y";
                    DebugLogger.Log($"Framing {framing.Id}: Calculated orientation = {orientation} (Direction: X={direction.X:F3}, Y={direction.Y:F3}, Z={direction.Z:F3})");
                    return orientation;
                }
            }
            DebugLogger.Log($"Framing {framing.Id}: Unable to determine orientation (no valid LocationCurve or Line)");
            return "Unknown";
        }

        public static (string hostType, string orientation) GetHostTypeAndOrientation(Element host)
        {
            if (host is Wall wall)
                return ("Wall", GetWallOrientation(wall));
            if (host is FamilyInstance fi && host.Category != null && host.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
                return ("Framing", GetFramingOrientation(fi));
            if (host is Floor floor)
            {
                var structuralParam = floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                if (structuralParam != null && structuralParam.AsInteger() == 1)
                    return ("Floor", "");
                else
                    return ("Unknown", "");
            }
            return ("Unknown", "");
        }
    }
}