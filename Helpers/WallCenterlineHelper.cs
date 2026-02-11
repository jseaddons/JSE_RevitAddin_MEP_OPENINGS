using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class WallCenterlineHelper
    {
        /// <summary>
        /// ✅ OPTIMIZED: Get invariant wall data (orientation, center coordinate, width) for caching.
        /// Extracts the geometric center (Y for X-walls, X for Y-walls) from the bounding box.
        /// Also returns wall width.
        /// Use this when processing multiple sleeves on the same wall to avoid repeated Bbox/Location/Parameter calls.
        /// ✅ LINKED DOC SUPPORT: Handles transforms if wall is in a linked model.
        /// </summary>
        public static (bool IsXWall, double CenterCoordinate, double Width, bool Success) GetWallInvariantData(Wall wall, Document hostDocument = null)
        {
            if (wall == null) return (false, 0, 0, false);
            
            try
            {
                // Get Width
                double width = wall.Width;

                BoundingBoxXYZ wallBbox = wall.get_BoundingBox(null);
                if (wallBbox == null) return (false, 0, width, false);
                
                // Calculate bbox center (Local Coordinate System)
                XYZ wallBboxCenter = new XYZ(
                    (wallBbox.Min.X + wallBbox.Max.X) / 2.0,
                    (wallBbox.Min.Y + wallBbox.Max.Y) / 2.0,
                    (wallBbox.Min.Z + wallBbox.Max.Z) / 2.0
                );
                
                // Handle Linked File Transform
                if (hostDocument != null && wall.Document != null && !wall.Document.Equals(hostDocument))
                {
                    // Find transform
                    var linkInstances = new FilteredElementCollector(hostDocument)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>();
                    
                    foreach (var linkInstance in linkInstances)
                    {
                        var linkDoc = linkInstance.GetLinkDocument();
                        if (linkDoc != null && linkDoc.Equals(wall.Document))
                        {
                            var transform = linkInstance.GetTotalTransform();
                            if (transform != null && !transform.IsIdentity)
                            {
                                wallBboxCenter = transform.OfPoint(wallBboxCenter);
                            }
                            break;
                        }
                    }
                }
                
                // Get orientation
                var locationCurve = wall.Location as LocationCurve;
                if (locationCurve == null || locationCurve.Curve == null)
                    return (false, 0, width, false);
                
                var curve = locationCurve.Curve;
                XYZ wallDirection;
                if (curve is Line line)
                {
                    wallDirection = line.Direction.Normalize();
                }
                else
                {
                    var start = curve.GetEndPoint(0);
                    var end = curve.GetEndPoint(1);
                    wallDirection = (end - start).Normalize();
                }
                
                double absX = Math.Abs(wallDirection.X);
                double absY = Math.Abs(wallDirection.Y);
                bool isXWall = absX > absY;
                
                // Return coordinate: Y for X-walls (perpendicular axis), X for Y-walls
                double centerCoord = isXWall ? wallBboxCenter.Y : wallBboxCenter.X;
                
                return (isXWall, centerCoord, width, true);
            }
            catch
            {
                return (false, 0, 0, false);
            }
        }

        /// <summary>
        /// ✅ OPTIMIZED: Get invariant structural framing data (Normal, Thickness) for caching.
        /// Calculates normal (with transform) and thickness ('b' or 'Width').
        /// </summary>
        public static (XYZ Normal, double Thickness, bool Success) GetFramingInvariantData(Element framing, Document hostDocument = null)
        {
            try
            {
                // Calculate Normal (handles linked transforms internally)
                // We'll reuse the logic from GetStructuralFramingCenterlinePoint via helper or duplicate it slightly efficiently?
                // DRY: Let's extract the Logic from GetStructuralFramingCenterlinePoint into this method?
                // Or just implement it here cleanly.
                
                // 1. Get Thickness
                double thickness = GetStructuralFramingThickness(framing);

                // 2. Get Normal (Local)
                XYZ normal = GetStructuralFramingNormal(framing);
                
                // 3. Transform Normal if needed
                if (hostDocument != null && framing.Document != null && !framing.Document.Equals(hostDocument))
                {
                     // Need transform. 
                     // Ideally we cache the transform too? But finding it is the expensive part.
                     // Copy-paste the transform finding logic for now or refactor to helper.
                     // For brevity/speed, I'll copy the robust loop.
                    var linkInstances = new FilteredElementCollector(hostDocument)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>();
                    
                    foreach (var linkInstance in linkInstances)
                    {
                        var linkDoc = linkInstance.GetLinkDocument();
                        if (linkDoc != null && linkDoc.Equals(framing.Document))
                        {
                            var transform = linkInstance.GetTotalTransform();
                            if (transform != null && !transform.IsIdentity)
                            {
                                normal = transform.OfVector(normal).Normalize();
                            }
                            break;
                        }
                    }
                }
                
                return (normal, thickness, true);
            }
            catch
            {
                return (XYZ.BasisZ, 0, false);
            }
        }

        /// <summary>
        /// ✅ OPTIMIZED: Get invariant floor data (Normal, Thickness) for caching.
        /// </summary>
        public static (XYZ Normal, double Thickness, bool Success) GetFloorInvariantData(Element floor)
        {
            try
            {
                double thickness = GetFloorThickness(floor);
                XYZ normal = GetFloorNormal(floor); // Usually 0,0,1
                return (normal, thickness, true);
            }
            catch
            {
                 return (XYZ.BasisZ, 0, false);
            }
        }

        /// <summary>
        /// ✅ SIMPLE METHOD: Get wall centerline point using bounding box center (no ray tracing/projection).
        /// Gets wall bounding box, calculates center point, and merges with intersection point based on orientation.
        /// This is more reliable than projection/ray tracing methods which can find wall face instead of centerline.
        /// </summary>
        public static XYZ GetWallCenterlinePointFromBbox(Wall wall, XYZ intersectionPoint, Document hostDocument = null)
        {
            // Use the efficient invariant data method
            // Note: This method re-calculates it every time. Cache-aware callers should use GetWallInvariantData directly.
            var data = GetWallInvariantData(wall, hostDocument);
            
            if (!data.Success)
            {
                if (OptimizationFlags.UseDiagnosticMode) DebugLogger.Log($"[CENTERLINE-DEBUG] Wall {wall?.Id} data extraction failed, returning input point");
                return intersectionPoint;
            }

            if (data.IsXWall)
            {
                // X-Wall: Replace Y with CenterCoordinate
                return new XYZ(intersectionPoint.X, data.CenterCoordinate, intersectionPoint.Z);
            }
            else
            {
                // Y-Wall (or others defaulting here): Replace X with CenterCoordinate
                return new XYZ(data.CenterCoordinate, intersectionPoint.Y, intersectionPoint.Z);
            }
        }

        
        // Returns the centerline point of the wall at a given intersection point, using robust exterior normal
        // ✅ CRITICAL FIX: Handles walls from linked documents by transforming wall normal to host coordinate system
        public static XYZ GetWallCenterlinePoint(Wall wall, XYZ intersectionPoint, Document hostDocument = null)
        {
            DebugLogger.Log($"[CENTERLINE-DEBUG] ===== WALL CENTERLINE CALCULATION =====");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Wall ID: {wall?.Id?.GetIntegerValue()}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Input intersectionPoint: {intersectionPoint}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Wall orientation: {wall?.Orientation}");
            
            if (wall == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Wall is null, returning input point");
                return intersectionPoint;
            }
            try
            {
                // ✅ CRITICAL FIX: Get wall normal and transform it if wall is from linked document
                // The intersection point is in host document coordinates, so wall normal must be in host coordinates too
                XYZ wallNormal = wall.Orientation.Normalize();
                
                // ✅ LINKED DOCUMENT FIX: If wall is from linked document, transform normal to host coordinate system
                Document wallDoc = wall.Document;
                Transform linkTransform = null;
                
                // Check if wall is from a linked document and get transform
                if (hostDocument != null && wallDoc != null && wallDoc != hostDocument)
                {
                    // Wall is from linked document, find the RevitLinkInstance
                    var linkInstances = new FilteredElementCollector(hostDocument)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>();
                    
                    foreach (var linkInstance in linkInstances)
                    {
                        var linkDoc = linkInstance.GetLinkDocument();
                        if (linkDoc != null && linkDoc.Equals(wallDoc))
                        {
                            linkTransform = linkInstance.GetTotalTransform();
                            DebugLogger.Log($"[CENTERLINE-DEBUG] Found linked wall, transforming normal from link to host coordinates");
                            break;
                        }
                    }
                    
                    // Transform wall normal to host coordinate system
                    if (linkTransform != null && !linkTransform.IsIdentity)
                    {
                        wallNormal = linkTransform.OfVector(wallNormal).Normalize();
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Transformed wall normal: {wallNormal}");
                    }
                }
                
                DebugLogger.Log($"[CENTERLINE-DEBUG] Wall normal (orientation): {wallNormal}");
                
                double wallWidth = wall.Width;
                double halfWidth = wallWidth / 2.0;
                DebugLogger.Log($"[CENTERLINE-DEBUG] Wall width: {UnitUtils.ConvertFromInternalUnits(wallWidth, UnitTypeId.Millimeters):F1}mm, halfWidth: {UnitUtils.ConvertFromInternalUnits(halfWidth, UnitTypeId.Millimeters):F1}mm");
                
                // ✅ CRITICAL FIX: Project intersection point onto wall centerline
                // Method: Project intersection point onto wall curve, then move to centerline
                // This works for both active and linked document walls
                
                try
                {
                    var locationCurve = wall.Location as LocationCurve;
                    if (locationCurve != null && locationCurve.Curve != null)
                    {
                        var curve = locationCurve.Curve;
                        
                        // ✅ STEP 1: Project intersection point onto wall curve (along wall length)
                        // This gives us a point on the wall centerline at the same position along the wall
                        double curveParam = curve.Project(intersectionPoint).Parameter;
                        XYZ pointOnCurve = curve.Evaluate(curveParam, true);
                        
                        // ✅ LINKED DOCUMENT FIX: Transform point on curve to host coordinates if needed
                        if (linkTransform != null && !linkTransform.IsIdentity)
                        {
                            pointOnCurve = linkTransform.OfPoint(pointOnCurve);
                            DebugLogger.Log($"[CENTERLINE-DEBUG] Transformed point on wall curve: {pointOnCurve}");
                        }
                        
                        // ✅ CRITICAL FIX: pointOnCurve is already on the wall centerline (LocationCurve IS the centerline)
                        // For dampers, we need to merge coordinates: keep X/Z from intersection point, use Y from centerline (for X-walls)
                        // or keep Y/Z from intersection point, use X from centerline (for Y-walls)
                        // This ensures the centerline point is at the correct position along the wall length
                        
                        // ✅ STEP 2: Use pointOnCurve directly since it's already on the wall centerline (LocationCurve)
                        // However, we need to preserve the intersection point's position along the wall length
                        // The projection onto the curve gives us the correct perpendicular coordinate (Y for X-walls, X for Y-walls)
                        // but we need to keep the intersection point's coordinate along the wall length
                        
                        // Get wall direction from the curve to determine orientation
                        XYZ wallDirection;
                        if (curve is Line line)
                        {
                            wallDirection = line.Direction.Normalize();
                        }
                        else
                        {
                            // For non-linear curves, use start-to-end direction
                            var start = curve.GetEndPoint(0);
                            var end = curve.GetEndPoint(1);
                            wallDirection = (end - start).Normalize();
                        }
                        
                        // Determine if wall is X-wall or Y-wall based on curve direction
                        double absX = Math.Abs(wallDirection.X);
                        double absY = Math.Abs(wallDirection.Y);
                        bool isXWall = absX > absY; // Wall runs primarily along X-axis
                        bool isYWall = absY > absX; // Wall runs primarily along Y-axis
                        
                        XYZ centerlinePoint;
                        if (isXWall)
                        {
                            // X-wall: Use centerline Y coordinate (perpendicular to wall), keep intersection point X and Z (along wall length and height)
                            centerlinePoint = new XYZ(intersectionPoint.X, pointOnCurve.Y, intersectionPoint.Z);
                            DebugLogger.Log($"[CENTERLINE-DEBUG] X-wall detected: Using centerline Y={pointOnCurve.Y}, keeping intersection X={intersectionPoint.X}, Z={intersectionPoint.Z}");
                        }
                        else if (isYWall)
                        {
                            // Y-wall: Use centerline X coordinate (perpendicular to wall), keep intersection point Y and Z (along wall length and height)
                            centerlinePoint = new XYZ(pointOnCurve.X, intersectionPoint.Y, intersectionPoint.Z);
                            DebugLogger.Log($"[CENTERLINE-DEBUG] Y-wall detected: Using centerline X={pointOnCurve.X}, keeping intersection Y={intersectionPoint.Y}, Z={intersectionPoint.Z}");
                        }
                        else
                        {
                            // Unknown orientation or slanted wall: Use pointOnCurve directly (it's already on centerline)
                            centerlinePoint = pointOnCurve;
                            DebugLogger.Log($"[CENTERLINE-DEBUG] Unknown wall orientation: Using pointOnCurve directly");
                        }
                        
                        // ✅ DIAGNOSTIC: Log the calculation
                        XYZ offset = centerlinePoint - intersectionPoint;
                        double offsetDistance = offset.GetLength();
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Point on wall curve (centerline): {pointOnCurve}");
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Final centerline point: {centerlinePoint}");
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Offset from input: {offset}, distance: {UnitUtils.ConvertFromInternalUnits(offsetDistance, UnitTypeId.Millimeters):F1}mm");
                        DebugLogger.Log($"[CENTERLINE-DEBUG] ===== END WALL CENTERLINE CALCULATION =====");
                        
                        return centerlinePoint;
                    }
                }
                catch (System.Exception ex)
                {
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Error projecting onto wall curve: {ex.Message}");
                }
                
                // ✅ FALLBACK: If we can't get centerline point, use half-width method (original approach)
                // This assumes intersection point is on one face and we move inward by half width
                DebugLogger.Log($"[CENTERLINE-DEBUG] Using fallback method: moving by half width");
                XYZ fallbackCenterlinePoint = intersectionPoint + wallNormal * (-halfWidth);
                DebugLogger.Log($"[CENTERLINE-DEBUG] Movement vector: {wallNormal * (-halfWidth)}");
                DebugLogger.Log($"[CENTERLINE-DEBUG] Final centerline point: {fallbackCenterlinePoint}");
                
                // Calculate and log the offset from original point
                XYZ fallbackOffset = fallbackCenterlinePoint - intersectionPoint;
                double fallbackOffsetDistance = fallbackOffset.GetLength();
                DebugLogger.Log($"[CENTERLINE-DEBUG] Offset from input: {fallbackOffset}, distance: {UnitUtils.ConvertFromInternalUnits(fallbackOffsetDistance, UnitTypeId.Millimeters):F1}mm");
                DebugLogger.Log($"[CENTERLINE-DEBUG] ===== END WALL CENTERLINE CALCULATION =====");
                
                return fallbackCenterlinePoint;
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Exception: {ex.Message}");
            }
            return intersectionPoint;
        }
        
        /// <summary>
        /// Returns the centerline point of a structural framing element at a given intersection point
        /// Follows the same logic pattern as GetWallCenterlinePoint but adapted for structural framing
        /// Only processes structural framing elements, following StructuralSleevePlacementCommand pattern
        /// ✅ CRITICAL FIX: Handles framing from linked documents by transforming framing normal to host coordinate system
        /// </summary>
        public static XYZ GetStructuralFramingCenterlinePoint(Element structuralFraming, XYZ intersectionPoint, Document hostDocument = null)
        {
            DebugLogger.Log($"[CENTERLINE-DEBUG] ===== STRUCTURAL FRAMING CENTERLINE CALCULATION =====");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Structural Framing ID: {structuralFraming?.Id?.GetIntegerValue()}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Input intersectionPoint: {intersectionPoint}");
            
            if (structuralFraming == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Structural framing is null, returning input point");
                return intersectionPoint;
            }
            
            // Check if this is a structural framing element
            if (structuralFraming.Category.Id.GetIntegerValue() != (int)BuiltInCategory.OST_StructuralFraming)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Element is not structural framing (Category: {structuralFraming.Category.Name}), returning input point");
                return intersectionPoint;
            }
            
            // STRUCTURAL TYPE FILTERING: Ensure this is truly a structural framing element (following StructuralSleevePlacementCommand pattern)
            var familyInstance = structuralFraming as FamilyInstance;
            if (familyInstance != null)
            {
                // Check if this has structural usage/type (beams, braces, columns, etc.)
                // Use the StructuralType property instead of parameter
                var structuralType = familyInstance.StructuralType;
                // StructuralType enum: NonStructural=0, Beam=1, Brace=2, Column=3, Footing=4, UnknownFraming=5
                if (structuralType == Autodesk.Revit.DB.Structure.StructuralType.NonStructural)
                {
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Framing {structuralFraming.Id.GetIntegerValue()} is not structural (StructuralType={structuralType}), returning input point");
                    return intersectionPoint;
                }
                DebugLogger.Log($"[CENTERLINE-DEBUG] Framing {structuralFraming.Id.GetIntegerValue()} is structural (StructuralType={structuralType}), proceeding with centerline calculation");
            }
            
            try
            {
                // ✅ CRITICAL FIX: Get framing normal and transform it if framing is from linked document
                // The intersection point is in host document coordinates, so framing normal must be in host coordinates too
                XYZ framingNormal = GetStructuralFramingNormal(structuralFraming);
                
                // ✅ LINKED DOCUMENT FIX: If framing is from linked document, transform normal to host coordinate system
                Document framingDoc = structuralFraming.Document;
                Transform linkTransform = null;
                
                // Check if framing is from a linked document and get transform
                if (hostDocument != null && framingDoc != null && framingDoc != hostDocument)
                {
                    // Framing is from linked document, find the RevitLinkInstance
                    var linkInstances = new FilteredElementCollector(hostDocument)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>();
                    
                    foreach (var linkInstance in linkInstances)
                    {
                        var linkDoc = linkInstance.GetLinkDocument();
                        if (linkDoc != null && linkDoc.Equals(framingDoc))
                        {
                            linkTransform = linkInstance.GetTotalTransform();
                            DebugLogger.Log($"[CENTERLINE-DEBUG] Found linked framing, transforming normal from link to host coordinates");
                            break;
                        }
                    }
                    
                    // Transform framing normal to host coordinate system
                    if (linkTransform != null && !linkTransform.IsIdentity)
                    {
                        framingNormal = linkTransform.OfVector(framingNormal).Normalize();
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Transformed framing normal: {framingNormal}");
                    }
                }
                
                DebugLogger.Log($"[CENTERLINE-DEBUG] Framing normal (orientation): {framingNormal}");
                
                // Get structural framing thickness following StructuralSleevePlacementCommand pattern
                double framingThickness = GetStructuralFramingThickness(structuralFraming);
                double halfThickness = framingThickness / 2.0;
                DebugLogger.Log($"[CENTERLINE-DEBUG] Framing thickness: {UnitUtils.ConvertFromInternalUnits(framingThickness, UnitTypeId.Millimeters):F1}mm, halfThickness: {UnitUtils.ConvertFromInternalUnits(halfThickness, UnitTypeId.Millimeters):F1}mm");
                
                // Move from intersection point toward framing centerline by half thickness
                XYZ centerlinePoint = intersectionPoint + framingNormal * (-halfThickness);
                DebugLogger.Log($"[CENTERLINE-DEBUG] Movement vector: {framingNormal * (-halfThickness)}");
                DebugLogger.Log($"[CENTERLINE-DEBUG] Final centerline point: {centerlinePoint}");
                
                // Calculate and log the offset from original point
                XYZ offset = centerlinePoint - intersectionPoint;
                double offsetDistance = offset.GetLength();
                DebugLogger.Log($"[CENTERLINE-DEBUG] Offset from input: {offset}, distance: {UnitUtils.ConvertFromInternalUnits(offsetDistance, UnitTypeId.Millimeters):F1}mm");
                DebugLogger.Log($"[CENTERLINE-DEBUG] ===== END STRUCTURAL FRAMING CENTERLINE CALCULATION =====");
                
                return centerlinePoint;
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Exception: {ex.Message}");
                return intersectionPoint;
            }
        }
        
        /// <summary>
        /// Returns the centerline point of a structural floor at a given intersection point
        /// Follows the same logic pattern as GetWallCenterlinePoint but adapted for floors
        /// Only processes structural floors, following StructuralSleevePlacementCommand pattern
        /// </summary>
        public static XYZ GetFloorCenterlinePoint(Element floor, XYZ intersectionPoint)
        {
            DebugLogger.Log($"[CENTERLINE-DEBUG] ===== FLOOR CENTERLINE CALCULATION =====");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Floor ID: {floor?.Id?.GetIntegerValue()}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Input intersectionPoint: {intersectionPoint}");
            
            if (floor == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Floor is null, returning input point");
                return intersectionPoint;
            }
            
            // Check if this is a floor element
            if (floor.Category.Id.GetIntegerValue() != (int)BuiltInCategory.OST_Floors)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Element is not a floor (Category: {floor.Category.Name}), returning input point");
                return intersectionPoint;
            }
            
            // STRUCTURAL TYPE FILTERING: Only process structural floors (following StructuralSleevePlacementCommand pattern)
            var floorElement = floor as Floor;
            if (floorElement != null)
            {
                // Check if this is a structural floor
                var structuralUsageParam = floorElement.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL);
                if (structuralUsageParam != null && structuralUsageParam.HasValue)
                {
                    bool isStructural = structuralUsageParam.AsInteger() == 1;
                    if (!isStructural)
                    {
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Floor {floor.Id.GetIntegerValue()} is not structural (isStructural={isStructural}), returning input point");
                        return intersectionPoint;
                    }
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Floor {floor.Id.GetIntegerValue()} is structural (isStructural={isStructural}), proceeding with centerline calculation");
                }
                else
                {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Floor {floor.Id.GetIntegerValue()} has no structural usage parameter, proceeding with centerline calculation");
                }
            }
            
            try
            {
                // Get floor thickness following StructuralSleevePlacementCommand pattern
                double floorThickness = GetFloorThickness(floor);
                double halfThickness = floorThickness / 2.0;
                DebugLogger.Log($"[CENTERLINE-DEBUG] Floor thickness: {UnitUtils.ConvertFromInternalUnits(floorThickness, UnitTypeId.Millimeters):F1}mm, halfThickness: {UnitUtils.ConvertFromInternalUnits(halfThickness, UnitTypeId.Millimeters):F1}mm");
                
                // For floors, the normal is typically vertical (Z-direction)
                // This follows the approach from StructuralSleevePlacementCommand for floor centerline calculation
                XYZ floorNormal = GetFloorNormal(floor);
                DebugLogger.Log($"[CENTERLINE-DEBUG] Floor normal: {floorNormal}");
                
                // Move from intersection point toward floor centerline by half thickness
                XYZ centerlinePoint = intersectionPoint + floorNormal * (-halfThickness);
                DebugLogger.Log($"[CENTERLINE-DEBUG] Movement vector: {floorNormal * (-halfThickness)}");
                DebugLogger.Log($"[CENTERLINE-DEBUG] Final centerline point: {centerlinePoint}");
                
                // Calculate and log the offset from original point
                XYZ offset = centerlinePoint - intersectionPoint;
                double offsetDistance = offset.GetLength();
                DebugLogger.Log($"[CENTERLINE-DEBUG] Offset from input: {offset}, distance: {UnitUtils.ConvertFromInternalUnits(offsetDistance, UnitTypeId.Millimeters):F1}mm");
                DebugLogger.Log($"[CENTERLINE-DEBUG] ===== END FLOOR CENTERLINE CALCULATION =====");
                
                return centerlinePoint;
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Exception: {ex.Message}");
                return intersectionPoint;
            }
        }
        
        /// <summary>
        /// Get structural framing thickness following StructuralSleevePlacementCommand pattern
        /// Prioritizes 'b' parameter, then falls back to 'Width'
        /// </summary>
        private static double GetStructuralFramingThickness(Element structuralFraming)
        {
            try
            {
                var familyInstance = structuralFraming as FamilyInstance;
                if (familyInstance != null)
                {
                    // Try 'b' parameter first (per DUPLICATION_SUPPRESSION_README.md)
                    var bParam = familyInstance.LookupParameter("b");
                    if (bParam != null && bParam.HasValue)
                    {
                        double bValue = bParam.AsDouble();
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Found 'b' parameter: {UnitUtils.ConvertFromInternalUnits(bValue, UnitTypeId.Millimeters):F1}mm");
                        return bValue;
                    }
                    
                    // Fall back to 'Width' parameter
                    var widthParam = familyInstance.LookupParameter("Width");
                    if (widthParam != null && widthParam.HasValue)
                    {
                        double widthValue = widthParam.AsDouble();
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Found 'Width' parameter: {UnitUtils.ConvertFromInternalUnits(widthValue, UnitTypeId.Millimeters):F1}mm");
                        return widthValue;
                    }
                }
                
                // If no parameters found, throw error as per StructuralSleevePlacementCommand pattern
                throw new InvalidOperationException($"Cannot determine structural framing depth: 'b' or 'Width' parameter not set for element ID {structuralFraming.Id.GetIntegerValue()}");
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Error getting structural framing thickness: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Get floor thickness following StructuralSleevePlacementCommand pattern
        /// Uses 'Default Thickness' parameter
        /// </summary>
        private static double GetFloorThickness(Element floor)
        {
            try
            {
                var floorElement = floor as Floor;
                if (floorElement != null)
                {
                    // Get thickness from Default Thickness parameter (per DUPLICATION_SUPPRESSION_README.md)
                    var thicknessParam = floorElement.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM);
                    if (thicknessParam != null && thicknessParam.HasValue)
                    {
                        double thicknessValue = thicknessParam.AsDouble();
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Found floor 'Default Thickness' parameter: {UnitUtils.ConvertFromInternalUnits(thicknessValue, UnitTypeId.Millimeters):F1}mm");
                        return thicknessValue;
                    }
                }
                
                // If no thickness parameter found, use a default
                double defaultThickness = UnitUtils.ConvertToInternalUnits(200.0, UnitTypeId.Millimeters); // 200mm default
                DebugLogger.Log($"[CENTERLINE-DEBUG] No thickness parameter found, using default: {UnitUtils.ConvertFromInternalUnits(defaultThickness, UnitTypeId.Millimeters):F1}mm");
                return defaultThickness;
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Error getting floor thickness: {ex.Message}");
                double defaultThickness = UnitUtils.ConvertToInternalUnits(200.0, UnitTypeId.Millimeters);
                return defaultThickness;
            }
        }
        
        /// <summary>
        /// Get structural framing normal direction for centerline calculation
        /// Following StructuralSleevePlacementCommand approach for determining orientation
        /// </summary>
        private static XYZ GetStructuralFramingNormal(Element structuralFraming)
        {
            try
            {
                var familyInstance = structuralFraming as FamilyInstance;
                if (familyInstance != null)
                {
                    // Use the family instance's facing orientation as the normal
                    XYZ facingOrientation = familyInstance.FacingOrientation;
                    if (facingOrientation != null && !facingOrientation.IsZeroLength())
                    {
                        return facingOrientation.Normalize();
                    }
                    
                    // Fall back to hand orientation
                    XYZ handOrientation = familyInstance.HandOrientation;
                    if (handOrientation != null && !handOrientation.IsZeroLength())
                    {
                        return handOrientation.Normalize();
                    }
                }
                
                // Default to Y-axis if no orientation available
                DebugLogger.Log($"[CENTERLINE-DEBUG] No valid orientation found, using default Y-axis normal");
                return new XYZ(0, 1, 0);
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Error getting structural framing normal: {ex.Message}");
                return new XYZ(0, 1, 0);
            }
        }
        
        /// <summary>
        /// Get floor normal direction for centerline calculation
        /// For floors, this is typically the vertical (Z) direction
        /// </summary>
        private static XYZ GetFloorNormal(Element floor)
        {
            try
            {
                // For floors, the normal is typically vertical (Z-direction)
                // This is consistent with StructuralSleevePlacementCommand approach
                return new XYZ(0, 0, 1); // Upward Z-direction
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Error getting floor normal: {ex.Message}");
                return new XYZ(0, 0, 1);
            }
        }
        
        /// <summary>
        /// ✅ LEGACY RAY-TRACE METHOD: Get wall centerline point by finding 2 wall faces and calculating midpoint.
        /// This is the "half in and half out" method for ducts, pipes, and cable trays.
        /// Uses ReferenceIntersector to find wall faces in both directions, then calculates midpoint.
        /// </summary>
        public static XYZ GetWallCenterlinePointFromRayTrace(Wall wall, XYZ intersectionPoint, Document hostDocument = null)
        {
            DebugLogger.Log($"[CENTERLINE-DEBUG] ===== WALL CENTERLINE FROM RAY-TRACE (LEGACY METHOD) =====");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Wall ID: {wall?.Id?.GetIntegerValue()}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Input intersectionPoint: {intersectionPoint}");
            
            if (wall == null || hostDocument == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Wall or hostDocument is null, returning input point");
                return intersectionPoint;
            }
            
            try
            {
                // ✅ STEP 1: Get wall normal to determine ray direction
                XYZ wallNormal = wall.Orientation.Normalize();
                
                // ✅ LINKED DOCUMENT FIX: Transform normal if wall is from linked document
                Document wallDoc = wall.Document;
                Transform linkTransform = null;
                
                if (wallDoc != null && wallDoc != hostDocument)
                {
                    var linkInstances = new FilteredElementCollector(hostDocument)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>();
                    
                    foreach (var linkInstance in linkInstances)
                    {
                        var linkDoc = linkInstance.GetLinkDocument();
                        if (linkDoc != null && linkDoc.Equals(wallDoc))
                        {
                            linkTransform = linkInstance.GetTotalTransform();
                            wallNormal = linkTransform.OfVector(wallNormal).Normalize();
                            DebugLogger.Log($"[CENTERLINE-DEBUG] Transformed wall normal from link to host coordinates");
                            break;
                        }
                    }
                }
                
                // ✅ STEP 2: Create ReferenceIntersector to find wall faces
                var view3D = new FilteredElementCollector(hostDocument)
                    .OfClass(typeof(View3D))
                    .Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate);
                
                if (view3D == null)
                {
                    DebugLogger.Log($"[CENTERLINE-DEBUG] No 3D view found, using fallback method");
                    return GetWallCenterlinePoint(wall, intersectionPoint, hostDocument);
                }
                
                var refIntersector = new ReferenceIntersector(
                    new ElementClassFilter(typeof(Wall)),
                    FindReferenceTarget.Face,
                    view3D);
                refIntersector.FindReferencesInRevitLinks = true;
                
                // ✅ STEP 3: Cast rays in both directions to find wall faces
                var rayDir = wallNormal;
                var hitsFwd = refIntersector.Find(intersectionPoint, rayDir)?.Where(h => h != null).OrderBy(h => h.Proximity).Take(2).ToList();
                var hitsBack = refIntersector.Find(intersectionPoint, rayDir.Negate())?.Where(h => h != null).OrderBy(h => h.Proximity).Take(2).ToList();
                
                // ✅ STEP 4: Find the two closest wall faces (one in each direction)
                ReferenceWithContext? face1 = null;
                ReferenceWithContext? face2 = null;
                
                if (hitsFwd != null && hitsFwd.Count > 0)
                {
                    // Check if the forward hit is the same wall
                    var forwardHit = hitsFwd.FirstOrDefault(h => 
                    {
                        var refElem = hostDocument.GetElement(h.GetReference().ElementId);
                        return refElem != null && refElem.Id == wall.Id;
                    });
                    if (forwardHit != null)
                    {
                        face1 = forwardHit;
                    }
                }
                
                if (hitsBack != null && hitsBack.Count > 0)
                {
                    // Check if the backward hit is the same wall
                    var backwardHit = hitsBack.FirstOrDefault(h => 
                    {
                        var refElem = hostDocument.GetElement(h.GetReference().ElementId);
                        return refElem != null && refElem.Id == wall.Id;
                    });
                    if (backwardHit != null)
                    {
                        face2 = backwardHit;
                    }
                }
                
                // ✅ STEP 5: Calculate midpoint between the two faces (half in and half out)
                if (face1 != null && face2 != null)
                {
                    var point1 = intersectionPoint + rayDir * face1.Proximity;
                    var point2 = intersectionPoint - rayDir * face2.Proximity;
                    var midpoint = (point1 + point2) * 0.5;
                    
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Found 2 wall faces: Face1 distance={face1.Proximity * 304.8:F1}mm, Face2 distance={face2.Proximity * 304.8:F1}mm");
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Point1=({point1.X:F6}ft, {point1.Y:F6}ft, {point1.Z:F6}ft), Point2=({point2.X:F6}ft, {point2.Y:F6}ft, {point2.Z:F6}ft)");
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Midpoint (centerline)=({midpoint.X:F6}ft, {midpoint.Y:F6}ft, {midpoint.Z:F6}ft)");
                    DebugLogger.Log($"[CENTERLINE-DEBUG] ===== END WALL CENTERLINE FROM RAY-TRACE =====");
                    
                    // ✅ CRITICAL REFINEMENT: Ensure strictly only the wall-depth coordinate changes
                    // The sleeve must stay perfectly aligned with the duct intersection point (Length & Height axes)
                    // We only want to adjust the position 'in/out' of the wall (Depth axis)
                    
                    XYZ mergedPoint = midpoint;
                    
                    // Determine if wall is X-wall or Y-wall to lock correct axes
                    double absX = Math.Abs(wallNormal.X);
                    double absY = Math.Abs(wallNormal.Y);
                    bool isXWall = absX > absY; // Wall runs perpendicular to X (Normal is X-dominant) -> NO! Normal X means wall is Y-aligned?
                    // Wait, Normal X (1,0,0) means wall faces X. So wall runs along Y. That's a "Y-Wall".
                    // Vertical Wall: Normal is horizontal.
                    
                    // Let's stick to standard naming:
                    // If Normal is predominantly X (e.g. 1,0,0), the wall plane is YZ. This is a "North-South" wall. Sleeve moves along X.
                    // If Normal is predominantly Y (e.g. 0,1,0), the wall plane is XZ. This is an "East-West" wall. Sleeve moves along Y.
                    
                    if (absX > absY) 
                    {
                        // Normal is X-dominant (Wall plane is YZ). we adjust X. Keep Y and Z from intersection.
                        mergedPoint = new XYZ(midpoint.X, intersectionPoint.Y, intersectionPoint.Z);
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Wall Normal X-dominant ({wallNormal.X:F2}): Adjusting X only. Keeping Intersect Y,Z.");
                    }
                    else 
                    {
                        // Normal is Y-dominant (Wall plane is XZ). We adjust Y. Keep X and Z from intersection.
                        mergedPoint = new XYZ(intersectionPoint.X, midpoint.Y, intersectionPoint.Z);
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Wall Normal Y-dominant ({wallNormal.Y:F2}): Adjusting Y only. Keeping Intersect X,Z.");
                    }
                    
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Midpoint (raw): ({midpoint.X:F6}, {midpoint.Y:F6}, {midpoint.Z:F6})");
                    DebugLogger.Log($"[CENTERLINE-DEBUG] MergedPoint (aligned): ({mergedPoint.X:F6}, {mergedPoint.Y:F6}, {mergedPoint.Z:F6})");
                    
                    return mergedPoint;
                }
                else
                {
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Could not find 2 wall faces (face1={face1 != null}, face2={face2 != null}), using fallback method");
                    return GetWallCenterlinePoint(wall, intersectionPoint, hostDocument);
                }
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Exception in GetWallCenterlinePointFromRayTrace: {ex.Message}");
                return GetWallCenterlinePoint(wall, intersectionPoint, hostDocument);
            }
        }
        
        /// <summary>
        /// Generic centerline point calculator that automatically determines element type
        /// and applies the appropriate centerline calculation method
        /// </summary>
        public static XYZ GetElementCenterlinePoint(Element element, XYZ intersectionPoint, Document hostDocument = null)
        {
            if (element == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Element is null, returning input point");
                return intersectionPoint;
            }
            
            // Determine element type and apply appropriate centerline calculation
            if (element.Category.Id.GetIntegerValue() == (int)BuiltInCategory.OST_Walls)
            {
                if (element is Wall wall)
                {
                    // ✅ STANCE UNIFICATION: Use GetWallCenterlinePointFromBbox (Simple Method)
                    // This uses the wall bounding box center, which is the most reliable geometric center
                    // regardless of Location Line (Face vs Center).
                    // This avoids the complexity of RayTrace and the potential Face/Center confusion of LocationCurve.
                    return GetWallCenterlinePointFromBbox(wall, intersectionPoint, hostDocument);
                }
                else
                    return intersectionPoint;
            }
            else if (element.Category.Id.GetIntegerValue() == (int)BuiltInCategory.OST_StructuralFraming)
            {
                // ✅ CRITICAL FIX: Pass host document so helper can find transform for linked document framing
                return GetStructuralFramingCenterlinePoint(element, intersectionPoint, hostDocument);
            }
            else if (element.Category.Id.GetIntegerValue() == (int)BuiltInCategory.OST_Floors)
            {
                return GetFloorCenterlinePoint(element, intersectionPoint);
            }
            else
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Unsupported element type: {element.Category.Name}, returning input point");
                return intersectionPoint;
            }
        }
    }
}