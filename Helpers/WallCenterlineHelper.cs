using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class WallCenterlineHelper
    {
        // Returns the centerline point of the wall at a given intersection point, using robust exterior normal
        public static XYZ GetWallCenterlinePoint(Wall wall, XYZ intersectionPoint)
        {
            DebugLogger.Log($"[CENTERLINE-DEBUG] ===== WALL CENTERLINE CALCULATION =====");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Wall ID: {wall?.Id?.Value}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Input intersectionPoint: {intersectionPoint}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Wall orientation: {wall?.Orientation}");
            
            if (wall == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Wall is null, returning input point");
                return intersectionPoint;
            }
            try
            {
                // *** FIXED: Don't project onto wall line, just move perpendicular to wall face ***
                XYZ wallNormal = wall.Orientation.Normalize();
                DebugLogger.Log($"[CENTERLINE-DEBUG] Wall normal (orientation): {wallNormal}");
                
                double wallWidth = wall.Width;
                double halfWidth = wallWidth / 2.0;
                DebugLogger.Log($"[CENTERLINE-DEBUG] Wall width: {UnitUtils.ConvertFromInternalUnits(wallWidth, UnitTypeId.Millimeters):F1}mm, halfWidth: {UnitUtils.ConvertFromInternalUnits(halfWidth, UnitTypeId.Millimeters):F1}mm");
                
                // Move from intersection point toward wall centerline by half wall width
                // Use negative direction to move inward from exterior face
                XYZ centerlinePoint = intersectionPoint + wallNormal * (-halfWidth);
                DebugLogger.Log($"[CENTERLINE-DEBUG] Movement vector: {wallNormal * (-halfWidth)}");
                DebugLogger.Log($"[CENTERLINE-DEBUG] Final centerline point: {centerlinePoint}");
                
                // Calculate and log the offset from original point
                XYZ offset = centerlinePoint - intersectionPoint;
                double offsetDistance = offset.GetLength();
                DebugLogger.Log($"[CENTERLINE-DEBUG] Offset from input: {offset}, distance: {UnitUtils.ConvertFromInternalUnits(offsetDistance, UnitTypeId.Millimeters):F1}mm");
                DebugLogger.Log($"[CENTERLINE-DEBUG] ===== END WALL CENTERLINE CALCULATION =====");
                
                return centerlinePoint;
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
        /// </summary>
        public static XYZ GetStructuralFramingCenterlinePoint(Element structuralFraming, XYZ intersectionPoint)
        {
            DebugLogger.Log($"[CENTERLINE-DEBUG] ===== STRUCTURAL FRAMING CENTERLINE CALCULATION =====");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Structural Framing ID: {structuralFraming?.Id?.Value}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Input intersectionPoint: {intersectionPoint}");
            
            if (structuralFraming == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Structural framing is null, returning input point");
                return intersectionPoint;
            }
            
            // Check if this is a structural framing element
            if (structuralFraming.Category.Id.Value != (int)BuiltInCategory.OST_StructuralFraming)
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
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Framing {structuralFraming.Id.Value} is not structural (StructuralType={structuralType}), returning input point");
                    return intersectionPoint;
                }
                DebugLogger.Log($"[CENTERLINE-DEBUG] Framing {structuralFraming.Id.Value} is structural (StructuralType={structuralType}), proceeding with centerline calculation");
            }
            
            try
            {
                // Get structural framing thickness following StructuralSleevePlacementCommand pattern
                double framingThickness = GetStructuralFramingThickness(structuralFraming);
                double halfThickness = framingThickness / 2.0;
                DebugLogger.Log($"[CENTERLINE-DEBUG] Framing thickness: {UnitUtils.ConvertFromInternalUnits(framingThickness, UnitTypeId.Millimeters):F1}mm, halfThickness: {UnitUtils.ConvertFromInternalUnits(halfThickness, UnitTypeId.Millimeters):F1}mm");
                
                // For structural framing, we need to determine the normal direction
                // This follows the approach from StructuralSleevePlacementCommand for centerline calculation
                XYZ framingNormal = GetStructuralFramingNormal(structuralFraming);
                DebugLogger.Log($"[CENTERLINE-DEBUG] Framing normal: {framingNormal}");
                
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
            DebugLogger.Log($"[CENTERLINE-DEBUG] Floor ID: {floor?.Id?.Value}");
            DebugLogger.Log($"[CENTERLINE-DEBUG] Input intersectionPoint: {intersectionPoint}");
            
            if (floor == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Floor is null, returning input point");
                return intersectionPoint;
            }
            
            // Check if this is a floor element
            if (floor.Category.Id.Value != (int)BuiltInCategory.OST_Floors)
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
                        DebugLogger.Log($"[CENTERLINE-DEBUG] Floor {floor.Id.Value} is not structural (isStructural={isStructural}), returning input point");
                        return intersectionPoint;
                    }
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Floor {floor.Id.Value} is structural (isStructural={isStructural}), proceeding with centerline calculation");
                }
                else
                {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Floor {floor.Id.Value} has no structural usage parameter, proceeding with centerline calculation");
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
                throw new InvalidOperationException($"Cannot determine structural framing depth: 'b' or 'Width' parameter not set for element ID {structuralFraming.Id.Value}");
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
        /// Generic centerline point calculator that automatically determines element type
        /// and applies the appropriate centerline calculation method
        /// </summary>
        public static XYZ GetElementCenterlinePoint(Element element, XYZ intersectionPoint)
        {
            if (element == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Element is null, returning input point");
                return intersectionPoint;
            }
            
            // Determine element type and apply appropriate centerline calculation
            if (element.Category.Id.Value == (int)BuiltInCategory.OST_Walls)
            {
                if (element is Wall wall)
                    return GetWallCenterlinePoint(wall, intersectionPoint);
                else
                    return intersectionPoint;
            }
            else if (element.Category.Id.Value == (int)BuiltInCategory.OST_StructuralFraming)
            {
                return GetStructuralFramingCenterlinePoint(element, intersectionPoint);
            }
            else if (element.Category.Id.Value == (int)BuiltInCategory.OST_Floors)
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
