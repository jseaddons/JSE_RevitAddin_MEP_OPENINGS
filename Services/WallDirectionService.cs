using System;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Centralized service for calculating wall/framing direction and orientation.
    /// Eliminates code duplication between RefreshService and ClashZoneService.
    /// Follows OOP Single Responsibility Principle.
    /// </summary>
    public static class WallDirectionService
    {
        /// <summary>
        /// Calculates wall direction vector from a wall or structural framing element.
        /// Returns the direction the wall/framing runs along (not its normal/perpendicular).
        /// </summary>
        /// <param name="structuralElement">Wall or Structural Framing element</param>
        /// <returns>Direction vector (normalized), or null if cannot be determined</returns>
        public static XYZ GetWallDirection(Element structuralElement)
        {
            try
            {
                if (structuralElement?.Location is LocationCurve lc)
                {
                    var curve = lc.Curve;
                    if (curve is Line line)
                    {
                        return line.Direction.Normalize();
                    }
                    else
                    {
                        // For non-linear curves, use start-to-end direction
                        var start = curve.GetEndPoint(0);
                        var end = curve.GetEndPoint(1);
                        return (end - start).Normalize();
                    }
                }
            }
            catch
            {
                // Return null on error
            }
            
            return null;
        }
        
        /// <summary>
        /// Calculates wall normal vector (perpendicular to wall direction in XY plane).
        /// This is used for orientation calculations and caching structural normal.
        /// </summary>
        /// <param name="structuralElement">Wall or Structural Framing element</param>
        /// <returns>Normal vector (perpendicular to direction in XY plane), or null if cannot be determined</returns>
        public static XYZ GetWallNormal(Element structuralElement)
        {
            var direction = GetWallDirection(structuralElement);
            if (direction == null)
                return null;
            
            // Wall normal is perpendicular in XY plane: rotate direction 90° clockwise
            return new XYZ(-direction.Y, direction.X, 0).Normalize();
        }
        
        /// <summary>
        /// Gets host orientation string ("X" or "Y") based on wall/framing direction.
        /// Uses the same logic as CreateClashZone.GetHostOrientation to ensure consistency.
        /// 
        /// Orientation Logic:
        /// - If wall runs along X-axis (direction primarily in X), orientation = "X"
        /// - If wall runs along Y-axis (direction primarily in Y), orientation = "Y"
        /// </summary>
        /// <param name="structuralElement">Wall or Structural Framing element</param>
        /// <returns>"X" if wall runs along X-axis, "Y" if along Y-axis, empty string for floors/unknown</returns>
        public static string GetHostOrientation(Element structuralElement)
        {
            try
            {
                if (structuralElement is Wall wall)
                {
                    var direction = GetWallDirection(wall);
                    if (direction != null)
                    {
                        // Check if wall is more aligned with X or Y axis
                        double absX = Math.Abs(direction.X);
                        double absY = Math.Abs(direction.Y);
                        
                        if (absX > absY)
                        {
                            return "X"; // Wall runs along X axis
                        }
                        else
                        {
                            return "Y"; // Wall runs along Y axis
                        }
                    }
                }
                else if (structuralElement is FamilyInstance famInst && 
                         famInst.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    var direction = GetWallDirection(famInst);
                    if (direction != null)
                    {
                        double absX = Math.Abs(direction.X);
                        double absY = Math.Abs(direction.Y);
                        
                        if (absX > absY)
                        {
                            return "X";
                        }
                        else
                        {
                            return "Y";
                        }
                    }
                }
                
                // Floors and unknown elements don't need orientation
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Gets wall direction type string (e.g., "X-WALL", "Y-WALL") for wall direction detection.
        /// This is used by UniversalSleevePlacerService for rotation logic.
        /// </summary>
        /// <param name="structuralElement">Wall element</param>
        /// <param name="wallDirection">Pre-calculated wall direction vector (optional - will calculate if null)</param>
        /// <returns>"X-WALL" if wall runs along X-axis, "Y-WALL" if along Y-axis, null if cannot determine</returns>
        public static string GetWallDirectionType(Element structuralElement, XYZ wallDirection = null)
        {
            try
            {
                if (structuralElement is Wall wall)
                {
                    var direction = wallDirection ?? GetWallDirection(wall);
                    if (direction != null)
                    {
                        double absX = Math.Abs(direction.X);
                        double absY = Math.Abs(direction.Y);
                        
                        if (absX > absY)
                        {
                            return "X-WALL";
                        }
                        else
                        {
                            return "Y-WALL";
                        }
                    }
                }
            }
            catch
            {
                // Return null on error
            }
            
            return null;
        }
        
        /// <summary>
        /// Gets structural element normal (perpendicular to direction).
        /// Used for caching and other orientation calculations.
        /// </summary>
        /// <param name="structuralElement">Wall or Structural Framing element</param>
        /// <returns>Normal vector (perpendicular to direction in XY plane), or null if cannot be determined</returns>
        public static XYZ GetStructuralElementNormal(Element structuralElement)
        {
            return GetWallNormal(structuralElement);
        }
    }
}

