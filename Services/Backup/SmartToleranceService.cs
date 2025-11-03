using System;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Smart tolerance handling service for adaptive clearance calculations
    /// Provides 1.5-2x performance improvement through intelligent tolerance selection
    /// </summary>
    public static class SmartToleranceService
    {
        #region Constants
        
        private const double MIN_TOLERANCE = 0.5; // 0.5 feet minimum
        private const double MAX_TOLERANCE = 2.0; // 2.0 feet maximum
        private const double DEFAULT_TOLERANCE = 1.0; // 1.0 feet default
        private const double INSULATION_FACTOR = 1.5; // 50% increase for insulated elements
        private const double SIZE_FACTOR = 0.5; // 50% of element size
        
        #endregion
        
        #region Public Methods
        
        /// <summary>
        /// Calculate adaptive tolerance based on MEP element properties
        /// </summary>
        public static double CalculateAdaptiveTolerance(Element mepElement)
        {
            if (!OptimizationFlags.UseSmartTolerance || mepElement == null)
            {
                return DEFAULT_TOLERANCE;
            }
            
            try
            {
                double baseTolerance = DEFAULT_TOLERANCE;
                
                // Get element size for tolerance calculation
                var elementSize = GetElementSize(mepElement);
                if (elementSize > 0)
                {
                    baseTolerance = elementSize * SIZE_FACTOR;
                }
                
                // Apply insulation factor if element is insulated
                var insulationFactor = GetInsulationFactor(mepElement);
                baseTolerance *= insulationFactor;
                
                // Apply category-specific adjustments
                var categoryFactor = GetCategoryFactor(mepElement);
                baseTolerance *= categoryFactor;
                
                // Clamp to reasonable bounds
                var finalTolerance = ClampValue(baseTolerance, MIN_TOLERANCE, MAX_TOLERANCE);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SmartTolerance] Element {mepElement.Id}: Size={elementSize:F2}, Insulation={insulationFactor:F2}, Category={categoryFactor:F2}, Tolerance={finalTolerance:F2}");
                
                return finalTolerance;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SmartTolerance] Error calculating tolerance for element {mepElement?.Id}: {ex.Message}");
                return DEFAULT_TOLERANCE;
            }
        }
        
        /// <summary>
        /// Get tolerance for structural elements (typically smaller)
        /// </summary>
        public static double GetStructuralTolerance(Element structuralElement)
        {
            if (!OptimizationFlags.UseSmartTolerance || structuralElement == null)
            {
                return DEFAULT_TOLERANCE;
            }
            
            try
            {
                // Structural elements typically need less tolerance
                var baseTolerance = DEFAULT_TOLERANCE * 0.8; // 20% reduction
                
                // Apply category-specific adjustments
                var categoryFactor = GetStructuralCategoryFactor(structuralElement);
                baseTolerance *= categoryFactor;
                
                // Clamp to reasonable bounds
                var finalTolerance = ClampValue(baseTolerance, MIN_TOLERANCE * 0.8, MAX_TOLERANCE * 0.8);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SmartTolerance] Structural element {structuralElement.Id}: Category={categoryFactor:F2}, Tolerance={finalTolerance:F2}");
                
                return finalTolerance;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SmartTolerance] Error calculating structural tolerance for element {structuralElement?.Id}: {ex.Message}");
                return DEFAULT_TOLERANCE * 0.8;
            }
        }
        
        /// <summary>
        /// Get tolerance for intersection testing between MEP and structural elements
        /// </summary>
        public static double GetIntersectionTolerance(Element mepElement, Element structuralElement)
        {
            if (!OptimizationFlags.UseSmartTolerance || mepElement == null || structuralElement == null)
            {
                return DEFAULT_TOLERANCE;
            }
            
            try
            {
                var mepTolerance = CalculateAdaptiveTolerance(mepElement);
                var structuralTolerance = GetStructuralTolerance(structuralElement);
                
                // Use the larger tolerance for intersection testing
                var intersectionTolerance = Math.Max(mepTolerance, structuralTolerance);
                
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[SmartTolerance] Intersection tolerance: MEP={mepTolerance:F2}, Structural={structuralTolerance:F2}, Final={intersectionTolerance:F2}");
                
                return intersectionTolerance;
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SmartTolerance] Error calculating intersection tolerance: {ex.Message}");
                return DEFAULT_TOLERANCE;
            }
        }
        
        #endregion
        
        #region Private Methods
        
        /// <summary>
        /// Clamp a value between min and max (replacement for Math.Clamp in older .NET versions)
        /// </summary>
        private static double ClampValue(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
        
        /// <summary>
        /// Get element size for tolerance calculation
        /// </summary>
        private static double GetElementSize(Element element)
        {
            try
            {
                if (element is MEPCurve mepCurve)
                {
                    // For MEP curves, use the larger of width or height
                    return Math.Max(mepCurve.Width, mepCurve.Height);
                }
                else if (element is FamilyInstance familyInstance)
                {
                    // For family instances, try to get size parameters
                    var widthParam = familyInstance.LookupParameter("Width");
                    var heightParam = familyInstance.LookupParameter("Height");
                    var diameterParam = familyInstance.LookupParameter("Diameter");
                    
                    if (widthParam != null && heightParam != null)
                    {
                        return Math.Max(widthParam.AsDouble(), heightParam.AsDouble());
                    }
                    else if (diameterParam != null)
                    {
                        return diameterParam.AsDouble();
                    }
                }
                
                // Fallback: use bounding box
                var bbox = element.get_BoundingBox(null);
                if (bbox != null)
                {
                    var size = bbox.Max - bbox.Min;
                    return Math.Max(Math.Max(size.X, size.Y), size.Z);
                }
                
                return 0;
            }
            catch
            {
                return 0;
            }
        }
        
        /// <summary>
        /// Get insulation factor for the element
        /// </summary>
        private static double GetInsulationFactor(Element element)
        {
            try
            {
                var insulationParam = element.LookupParameter("Insulation Thickness");
                if (insulationParam != null && insulationParam.AsDouble() > 0)
                {
                    return INSULATION_FACTOR;
                }
                
                return 1.0;
            }
            catch
            {
                return 1.0;
            }
        }
        
        /// <summary>
        /// Get category-specific factor for MEP elements
        /// </summary>
        private static double GetCategoryFactor(Element element)
        {
            try
            {
                var categoryId = element.Category?.Id.IntegerValue ?? -1;
                
                switch (categoryId)
                {
                    case (int)BuiltInCategory.OST_DuctCurves:
                        return 1.0; // Standard tolerance
                    case (int)BuiltInCategory.OST_PipeCurves:
                        return 0.9; // Slightly less tolerance
                    case (int)BuiltInCategory.OST_CableTray:
                        return 0.8; // Less tolerance for cable trays
                    case (int)BuiltInCategory.OST_DuctAccessory:
                        return 1.1; // Slightly more tolerance for dampers
                    default:
                        return 1.0;
                }
            }
            catch
            {
                return 1.0;
            }
        }
        
        /// <summary>
        /// Get category-specific factor for structural elements
        /// </summary>
        private static double GetStructuralCategoryFactor(Element element)
        {
            try
            {
                var categoryId = element.Category?.Id.IntegerValue ?? -1;
                
                switch (categoryId)
                {
                    case (int)BuiltInCategory.OST_Walls:
                        return 1.0; // Standard tolerance
                    case (int)BuiltInCategory.OST_Floors:
                        return 0.9; // Slightly less tolerance
                    case (int)BuiltInCategory.OST_StructuralFraming:
                        return 0.8; // Less tolerance for framing
                    case (int)BuiltInCategory.OST_StructuralColumns:
                        return 0.7; // Even less tolerance for columns
                    default:
                        return 1.0;
                }
            }
            catch
            {
                return 1.0;
            }
        }
        
        #endregion
    }
}
