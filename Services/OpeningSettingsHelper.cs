using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Helper service for handling opening-related settings from SettingsModel
    /// </summary>
    public static class OpeningSettingsHelper
    {
        /// <summary>
        /// Determines the opening type for a given MEP category based on settings
        /// </summary>
        /// <param name="mepCategory">The MEP category (e.g., "Pipes", "Ducts", "Cable Trays")</param>
        /// <returns>The opening type: "Circular" or "Rectangular"</returns>
        public static string GetOpeningTypeForCategory(string mepCategory)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                
                if (mepCategory.Equals("Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    // REMOVED: PipeOpeningTypeRectangular global setting that forced ALL pipes to rectangular
                    // Now pipes will respect UI selection and size threshold rules
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[OPENING_SETTINGS] Pipe opening type will be determined by UI selection and size threshold rules");
                    return "Circular"; // Default, will be overridden by UI selection and size rules
                }
                else if (mepCategory.Equals("Ducts", StringComparison.OrdinalIgnoreCase))
                {
                    // For ducts, we need to get the opening type from the UI (EmergencyMainDialog)
                    // This will be handled by the DuctSleeveCommand which reads from the UI directly
                    // For now, return a placeholder that will be overridden by the command
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[OPENING_SETTINGS] Duct opening type will be determined by UI selection for round ducts");
                    return "Circular"; // Default, will be overridden by UI selection
                }
                else
                {
                    // All other categories (Duct Accessories, Cable Trays) are always rectangular
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[OPENING_SETTINGS] {mepCategory} opening type set to Rectangular (default for non-pipe categories)");
                    return "Rectangular";
                }
            }
            catch (Exception ex)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[OPENING_SETTINGS] Error getting opening type for category '{mepCategory}': {ex.Message}");
                // Fallback to rectangular for safety
                return "Rectangular";
            }
        }

        /// <summary>
        /// Rounds opening dimensions based on user-configured rounding value and always-up setting
        /// </summary>
        /// <param name="dimension">The dimension to potentially round</param>
        /// <returns>The rounded dimension, or original dimension if rounding is disabled</returns>
        public static double RoundDimensionToNearest5mm(double dimension)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                
                // Get the rounding value from settings (default 5mm if 0)
                double roundingValue = settings.RoundingValue;
                if (roundingValue <= 0)
                {
                    // Rounding is disabled, return original
                    return dimension;
                }
                
                // Convert to millimeters
                double mmDimension = UnitUtils.ConvertFromInternalUnits(dimension, UnitTypeId.Millimeters);
                double roundedMm;
                
                if (settings.RoundAlwaysUp)
                {
                    // ✅ FIX: Add tolerance for floating point errors (e.g. 650.0001 -> 700)
                    // If we are effectively AT the rounding boundary (within 0.1mm), stay there.
                    double remainder = mmDimension % roundingValue;
                    if (remainder > 0 && remainder < 0.1) // Tolerance: 0.1mm
                    {
                        // We are just slightly over (e.g. 650.0001 with rounding 50) -> Treat as 650
                        mmDimension -= remainder; 
                    }
                    else if (remainder > (roundingValue - 0.1)) // Tolerance: 0.1mm on the other side
                    {
                         // We are just slightly under next step (e.g. 699.999) -> Treat as 700
                         mmDimension += (roundingValue - remainder);
                    }

                    // Always round up: 453 → 500 (with rounding value 50)
                    roundedMm = Math.Ceiling(mmDimension / roundingValue) * roundingValue;
                }
                else
                {
                    // Round to nearest: 453 → 450 (with rounding value 50)
                    roundedMm = Math.Round(mmDimension / roundingValue) * roundingValue;
                }
                
                // Convert back to internal units
                double roundedDimension = UnitUtils.ConvertToInternalUnits(roundedMm, UnitTypeId.Millimeters);
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[OPENING_SETTINGS] Rounded dimension from {UnitUtils.ConvertFromInternalUnits(dimension, UnitTypeId.Millimeters):F3}mm to {roundedMm:F1}mm (rounding value: {roundingValue}, always up: {settings.RoundAlwaysUp})");
                return roundedDimension;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[OPENING_SETTINGS] Error rounding dimension: {ex.Message}");
                // Return original dimension on error
                return dimension;
            }
        }

        /// <summary>
        /// Rounds both width and height dimensions based on user-configured rounding value
        /// </summary>
        /// <param name="width">The width dimension</param>
        /// <param name="height">The height dimension</param>
        /// <returns>A tuple containing the rounded dimensions</returns>
        public static (double width, double height) RoundDimensionsToNearest5mm(double width, double height)
        {
            return (RoundDimensionToNearest5mm(width), RoundDimensionToNearest5mm(height));
        }

        /// <summary>
        /// ✅ CLUSTER-SPECIFIC: Rounds dimension for clusters with special logic:
        /// If RoundAlwaysUp is true and the difference to the next increment is less than 1mm, round down instead.
        /// This prevents small rounding errors (e.g., 750.1mm) from jumping to the next increment (800mm).
        /// </summary>
        /// <param name="dimension">The dimension to round</param>
        /// <returns>The rounded dimension</returns>
        public static double RoundDimensionForCluster(double dimension)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                
                // Get the rounding value from settings (default 5mm if 0)
                double roundingValue = settings.RoundingValue;
                if (roundingValue <= 0)
                {
                    // Rounding is disabled, return original
                    return dimension;
                }
                
                // Convert to millimeters
                double mmDimension = UnitUtils.ConvertFromInternalUnits(dimension, UnitTypeId.Millimeters);
                double roundedMm;
                
                if (settings.RoundAlwaysUp)
                {
                    // Calculate lower and upper increments
                    double lowerIncrement = Math.Floor(mmDimension / roundingValue) * roundingValue;
                    double upperIncrement = Math.Ceiling(mmDimension / roundingValue) * roundingValue;
                    
                    // Calculate difference from the lower increment
                    double differenceFromLower = mmDimension - lowerIncrement;
                    
                    // ✅ CLUSTER-SPECIFIC: If the value is within 1mm of the lower increment, use lower instead of rounding up
                    // Example: 750.1mm with 50mm increment → difference from 750mm is 0.1mm (< 1mm) → stay at 750mm instead of 800mm
                    if (differenceFromLower < 1.0 && differenceFromLower > 0.0)
                    {
                        // Use the lower increment (don't round up)
                        roundedMm = lowerIncrement;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[OPENING_SETTINGS] [CLUSTER] Rounded down due to <1mm difference from lower increment: {mmDimension:F3}mm → {roundedMm:F1}mm (would have been {upperIncrement:F1}mm, diff={differenceFromLower:F3}mm)");
                        }
                    }
                    else
                    {
                        // Normal round up behavior
                        roundedMm = upperIncrement;
                    }
                }
                else
                {
                    // Round to nearest: 453 → 450 (with rounding value 50)
                    roundedMm = Math.Round(mmDimension / roundingValue) * roundingValue;
                }
                
                // Convert back to internal units
                double roundedDimension = UnitUtils.ConvertToInternalUnits(roundedMm, UnitTypeId.Millimeters);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[OPENING_SETTINGS] [CLUSTER] Rounded dimension from {mmDimension:F3}mm to {roundedMm:F1}mm (rounding value: {roundingValue}, always up: {settings.RoundAlwaysUp})");
                }
                return roundedDimension;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[OPENING_SETTINGS] [CLUSTER] Error rounding dimension: {ex.Message}");
                }
                // Return original dimension on error
                return dimension;
            }
        }

        /// <summary>
        /// ✅ CLUSTER-SPECIFIC: Rounds both width and height dimensions for clusters with special logic
        /// </summary>
        /// <param name="width">The width dimension</param>
        /// <param name="height">The height dimension</param>
        /// <returns>A tuple containing the rounded dimensions</returns>
        public static (double width, double height) RoundDimensionsForCluster(double width, double height)
        {
            return (RoundDimensionForCluster(width), RoundDimensionForCluster(height));
        }

        /// <summary>
        /// Rounds diameter dimension based on user-configured rounding value
        /// </summary>
        /// <param name="diameter">The diameter dimension</param>
        /// <returns>The rounded diameter</returns>
        public static double RoundDiameterToNearest5mm(double diameter)
        {
            return RoundDimensionToNearest5mm(diameter);
        }

        /// <summary>
        /// Gets a summary of current opening settings for logging/debugging
        /// </summary>
        /// <returns>A string describing the current opening settings</returns>
        public static string GetOpeningSettingsSummary()
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                
                var pipeType = settings.PipeOpeningTypeRectangular ? "Rectangular" : "Circular";
                var roundingMode = settings.RoundingValue > 0 
                    ? $"Rounding to nearest {settings.RoundingValue}mm (Always Up: {settings.RoundAlwaysUp})" 
                    : "Rounding disabled";
                
                return $"Opening Settings - Pipes: {pipeType}, {roundingMode}";
            }
            catch (Exception ex)
            {
                return $"Error getting opening settings: {ex.Message}";
            }
        }
    }
}
