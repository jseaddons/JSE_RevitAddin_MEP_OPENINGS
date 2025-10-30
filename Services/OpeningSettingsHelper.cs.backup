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
                    DebugLogger.Info($"[OPENING_SETTINGS] Pipe opening type will be determined by UI selection and size threshold rules");
                    return "Circular"; // Default, will be overridden by UI selection and size rules
                }
                else if (mepCategory.Equals("Ducts", StringComparison.OrdinalIgnoreCase))
                {
                    // For ducts, we need to get the opening type from the UI (EmergencyMainDialog)
                    // This will be handled by the DuctSleeveCommand which reads from the UI directly
                    // For now, return a placeholder that will be overridden by the command
                    DebugLogger.Info($"[OPENING_SETTINGS] Duct opening type will be determined by UI selection for round ducts");
                    return "Circular"; // Default, will be overridden by UI selection
                }
                else
                {
                    // All other categories (Duct Accessories, Cable Trays) are always rectangular
                    DebugLogger.Info($"[OPENING_SETTINGS] {mepCategory} opening type set to Rectangular (default for non-pipe categories)");
                    return "Rectangular";
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OPENING_SETTINGS] Error getting opening type for category '{mepCategory}': {ex.Message}");
                // Fallback to rectangular for safety
                return "Rectangular";
            }
        }

        /// <summary>
        /// Rounds opening dimensions to the nearest 5mm if the setting is enabled
        /// </summary>
        /// <param name="dimension">The dimension to potentially round</param>
        /// <returns>The rounded dimension if setting is enabled, otherwise the original dimension</returns>
        public static double RoundDimensionToNearest5mm(double dimension)
        {
            try
            {
                var settings = ApplicationProfileService.Instance.GetCurrentSettings();
                
                if (settings.RoundOpeningSizesToNearest5mm)
                {
                    // Convert to millimeters, round to nearest 5mm, then convert back to internal units
                    double mmDimension = UnitUtils.ConvertFromInternalUnits(dimension, UnitTypeId.Millimeters);
                    double roundedMm = Math.Round(mmDimension / 5.0) * 5.0;
                    double roundedDimension = UnitUtils.ConvertToInternalUnits(roundedMm, UnitTypeId.Millimeters);
                    
                    DebugLogger.Info($"[OPENING_SETTINGS] Rounded dimension from {mmDimension:F1}mm to {roundedMm:F1}mm");
                    return roundedDimension;
                }
                else
                {
                    // Setting is disabled, return original dimension
                    return dimension;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OPENING_SETTINGS] Error rounding dimension: {ex.Message}");
                // Return original dimension on error
                return dimension;
            }
        }

        /// <summary>
        /// Rounds both width and height dimensions to the nearest 5mm if the setting is enabled
        /// </summary>
        /// <param name="width">The width dimension</param>
        /// <param name="height">The height dimension</param>
        /// <returns>A tuple containing the rounded dimensions</returns>
        public static (double width, double height) RoundDimensionsToNearest5mm(double width, double height)
        {
            return (RoundDimensionToNearest5mm(width), RoundDimensionToNearest5mm(height));
        }

        /// <summary>
        /// Rounds diameter dimension to the nearest 5mm if the setting is enabled
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
                var roundingEnabled = settings.RoundOpeningSizesToNearest5mm ? "Enabled" : "Disabled";
                
                return $"Opening Settings - Pipes: {pipeType}, Rounding to 5mm: {roundingEnabled}";
            }
            catch (Exception ex)
            {
                return $"Error getting opening settings: {ex.Message}";
            }
        }
    }
}
