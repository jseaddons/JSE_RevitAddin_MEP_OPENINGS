using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders
{
    /// <summary>
    /// Service for calculating clearance values from multiple sources (database, UI, strategy).
    /// Follows SRP by centralizing clearance calculation logic.
    /// 
    /// Priority order (when batch writing is enabled):
    /// 1. Database (_conditions.ClearanceSettings) - authoritative source
    /// 2. Strategy calculation (may use UI settings)
    /// 3. UI settings dictionary (_clearanceSettings)
    /// 4. Default fallback (50mm)
    /// </summary>
    public class ClearanceCalculationService
    {
        private readonly ClearanceProviderFactory _providerFactory;

        public ClearanceCalculationService()
        {
            _providerFactory = new ClearanceProviderFactory();
        }

        /// <summary>
        /// Get clearance for a clash zone, prioritizing database settings when batch writing is enabled.
        /// </summary>
        /// <param name="zone">The clash zone</param>
        /// <param name="conditions">Opening conditions with database clearance settings</param>
        /// <param name="uiClearances">UI clearance settings dictionary (optional)</param>
        /// <param name="strategy">Placement strategy for fallback calculation (optional)</param>
        /// <returns>Clearance value in internal units (feet)</returns>
        public double GetClearance(
            ClashZone zone,
            OpeningConditions conditions = null,
            Dictionary<string, double> uiClearances = null,
            ISleevePlacementStrategy strategy = null)
        {
            // ✅ PRIORITY 1: Read directly from conditions.ClearanceSettings (database) if available
            // This is the authoritative source and works correctly even when batching is enabled
            if (conditions?.ClearanceSettings != null)
            {
                try
                {
                    double clearanceMm = GetClearanceFromDatabase(zone, conditions.ClearanceSettings);
                    if (clearanceMm > 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[ClearanceCalculationService] ✅ Using DATABASE clearance: {clearanceMm}mm for {zone.MepElementCategory} (isInsulated={zone.IsInsulated})");
                        return RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceMm);
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[ClearanceCalculationService] Error reading clearance from database for zone {zone.Id}: {ex.Message}");
                }
            }

            // ✅ PRIORITY 2: Fallback to strategy calculation (may use UI settings)
            if (strategy != null)
            {
                try
                {
                    var mepSize = new MepElementSize
                    {
                        Width = zone.MepElementWidth,
                        Height = zone.MepElementHeight,
                        Diameter = zone.MepElementOuterDiameter,
                        Shape = zone.MepElementOuterDiameter > 0 ? "Round" : "Rectangular"
                    };

                    double clearance = strategy.GetClearance(mepSize, conditions);
                    if (clearance > 0)
                    {
                        return clearance;
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[ClearanceCalculationService] Strategy clearance calculation failed for zone {zone.Id}: {ex.Message}");
                }
            }

            // ✅ PRIORITY 3: Fallback to UI settings dictionary
            if (uiClearances != null && uiClearances.Count > 0)
            {
                string key = $"{zone.StructuralElementType}_{zone.MepElementCategory}";
                if (uiClearances.ContainsKey(key))
                {
                    return uiClearances[key];
                }
            }

            // ✅ PRIORITY 4: Final fallback: 50mm default
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Warning($"[ClearanceCalculationService] ⚠️ Using default clearance (50mm) for zone {zone.Id} - no database or UI settings found");
            return RevitUnitConversionService.Instance.ToInternalMillimeters(50);
        }

        /// <summary>
        /// Get clearance from database ClearanceSettings based on zone category and insulation status.
        /// </summary>
        private double GetClearanceFromDatabase(ClashZone zone, ClearanceSettings clearanceSettings)
        {
            bool isInsulated = zone.IsInsulated;
            double clearanceMm = 0.0;

            if (string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                clearanceMm = isInsulated
                    ? clearanceSettings.PipesInsulated
                    : clearanceSettings.PipesNormal;
            }
            else if (string.Equals(zone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                // Check if round or rectangular
                bool isRound = zone.MepElementOuterDiameter > 0;
                if (isRound)
                {
                    clearanceMm = isInsulated
                        ? clearanceSettings.RoundInsulated
                        : clearanceSettings.RoundNormal;
                }
                else
                {
                    clearanceMm = isInsulated
                        ? clearanceSettings.RectangularInsulated
                        : clearanceSettings.RectangularNormal;
                }
            }
            // Note: Cable trays use GetCableTrayPlacementAdjustment in strategy, not this method

            return clearanceMm;
        }
    }
}

