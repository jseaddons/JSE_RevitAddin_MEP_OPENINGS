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
            // ✅ PRIORITY 1: Read directly from conditions.ClearanceSettings (database) - ALWAYS USE DB VALUES
            // This is the authoritative source - use whatever is stored in database, no fallback
            if (conditions?.ClearanceSettings != null)
            {
                try
                {
                    double clearanceMm = GetClearanceFromDatabase(zone, conditions.ClearanceSettings);
                    
                    // ✅ DIAGNOSTIC: Log what we're reading from database
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[ClearanceCalculationService] 🔍 DATABASE CHECK: Zone {zone.Id}, Category={zone.MepElementCategory}, " +
                            $"IsInsulated={zone.IsInsulated}, ClearanceMm={clearanceMm}mm, " +
                            $"RectNormal={conditions.ClearanceSettings.RectangularNormal}mm, " +
                            $"RoundNormal={conditions.ClearanceSettings.RoundNormal}mm, " +
                            $"PipesNormal={conditions.ClearanceSettings.PipesNormal}mm, " +
                            $"DuctAccessoryMepNormal={conditions.ClearanceSettings.DuctAccessoryMepNormal}mm, " +
                            $"DuctAccessoryOtherNormal={conditions.ClearanceSettings.DuctAccessoryOtherNormal}mm");
                        
                        // ✅ COMPREHENSIVE LOGGING: Log to file for traceability
                        SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [CLEARANCE-DB] Zone {zone.Id}, Category={zone.MepElementCategory}, " +
                            $"IsInsulated={zone.IsInsulated}, ClearanceMm={clearanceMm}mm, " +
                            $"DuctAccessoryMepNormal={conditions.ClearanceSettings.DuctAccessoryMepNormal}mm, " +
                            $"DuctAccessoryMepInsulated={conditions.ClearanceSettings.DuctAccessoryMepInsulated}mm, " +
                            $"DuctAccessoryOtherNormal={conditions.ClearanceSettings.DuctAccessoryOtherNormal}mm, " +
                            $"DuctAccessoryOtherInsulated={conditions.ClearanceSettings.DuctAccessoryOtherInsulated}mm\n");
                    }
                    
                    // ✅ USE DATABASE VALUE DIRECTLY: No validation, no fallback - use whatever is in DB
                    if (!double.IsNaN(clearanceMm) && !double.IsInfinity(clearanceMm))
                    {
                        double clearanceInFeet = RevitUnitConversionService.Instance.ToInternalMillimeters(clearanceMm);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ClearanceCalculationService] ✅ Using DATABASE clearance: {clearanceMm}mm ({clearanceInFeet:F6}ft) for {zone.MepElementCategory} (isInsulated={zone.IsInsulated})");
                            SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [CLEARANCE-USED] Zone {zone.Id}: {clearanceMm}mm ({clearanceInFeet:F6}ft) from DATABASE\n");
                        }
                        return clearanceInFeet;
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
            {
                DebugLogger.Warning($"[ClearanceCalculationService] ⚠️ Using default clearance (50mm) for zone {zone.Id} - no database or UI settings found");
                SafeFileLogger.SafeAppendText("clearance_calculation_trace.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [CLEARANCE-FALLBACK] Zone {zone.Id}: Using DEFAULT 50mm - no database/UI/strategy clearance found\n");
            }
            return RevitUnitConversionService.Instance.ToInternalMillimeters(50);
        }

        /// <summary>
        /// Get clearance from database ClearanceSettings based on zone category and insulation status.
        /// ✅ CRITICAL: Uses database values directly - no fallback, no validation.
        /// </summary>
        private double GetClearanceFromDatabase(ClashZone zone, ClearanceSettings clearanceSettings)
        {
            bool isInsulated = zone.IsInsulated;
            double clearanceMm = 0.0;

            // ✅ FIX: Duct Accessories (Dampers) use special clearance values, NOT duct clearances
            if (string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                // ✅ CRITICAL: Dampers use DuctAccessoryOtherNormal/Insulated for general clearance
                // The asymmetric clearance (MEP side vs Other side) is handled in DamperPlacementStrategy
                // This method returns the "other" side clearance (used for non-connector sides)
                clearanceMm = isInsulated
                    ? clearanceSettings.DuctAccessoryOtherInsulated
                    : clearanceSettings.DuctAccessoryOtherNormal;
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[ClearanceCalculationService] 🔍 DUCT ACCESSORY: Using DuctAccessoryOther{(isInsulated ? "Insulated" : "Normal")}={clearanceMm}mm " +
                        $"(asymmetric clearance handled in strategy: MEP={(isInsulated ? clearanceSettings.DuctAccessoryMepInsulated : clearanceSettings.DuctAccessoryMepNormal)}mm, " +
                        $"Other={clearanceMm}mm, IsInsulated={isInsulated})");
                }
            }
            else if (string.Equals(zone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
            {
                clearanceMm = isInsulated
                    ? clearanceSettings.PipesInsulated
                    : clearanceSettings.PipesNormal;
            }
            else if (string.Equals(zone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase))
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

            // ✅ USE DATABASE VALUE DIRECTLY: Return whatever is in database, even if 0.0
            return clearanceMm;
        }
    }
}

