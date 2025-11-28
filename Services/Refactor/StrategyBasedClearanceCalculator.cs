using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactor
{
    /// <summary>
    /// Strategy-based clearance calculator that delegates to category-specific providers.
    /// Implements IClearanceStrategy for high-level clearance calculation.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    public class StrategyBasedClearanceCalculator : IClearanceStrategy
    {
        private readonly ClearanceProviderFactory _providerFactory;
        private readonly int _parallelThreshold;
        
        public string StrategyName => "StrategyBasedClearanceCalculator";
        
        /// <summary>
        /// Constructor with dependency injection for clearance provider factory.
        /// </summary>
        /// <param name="providerFactory">Factory for creating category-specific clearance providers</param>
        /// <param name="parallelThreshold">Zone count threshold for parallel processing (default: 10)</param>
        public StrategyBasedClearanceCalculator(
            ClearanceProviderFactory providerFactory = null,
            int parallelThreshold = 10)
        {
            _providerFactory = providerFactory ?? new ClearanceProviderFactory();
            _parallelThreshold = parallelThreshold;
        }
        
        /// <summary>
        /// Calculate clearance for a clash zone based on MEP element and structural host.
        /// Delegates to category-specific clearance provider.
        /// </summary>
        public double CalculateClearance(ClashZone zone, OpeningConditions conditions)
        {
            if (zone == null)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning("[StrategyBasedClearance] Null zone provided, returning default clearance");
                }
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50); // 50mm default
            }

            try
            {
                // Get MEP element from zone - ClashZone stores ID, not element reference
                // TODO: Accept Document parameter to resolve element
                // For now, use zone's pre-calculated category for clearance lookup
                if (string.IsNullOrEmpty(zone.MepElementCategory))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[StrategyBasedClearance] Zone {zone.Id} has no MEP element category");
                    }
                    return RevitUnitConversionService.Instance.ToInternalMillimeters(50);
                }

                // Use category string to get provider (fallback until Document injection added)
                // TODO: Need Document parameter to resolve zone.MepElementId to actual Element
                // For now, use default clearance based on category
                double clearance = RevitUnitConversionService.Instance.ToInternalMillimeters(50); // Default 50mm
                
                // Attempt to use category-based clearance if available
                if (!string.IsNullOrEmpty(zone.MepElementCategory))
                {
                    // Map category to default clearance values
                    // TODO: Replace with actual provider.GetClearance(element, uiClearances) when Document available
                    clearance = GetDefaultClearanceByCategory(zone.MepElementCategory);
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    // Convert clearance to mm for logging
                    double clearanceMm = UnitUtils.ConvertFromInternalUnits(clearance, UnitTypeId.Millimeters);
                    DebugLogger.Info($"[StrategyBasedClearance] Zone {zone.Id}, Category: {zone.MepElementCategory}, " +
                        $"Clearance: {clearanceMm:F1}mm");
                }
                
                return clearance;
            }
            catch (System.Exception ex)
            {
                // Fail-safe: Return default clearance on error
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[StrategyBasedClearance] Error calculating clearance for zone {zone.Id}: {ex.Message}");
                }
                return RevitUnitConversionService.Instance.ToInternalMillimeters(50);
            }
        }
        
        /// <summary>
        /// Determine if clearance calculation should use parallel processing.
        /// </summary>
        public bool ShouldUseParallelCalculation(int zoneCount)
        {
            return zoneCount > _parallelThreshold;
        }
        
        /// <summary>
        /// Extract UI clearance overrides from OpeningConditions.
        /// </summary>
        private Dictionary<string, double> ExtractUIClearances(OpeningConditions conditions)
        {
            // This is a placeholder - actual implementation would extract clearance values
            // from the OpeningConditions object based on its structure
            // For now, return null to use provider defaults
            return null;
        }
        
        /// <summary>
        /// Get default clearance by MEP category string (temporary until Document injection added).
        /// </summary>
        private double GetDefaultClearanceByCategory(string category)
        {
            // Default clearances in mm, converted to internal units (feet)
            var defaultMm = 50.0; // Standard clearance
            
            if (category.Contains("Pipe", StringComparison.OrdinalIgnoreCase))
                defaultMm = 50.0;
            else if (category.Contains("Duct", StringComparison.OrdinalIgnoreCase))
                defaultMm = 50.0;
            else if (category.Contains("CableTray", StringComparison.OrdinalIgnoreCase))
                defaultMm = 50.0;
            else if (category.Contains("Conduit", StringComparison.OrdinalIgnoreCase))
                defaultMm = 50.0;
            
            return RevitUnitConversionService.Instance.ToInternalMillimeters(defaultMm);
        }
    }
}
