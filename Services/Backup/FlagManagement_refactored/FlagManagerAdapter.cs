using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Team F: Adapter that wraps legacy FlagManager for coexistence.
    /// 
    /// This adapter allows gradual migration from legacy FlagManager to refactored services.
    /// When UseRefactoredClashZoneFlagServices is false, this adapter delegates to legacy FlagManager.
    /// When UseRefactoredClashZoneFlagServices is true, this adapter uses refactored services.
    /// 
    /// SOLID: Adapter Pattern - enables coexistence without breaking existing code.
    /// </summary>
    public class FlagManagerAdapter : IFlagManager
    {
        private readonly IFlagManager _refactoredService;
        private readonly FlagManager _legacyService;
        private readonly bool _useRefactored;
        
        /// <summary>
        /// Creates a flag manager adapter.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="refactoredService">Refactored flag manager service (can be null)</param>
        /// <param name="legacyService">Legacy flag manager (can be null)</param>
        public FlagManagerAdapter(
            Document document,
            IFlagManager refactoredService = null,
            FlagManager legacyService = null)
        {
            _useRefactored = OptimizationFlags.UseRefactoredClashZoneFlagServices;
            
            if (_useRefactored && refactoredService != null)
            {
                _refactoredService = refactoredService;
            }
            else
            {
                // Use legacy service
                _legacyService = legacyService ?? new FlagManager(document);
            }
        }
        
        /// <summary>
        /// Resets flags for deleted sleeves.
        /// Delegates to refactored service or legacy service based on feature flag.
        /// </summary>
        public int ResetFlagsForDeletedSleeves(
            List<ClashZone> clashZones, 
            List<string> categories,
            string refreshLogName = null)
        {
            if (_useRefactored && _refactoredService != null)
            {
                return _refactoredService.ResetFlagsForDeletedSleeves(clashZones, categories, refreshLogName);
            }
            else
            {
                // Legacy service signature is different - convert clashZones to clashZonesByCategory
                Dictionary<string, List<ClashZone>> clashZonesByCategory = null;
                if (clashZones != null && clashZones.Count > 0 && categories != null)
                {
                    clashZonesByCategory = new Dictionary<string, List<ClashZone>>(StringComparer.OrdinalIgnoreCase);
                    foreach (var category in categories)
                    {
                        var zonesForCategory = clashZones
                            .Where(cz => string.Equals(cz.MepElementCategory, category, StringComparison.OrdinalIgnoreCase))
                            .ToList();
                        if (zonesForCategory.Count > 0)
                        {
                            clashZonesByCategory[category] = zonesForCategory;
                        }
                    }
                }
                
                return _legacyService.ResetFlagsForDeletedSleeves(categories, clashZonesByCategory, refreshLogName);
            }
        }
        
        /// <summary>
        /// Resets instance IDs for deleted sleeves.
        /// Delegates to refactored service or legacy service based on feature flag.
        /// </summary>
        public int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null)
        {
            if (_useRefactored && _refactoredService != null)
            {
                return _refactoredService.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory, refreshLogName);
            }
            else
            {
                return _legacyService.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory, refreshLogName);
            }
        }
        
        /// <summary>
        /// Updates flags after sleeve placement.
        /// Delegates to refactored service or legacy service based on feature flag.
        /// </summary>
        public void UpdateFlagsAfterPlacement(
            List<(Guid clashZoneId, int sleeveInstanceId, bool isCluster)> placedSleeves)
        {
            if (_useRefactored && _refactoredService != null)
            {
                _refactoredService.UpdateFlagsAfterPlacement(placedSleeves);
            }
            else
            {
                // Legacy service doesn't have this method - convert to BatchUpdateFlagsForPlacement format
                // This requires loading ClashZone objects, which is complex
                // For now, log a warning - callers should use BatchUpdateFlagsForPlacement directly
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning("[FlagManagerAdapter] UpdateFlagsAfterPlacement not fully supported in legacy mode - use BatchUpdateFlagsForPlacement instead");
                }
            }
        }
        
        /// <summary>
        /// Batch update flags for placement.
        /// Delegates to refactored service or legacy service based on feature flag.
        /// </summary>
        public void BatchUpdateFlagsForPlacement(
            List<(ClashZone clashZone, int sleeveId)> clashZones, 
            bool isCluster, 
            string category, 
            string filterName = null)
        {
            if (_useRefactored && _refactoredService != null)
            {
                _refactoredService.BatchUpdateFlagsForPlacement(clashZones, isCluster, category, filterName);
            }
            else
            {
                _legacyService.BatchUpdateFlagsForPlacement(clashZones, isCluster, category, filterName);
            }
        }
        
        /// <summary>
        /// Get legacy FlagManager instance (for compatibility with services that require FlagManager directly).
        /// Returns null if using refactored service.
        /// </summary>
        public FlagManager? GetLegacyFlagManager()
        {
            return _legacyService;
        }
    }
}

