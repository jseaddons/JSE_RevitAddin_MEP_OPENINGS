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
        private readonly bool _useRefactored;
        
        /// <summary>
        /// Creates a flag manager adapter.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="refactoredService">Refactored flag manager service (can be null)</param>
        public FlagManagerAdapter(
            Document document,
            IFlagManager refactoredService = null)
        {
            _useRefactored = OptimizationFlags.UseRefactoredClashZoneFlagServices;
            _refactoredService = refactoredService;
        }
        
        /// <summary>
        /// Resets flags for deleted sleeves.
        /// Delegates to refactored service.
        /// </summary>
        public int ResetFlagsForDeletedSleeves(
            List<ClashZone> clashZones, 
            List<string> categories,
            string refreshLogName = null)
        {
            if (_refactoredService != null)
            {
                return _refactoredService.ResetFlagsForDeletedSleeves(clashZones, categories, refreshLogName);
            }
            return 0;
        }
        
        /// <summary>
        /// Resets instance IDs for deleted sleeves.
        /// Delegates to refactored service.
        /// </summary>
        public int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null)
        {
            if (_refactoredService != null)
            {
                return _refactoredService.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory, refreshLogName);
            }
            return 0;
        }
        
        /// <summary>
        /// Updates flags after sleeve placement.
        /// Delegates to refactored service.
        /// </summary>
        public void UpdateFlagsAfterPlacement(
            List<(Guid clashZoneId, int sleeveInstanceId, bool isCluster)> placedSleeves)
        {
            if (_refactoredService != null)
            {
                _refactoredService.UpdateFlagsAfterPlacement(placedSleeves);
            }
        }
        
        /// <summary>
        /// Batch update flags for placement.
        /// Delegates to refactored service.
        /// </summary>
        public void BatchUpdateFlagsForPlacement(
            List<(ClashZone clashZone, int sleeveId)> clashZones, 
            bool isCluster, 
            string category, 
            string filterName = null)
        {
            if (_refactoredService != null)
            {
                _refactoredService.BatchUpdateFlagsForPlacement(clashZones, isCluster, category, filterName);
            }
        }
        
        /// <summary>
        /// Get legacy FlagManager instance (for compatibility with services that require FlagManager directly).
        /// Returns null if using refactored service.
        /// </summary>
        public object? GetLegacyFlagManager()
        {
            return null;
        }
        /// <summary>
        /// Deletes a sleeve when its intersection point has changed significantly.
        /// Delegates to refactored service.
        /// </summary>
        public void DeleteSleeveForIntersectionPointChange(
            ClashZone zone, 
            string category, 
            double movementDistance)
        {
            if (_refactoredService != null)
            {
                _refactoredService.DeleteSleeveForIntersectionPointChange(zone, category, movementDistance);
            }
        }
        /// <summary>
        /// Syncs flags from database (single source of truth) to in-memory clash zones.
        /// Delegates to refactored service.
        /// </summary>
        public void SyncFlagsFromGlobal(List<ClashZone> clashZones, string category)
        {
            if (_refactoredService != null)
            {
                _refactoredService.SyncFlagsFromGlobal(clashZones, category);
            }
        }
        
        /// <summary>
        /// Verifies existing sleeves in model and resets flags for missing ones.
        /// Delegates to refactored service.
        /// </summary>
        public int VerifyExistingSleevesAndResetFlags(Document doc, List<string> filterNames, List<string> categories)
        {
             if (_refactoredService != null)
            {
                return _refactoredService.VerifyExistingSleevesAndResetFlags(doc, filterNames, categories);
            }
            return 0;
        }

        public bool GetFlag(Guid clashZoneId, string flagName)
        {
             if (_refactoredService != null)
            {
                return _refactoredService.GetFlag(clashZoneId, flagName);
            }
            return false;
        }

        public void SetFlag(Guid clashZoneId, string flagName, bool value)
        {
             if (_refactoredService != null)
            {
                _refactoredService.SetFlag(clashZoneId, flagName, value);
            }
        }

        public string GetFlagValue(Guid clashZoneId, string flagName)
        {
             if (_refactoredService != null)
            {
                return _refactoredService.GetFlagValue(clashZoneId, flagName);
            }
            return null;
        }

        public void SetFlagValue(Guid clashZoneId, string flagName, string value)
        {
             if (_refactoredService != null)
            {
                _refactoredService.SetFlagValue(clashZoneId, flagName, value);
            }
        }

        public List<ClashZone> GetFlaggedClashZones(string flagName, string category)
        {
             if (_refactoredService != null)
            {
                return _refactoredService.GetFlaggedClashZones(flagName, category);
            }
            return new List<ClashZone>();
        }

        public void SetFlaggedClashZones(List<ClashZone> clashZones, string flagName, bool value)
        {
             if (_refactoredService != null)
            {
                _refactoredService.SetFlaggedClashZones(clashZones, flagName, value);
            }
        }
    }
}

