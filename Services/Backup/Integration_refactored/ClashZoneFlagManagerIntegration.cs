using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Integration
{
    /// <summary>
    /// Team H: Integration service that wires ClashZoneService and FlagManager together.
    /// 
    /// This service provides a unified interface for operations that require both
    /// clash zone management and flag management services.
    /// 
    /// SOLID: Integration Pattern - coordinates multiple services without tight coupling.
    /// </summary>
    public class ClashZoneFlagManagerIntegration
    {
        private readonly IClashZoneService _clashZoneService;
        private readonly IFlagManager _flagManager;
        private readonly IInstanceIdManager _instanceIdManager;
        private readonly ISessionTracker _sessionTracker;
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new integration service.
        /// </summary>
        /// <param name="clashZoneService">Clash zone service</param>
        /// <param name="flagManager">Flag manager service</param>
        /// <param name="instanceIdManager">Instance ID manager service</param>
        /// <param name="sessionTracker">Session tracker service</param>
        /// <param name="logger">Optional logger (defaults to LoggerAdapter.Default)</param>
        public ClashZoneFlagManagerIntegration(
            IClashZoneService clashZoneService,
            IFlagManager flagManager,
            IInstanceIdManager instanceIdManager,
            ISessionTracker sessionTracker,
            ILogger logger = null)
        {
            _clashZoneService = clashZoneService ?? throw new ArgumentNullException(nameof(clashZoneService));
            _flagManager = flagManager ?? throw new ArgumentNullException(nameof(flagManager));
            _instanceIdManager = instanceIdManager ?? throw new ArgumentNullException(nameof(instanceIdManager));
            _sessionTracker = sessionTracker ?? throw new ArgumentNullException(nameof(sessionTracker));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Execute refresh with integrated services.
        /// 
        /// This method coordinates cleanup, flag reset, and instance ID reset operations.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="categories">List of MEP categories to process</param>
        /// <param name="clashZonesByCategory">Dictionary of clash zones by category</param>
        public void ExecuteRefresh(
            Document document,
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            if (categories == null || categories.Count == 0)
            {
                _logger.Info("No categories to process", "ClashZoneFlagManagerIntegration");
                return;
            }
            
            try
            {
                // ✅ STEP 1: Cleanup invalid clash zones
                _logger.Info("Starting refresh with integrated services", "ClashZoneFlagManagerIntegration");
                var cleanupCount = _clashZoneService.CleanupInvalidClashZones(document);
                _logger.Info($"Cleaned up {cleanupCount} invalid clash zones", "ClashZoneFlagManagerIntegration");
                
                // ✅ STEP 2: Reset flags for deleted sleeves
                int totalResetCount = 0;
                foreach (var category in categories)
                {
                    var clashZones = (clashZonesByCategory != null && clashZonesByCategory.ContainsKey(category)) 
                        ? clashZonesByCategory[category] 
                        : new List<ClashZone>();
                    var resetCount = _flagManager.ResetFlagsForDeletedSleeves(clashZones, new List<string> { category });
                    totalResetCount += resetCount;
                    _logger.Info($"Reset flags for {resetCount} deleted sleeves in category '{category}'", "ClashZoneFlagManagerIntegration");
                }
                
                // ✅ STEP 3: Reset instance IDs for deleted sleeves
                var instanceResetCount = _instanceIdManager.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory);
                _logger.Info($"Reset instance IDs for {instanceResetCount} deleted sleeves", "ClashZoneFlagManagerIntegration");
                
                _logger.Info($"Refresh complete: {cleanupCount} cleaned, {totalResetCount} flags reset, {instanceResetCount} instance IDs reset", "ClashZoneFlagManagerIntegration");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error during refresh: {ex.Message}", ex, "ClashZoneFlagManagerIntegration");
                throw; // Re-throw to let caller handle
            }
        }
        
        /// <summary>
        /// Clear recently placed cluster sleeves (session cleanup).
        /// </summary>
        public void ClearSession()
        {
            _sessionTracker.ClearRecentlyPlacedClusterSleeves();
            _logger.Info("Session cleared", "ClashZoneFlagManagerIntegration");
        }
    }
}

