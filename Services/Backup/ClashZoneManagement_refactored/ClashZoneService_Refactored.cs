using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Team G: SOLID-compliant clash zone service (orchestrator).
    /// 
    /// This service orchestrates clash zone operations by delegating to focused services.
    /// SOLID: Single Responsibility - orchestration only.
    /// 
    /// ✅ PRESERVES ALL FUNCTIONALITY:
    /// - Cleanup operations (delegated to ClashZoneCleanupService)
    /// - Filtering operations (delegated to ClashZoneFilterService)
    /// - Validation operations (delegated to ClashZoneValidationService)
    /// </summary>
    public class ClashZoneService : IClashZoneService
    {
        private readonly IClashZoneCleanupService _cleanupService;
        private readonly IClashZoneFilterService _filterService;
        private readonly IClashZoneValidationService _validationService;
        private readonly IClashZoneDetectionService _detectionService;
        private readonly ClashZoneStorage? _storage;
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new clash zone service.
        /// </summary>
        /// <param name="cleanupService">Cleanup service for invalid clash zones</param>
        /// <param name="filterService">Filter service for selection-based filtering</param>
        /// <param name="validationService">Validation service for element existence checks</param>
        /// <param name="detectionService">Detection service for converting intersections to clash zones</param>
        /// <param name="storage">Optional clash zone storage</param>
        /// <param name="logger">Optional logger for tracking operations</param>
        public ClashZoneService(
            IClashZoneCleanupService cleanupService,
            IClashZoneFilterService filterService,
            IClashZoneValidationService validationService,
            IClashZoneDetectionService detectionService,
            ClashZoneStorage? storage = null,
            ILogger logger = null)
        {
            _cleanupService = cleanupService ?? throw new ArgumentNullException(nameof(cleanupService));
            _filterService = filterService ?? throw new ArgumentNullException(nameof(filterService));
            _validationService = validationService ?? throw new ArgumentNullException(nameof(validationService));
            _detectionService = detectionService ?? throw new ArgumentNullException(nameof(detectionService));
            _storage = storage;
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Cleanup invalid clash zones (removes null elements and duplicates).
        /// </summary>
        public int CleanupInvalidClashZones(Document document)
        {
            if (_storage?.ClashZones == null)
            {
                _logger.Info("No clash zones to clean up", "ClashZoneService");
                return 0;
            }
            
            // ✅ DELEGATE: Use focused cleanup service
            var removed = _cleanupService.CleanupInvalidClashZones(_storage.ClashZones, document);
            
            _logger.Info($"Cleaned up {removed} invalid clash zones", "ClashZoneService");
            
            return removed;
        }
        
        /// <summary>
        /// Filter clash zones by current selection parameters.
        /// </summary>
        public List<ClashZone> FilterClashZonesByCurrentSelection(
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document)
        {
            if (_storage?.ClashZones == null)
            {
                _logger.Info("No clash zones to filter", "ClashZoneService");
                return new List<ClashZone>();
            }
            
            // ✅ DELEGATE: Use focused filter service
            var filteredZones = _filterService.FilterClashZonesByCurrentSelection(
                _storage.ClashZones,
                selectedReferenceFiles,
                currentClearanceSettings,
                currentPrefix,
                document);
            
            _logger.Info($"Filtered {_storage.ClashZones.Count} → {filteredZones.Count} clash zones", "ClashZoneService");
            
            return filteredZones;
        }
        
        /// <summary>
        /// Detect new clash zones from intersections.
        /// Delegates to ClashZoneDetectionService to convert intersections to ClashZone objects.
        /// </summary>
        public List<ClashZone> DetectNewClashZones(
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document,
            Dictionary<string, double> clearanceSettings = null,
            List<string> selectedCategories = null)
        {
            if (_storage == null)
            {
                _logger.Warning("ClashZoneStorage is null, cannot detect clash zones", "ClashZoneService");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning("[ClashZoneService] ❌ ERROR: ClashZoneStorage is null - detection will return empty list");
                }
                return new List<ClashZone>();
            }
            
            // ✅ DIAGNOSTIC: Log storage state before detection
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[ClashZoneService] DetectNewClashZones called: {currentIntersections?.Count ?? 0} intersections, storage has {_storage.ClashZones?.Count ?? 0} existing zones");
            }
            
            // ✅ DELEGATE: Use focused detection service
            var result = _detectionService.DetectNewClashZones(
                currentIntersections,
                document,
                _storage,
                clearanceSettings,
                selectedCategories);
            
            // ✅ DIAGNOSTIC: Log result after detection
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[ClashZoneService] DetectNewClashZones returned: {result?.Count ?? 0} new clash zones, storage now has {_storage.ClashZones?.Count ?? 0} total zones");
            }
            
            return result;
        }
    }
}
