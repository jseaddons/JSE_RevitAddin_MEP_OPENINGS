using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Team H: Factory for creating fully-wired clash zone services.
    /// 
    /// This factory handles dependency injection and coexistence between legacy and refactored services.
    /// SOLID: Factory Pattern - centralizes service creation logic.
    /// </summary>
    public static class ClashZoneServiceFactory
    {
        /// <summary>
        /// Create refactored clash zone service (SOLID-compliant).
        /// 
        /// ✅ WIRING: Creates all focused services and wires them together:
        /// - ClashZoneCleanupService (cleanup operations)
        /// - ClashZoneFilterService (filtering operations)
        /// - ClashZoneValidationService (validation operations)
        /// - DuctDamperFilterService (duct-damper filtering - SOLID: separate service)
        /// - DamperAutoDetectionService (auto-detection - SOLID: separate service)
        /// - ClashZoneDetectionService (detection operations - converts intersections to clash zones)
        /// - ClashZoneService (orchestrator)
        /// </summary>
        /// <param name="document">Revit document (required for detection service)</param>
        /// <param name="storage">Optional clash zone storage</param>
        /// <param name="logger">Optional logger (defaults to LoggerAdapter.Default)</param>
        /// <returns>IClashZoneService instance</returns>
        public static IClashZoneService CreateRefactored(
            Document document,
            ClashZoneStorage? storage = null,
            ILogger logger = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            logger = logger ?? LoggerAdapter.Default;
            
            // ✅ WIRING: Create focused services (SOLID: Single Responsibility)
            var cleanupService = new ClashZoneCleanupService(logger);
            var filterService = new ClashZoneFilterService(logger);
            var validationService = new ClashZoneValidationService(logger);
            
            // ✅ WIRING: Create duct-damper services (SOLID: separate concerns)
            var ductDamperFilterService = new DuctDamperFilterService(logger);
            var damperAutoDetectionService = new DamperAutoDetectionService(ductDamperFilterService, logger);
            
            // ✅ WIRING: Create detection service with duct-damper services injected (SOLID: Dependency Inversion)
            var detectionService = new ClashZoneDetectionService(
                document,
                ductDamperFilterService,
                damperAutoDetectionService,
                logger);
            
            // ✅ WIRING: Create main service (orchestrator)
            var service = new ClashZoneService(
                cleanupService,
                filterService,
                validationService,
                detectionService,
                storage,
                logger);
            
            return service;
        }
        
        // ⚠️ REMOVED: CreateAdapter method - legacy ClashZoneService has been moved to backup folder
        // The adapter is no longer needed since we always use refactored services
    }
}

