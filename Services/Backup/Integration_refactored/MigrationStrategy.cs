using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement;
using JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Integration
{
    /// <summary>
    /// Team H: Migration strategy for gradually moving from legacy to refactored services.
    /// 
    /// This strategy enables gradual migration using a feature flag, allowing safe rollout
    /// of refactored services without breaking existing code.
    /// 
    /// SOLID: Strategy Pattern - enables switching between legacy and refactored implementations.
    /// </summary>
    public static class MigrationStrategy
    {
        /// <summary>
        /// Create services based on feature flag (gradual migration).
        /// 
        /// When UseRefactoredClashZoneFlagServices is true:
        /// - Uses refactored services (SOLID-compliant)
        /// 
        /// When UseRefactoredClashZoneFlagServices is false:
        /// - Uses adapters wrapping legacy services (coexistence)
        /// 
        /// ✅ SAFETY: Feature flag allows safe rollout and rollback.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="storage">Optional clash zone storage</param>
        /// <param name="legacyFlagManager">Optional legacy flag manager (for coexistence)</param>
        /// <param name="logger">Optional logger (defaults to LoggerAdapter.Default)</param>
        /// <param name="useRefactored">Override feature flag (if null, uses OptimizationFlags.UseRefactoredClashZoneFlagServices)</param>
        /// <returns>Tuple of (clashZoneService, flagManager, instanceIdManager, sessionTracker)</returns>
        public static (
            IClashZoneService clashZoneService, 
            IFlagManager flagManager,
            IInstanceIdManager instanceIdManager,
            ISessionTracker sessionTracker) CreateServices(
            Document document,
            ClashZoneStorage? storage = null,
            FlagManager? legacyFlagManager = null,
            ILogger logger = null,
            bool? useRefactored = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            logger = logger ?? LoggerAdapter.Default;
            
            // ✅ FEATURE FLAG: Check OptimizationFlags.UseRefactoredClashZoneFlagServices
            // Allow override via parameter for testing
            bool shouldUseRefactored = useRefactored ?? OptimizationFlags.UseRefactoredClashZoneFlagServices;
            
            if (shouldUseRefactored)
            {
                // ✅ WIRING: Use refactored services (SOLID-compliant)
                logger.Info("Creating refactored services (SOLID-compliant)", "MigrationStrategy");
                
                var clashZoneService = ClashZoneServiceFactory.CreateRefactored(document, storage, logger);
                
                // ✅ WIRING: Create refactored flag management services
                // ⚠️ CRITICAL: Services require repositories, but repositories must be created per operation
                // The refresh service creates repositories per operation with using statements
                // For now, we'll use the adapter pattern which wraps legacy services
                // TODO: Update services to accept repository factories or create repositories per operation internally
                var sleeveCollector = new RevitSleeveCollector();
                var sessionTracker = new SessionTrackerService(logger);
                
                // ✅ TEMPORARY SOLUTION: Use adapter pattern which wraps legacy FlagManager
                // The adapter will delegate to legacy FlagManager which creates repositories internally
                // This allows us to use refactored ClashZoneService while keeping flag management working
                // Once services support repository factories, we can switch to fully refactored services
                var tempLegacyFlagManager = new Services.FlagManager(document);
                var flagManager = FlagManagerFactory.CreateAdapter(document, null, tempLegacyFlagManager, logger);
                
                // ⚠️ NOTE: InstanceIdManager and SessionTracker are not available when using adapter
                // The adapter wraps legacy FlagManager which doesn't expose these interfaces
                // TODO: Update services to support repository factories so we can use fully refactored services
                IInstanceIdManager instanceIdManager = null;
                
                logger.Info("✅ Refactored ClashZoneService created, using FlagManagerAdapter for flag management", "MigrationStrategy");
                logger.Warning("⚠️ NOTE: Using adapter pattern - fully refactored services require repository factory support", "MigrationStrategy");
                
                return (clashZoneService, flagManager, instanceIdManager, sessionTracker);
            }
            else
            {
                // ✅ WIRING: Use legacy services via adapters (coexistence)
                logger.Info("Creating legacy service adapters (coexistence mode)", "MigrationStrategy");

                var concreteLegacyFlagManager = legacyFlagManager ?? new Services.FlagManager(document);

                var legacyClashZoneService = new Services.ClashZoneService(
                    storage,
                    msg => logger.Info(msg, "LegacyClashZoneService"),
                    concreteLegacyFlagManager,
                    guidManager: null);

                var clashZoneService = new ClashZoneServiceAdapter(legacyClashZoneService, logger);
                
                var flagManager = FlagManagerFactory.CreateAdapter(document, null, concreteLegacyFlagManager, logger);

                var instanceIdManager = new LegacyInstanceIdManagerAdapter(concreteLegacyFlagManager);
                var sessionTracker = new LegacySessionTrackerAdapter();
                
                return (clashZoneService, flagManager, instanceIdManager, sessionTracker);
            }
        }
        
        /// <summary>
        /// Create integration service with all dependencies wired.
        /// 
        /// This is a convenience method that creates all services and wires them into
        /// a ClashZoneFlagManagerIntegration instance.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="storage">Optional clash zone storage</param>
        /// <param name="legacyFlagManager">Optional legacy flag manager (for coexistence)</param>
        /// <param name="logger">Optional logger (defaults to LoggerAdapter.Default)</param>
        /// <param name="useRefactored">Override feature flag (if null, uses OptimizationFlags.UseRefactoredClashZoneFlagServices)</param>
        /// <returns>ClashZoneFlagManagerIntegration instance with all services wired</returns>
        public static ClashZoneFlagManagerIntegration CreateIntegration(
            Document document,
            ClashZoneStorage? storage = null,
            FlagManager? legacyFlagManager = null,
            ILogger logger = null,
            bool? useRefactored = null)
        {
            var (clashZoneService, flagManager, instanceIdManager, sessionTracker) = CreateServices(
                document,
                storage,
                legacyFlagManager,
                logger,
                useRefactored);
            
            // Ensure all required services are available
            if (flagManager == null)
                throw new InvalidOperationException("FlagManager is required but was null");
            
            if (instanceIdManager == null)
                throw new InvalidOperationException("InstanceIdManager is required but was null");
            
            if (sessionTracker == null)
                throw new InvalidOperationException("SessionTracker is required but was null");
            
            return new ClashZoneFlagManagerIntegration(
                clashZoneService,
                flagManager,
                instanceIdManager,
                sessionTracker,
                logger);
        }
    }
}
