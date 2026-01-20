using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Team H: Factory for creating fully-wired flag management services.
    /// 
    /// This factory handles dependency injection and coexistence between legacy and refactored services.
    /// SOLID: Factory Pattern - centralizes service creation logic.
    /// </summary>
    public static class FlagManagerFactory
    {
        /* DISABLED: Refactored services not fully implemented
        /// <summary>
        /// Create refactored flag management services (SOLID-compliant).
        /// 
        /// ✅ WIRING: Creates all dependencies and wires them together:
        /// - SessionTrackerService (session tracking)
        /// - InstanceIdManagerService (instance ID management)
        /// - FlagManagerService (flag management)
        /// - IClashZoneRepository (database operations)
        /// - ISleeveCollector (Revit API operations)
        /// 
        /// ✅ TESTABILITY: All dependencies are injected via interfaces.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="repository">Optional repository (will be created if null)</param>
        /// <param name="sleeveCollector">Optional sleeve collector (will be created if null)</param>
        /// <param name="logger">Optional logger (defaults to LoggerAdapter.Default)</param>
        /// <returns>Tuple of (flagManager, instanceIdManager, sessionTracker)</returns>
        public static (IFlagManager flagManager, IInstanceIdManager instanceIdManager, ISessionTracker sessionTracker) CreateRefactored(
            Document document,
            IClashZoneRepository repository = null,
            ISleeveCollector sleeveCollector = null,
            ILogger logger = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            logger = logger ?? LoggerAdapter.Default;
            
            // ✅ TESTABILITY: Create repository if not provided (can be injected for testing)
            // ⚠️ NOTE: ClashZoneRepository requires SleeveDbContext which must be alive during operations.
            // For production, repositories are created per operation with a using statement.
            // For testing, inject a mock IClashZoneRepository that doesn't require a context.
            // This factory cannot create a production repository here because the context would be disposed.
            // Services will create repositories per operation when repository is null.
            if (repository == null)
            {
                // For production, we'll need to create repositories per operation
                // For now, pass null and let services handle it (they'll create per operation)
                // This is a known limitation - a repository factory would be ideal
                throw new InvalidOperationException("Repository must be provided. For production use, create repositories per operation. For testing, inject a mock IClashZoneRepository.");
            }
            
            // ✅ TESTABILITY: Create sleeve collector if not provided (can be injected for testing)
            if (sleeveCollector == null)
            {
                sleeveCollector = new RevitSleeveCollector();
            }
            
            var sessionTracker = new SessionTrackerService(logger);
            var instanceIdManager = new InstanceIdManagerService(document, repository, sleeveCollector, logger);
            var flagManager = new FlagManagerService(document, instanceIdManager, repository, sleeveCollector, logger);
            
            return (flagManager, instanceIdManager, sessionTracker);
        }
        */
        
        /// <summary>
        /// Create adapter for legacy FlagManager (coexistence).
        /// 
        /// This enables gradual migration from legacy FlagManager to refactored services.
        /// When UseRefactoredClashZoneFlagServices is false, this adapter delegates to legacy FlagManager.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="refactoredService">Optional refactored service (can be null)</param>
        /// <param name="legacyFlagManager">Legacy flag manager instance (can be null, will be created if needed)</param>
        /// <param name="logger">Optional logger (defaults to LoggerAdapter.Default)</param>
        /// <returns>IFlagManager adapter wrapping legacy service</returns>
        public static IFlagManager CreateAdapter(
            Document document,
            IFlagManager refactoredService = null,
            ILogger logger = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            // ✅ CRITICAL FIX: Create actual FlagManagerService if refactoredService is null
            // Previously, passing null caused the adapter to do nothing!
            if (refactoredService == null)
            {
                logger = logger ?? LoggerAdapter.Default;
                
                // Create required dependencies
                var dbContext = new SleeveDbContext(document);
                var repository = new ClashZoneRepository(dbContext);
                var sleeveCollector = new RevitSleeveCollector();
                
                // ✅ Use simple stub for IInstanceIdManager (not used by BatchUpdateFlagsForPlacement)
                var instanceIdManager = new StubInstanceIdManager();
                
                // Create actual FlagManagerService
                refactoredService = new FlagManagerService(document, instanceIdManager, repository, sleeveCollector, logger);
            }
            
            return new FlagManagerAdapter(document, refactoredService);
        }
    }

    /// <summary>
    /// Simple stub implementation of IInstanceIdManager.
    /// Only used by FlagManagerFactory when creating FlagManagerService.
    /// ✅ BatchUpdateFlagsForPlacement does NOT use IInstanceIdManager, so this stub is safe.
    /// </summary>
    internal class StubInstanceIdManager : IInstanceIdManager
    {
        public int ResetInstanceIdsForDeletedSleeves(
            System.Collections.Generic.List<string> categories,
            System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<Models.ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null)
        {
            // Stub - not used by BatchUpdateFlagsForPlacement
            return 0;
        }
    }
}

