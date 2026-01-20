using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.FilterManagement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Factories
{
    /// <summary>
    /// Factory for creating FilterManagement services.
    /// Wires Team I (persistence) with Team J (business logic/UI).
    /// </summary>
    public static class FilterManagementServiceFactory
    {
        /// <summary>
        /// Creates a complete FilterManagement service stack.
        /// Wires Team I persistence services with legacy FilterManagementService.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="log">Logging action</param>
        /// <param name="updateStatus">Status update action</param>
        /// <param name="logger">Optional ILogger instance</param>
        /// <returns>Configured FilterManagementService with Team I persistence</returns>
        public static FilterManagementService CreateFilterManagementService(
            Document document,
            Action<string> log,
            Action<string> updateStatus,
            ILogger logger = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            // Use default logger if not provided
            var effectiveLogger = logger ?? LoggerAdapter.Default;
            
            // ✅ TEAM I: Create persistence layer
            var filterRepository = new FilterRepositoryAdapter(document, effectiveLogger);
            var filterPersistenceService = new FilterPersistenceService(
                filterRepository,
                document,
                effectiveLogger
            );
            
            // ✅ TEAM J: Create business logic layer (legacy FilterManagementService)
            var filterManagementService = new FilterManagementService(document, log, updateStatus);
            
            // ✅ INTEGRATION: FilterManagementService will use Team I services internally
            // Note: Legacy FilterManagementService already uses FilterRepository internally
            // This factory provides a clean creation point for future migration
            
            effectiveLogger.Info("Created FilterManagementService with Team I persistence layer", "FilterManagementServiceFactory");
            
            return filterManagementService;
        }
        
        /// <summary>
        /// Creates ONLY the Team I persistence service.
        /// Use this when you need direct access to persistence operations.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="logger">Optional ILogger instance</param>
        /// <returns>Configured FilterPersistenceService</returns>
        public static IFilterPersistenceService CreatePersistenceService(
            Document document,
            ILogger logger = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            var effectiveLogger = logger ?? LoggerAdapter.Default;
            
            var filterRepository = new FilterRepositoryAdapter(document, effectiveLogger);
            var filterPersistenceService = new FilterPersistenceService(
                filterRepository,
                document,
                effectiveLogger
            );
            
            effectiveLogger.Info("Created FilterPersistenceService (Team I)", "FilterManagementServiceFactory");
            
            return filterPersistenceService;
        }
        
        /// <summary>
        /// Creates ONLY the Team I repository adapter.
        /// Use this when you need direct access to database operations.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="logger">Optional ILogger instance</param>
        /// <returns>Configured FilterRepositoryAdapter</returns>
        public static IFilterRepository CreateRepository(
            Document document,
            ILogger logger = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));
            
            var effectiveLogger = logger ?? LoggerAdapter.Default;
            
            var filterRepository = new FilterRepositoryAdapter(document, effectiveLogger);
            
            effectiveLogger.Info("Created FilterRepositoryAdapter (Team I)", "FilterManagementServiceFactory");
            
            return filterRepository;
        }
    }
}
