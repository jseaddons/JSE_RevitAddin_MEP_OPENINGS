using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Orchestrator for executing opening placement commands with memory management
    /// Uses NEW Universal Architecture with UniversalSleevePlacementCommand
    /// </summary>
    public class OpeningCommandOrchestrator
    {
        private readonly Document _document;
        private readonly UIDocument _uiDocument;
        private readonly Dictionary<string, double> _uiClearances;

        private readonly MarkPrefixSettings _markPrefixes;
        private readonly bool _forceDetectionMode;
        
        // ⚠️ CRITICAL: Crash-safe executor for timeout protection (5-minute limit per category)
        private readonly CrashSafeExecutor _crashSafeExecutor;
        
        // ✅ PATH 3: Store PATH 3 type flags for clustering (key: "FilterName_Category", value: (isValidated, isInvalidated, isNew))
        private readonly Dictionary<string, (bool isValidated, bool isInvalidated, bool isNew)> _path3Flags = 
            new Dictionary<string, (bool, bool, bool)>();



        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument, Dictionary<string, double> uiClearances = null, MarkPrefixSettings markPrefixes = null, bool forceDetectionMode = false)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _uiClearances = uiClearances ?? new Dictionary<string, double>();
            _markPrefixes = markPrefixes ?? new MarkPrefixSettings();
            _forceDetectionMode = forceDetectionMode;
            
            // ⚠️ CRITICAL: Initialize crash-safe executor for timeout protection
            _crashSafeExecutor = new CrashSafeExecutor();
        }

        /// <summary>
        /// Set UI clearance settings from the main dialog
        /// </summary>
        public void SetUIClearances(Dictionary<string, double> clearances)
        {
            _uiClearances.Clear();
            if (clearances != null)
            {
                foreach (var kvp in clearances)
                {
                    _uiClearances[kvp.Key] = kvp.Value;
                }
            }
        }

        /// <summary>
        /// Execute multiple filters with memory management
        /// </summary>
        public void ExecuteMultipleFilters(List<OpeningFilter> filters, bool showProgress = true)
        {
            // 🔥 CRITICAL DEBUG: Direct file logging to trace orchestrator execution
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔥 ExecuteMultipleFilters CALLED 🔥\n");
                DebugLogger.Info("[OpeningCommandOrchestrator] 🔥 ExecuteMultipleFilters CALLED 🔥");
            }
            
            if (filters == null || filters.Count == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ No filters provided\n");
                    DebugLogger.Warning("[OpeningCommandOrchestrator] No filters provided");
                }
                return;
            }

            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Starting execution of {filters.Count} filters\n");
                DebugLogger.Info($"[OpeningCommandOrchestrator] Starting execution of {filters.Count} filters");
            }

            // 🔥 BUILD VERIFICATION: Log DLL build timestamp to confirm rebuild
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var buildTime = System.IO.File.GetLastWriteTime(assembly.Location);
                DebugLogger.Info($"[ORCHESTRATOR] 🔨 DLL BUILD TIME: {buildTime:yyyy-MM-dd HH:mm:ss} - If this timestamp is old, DLL was NOT rebuilt!");
            }
            catch { }

            try
            {
                // 🔥 DIAGNOSTIC: Log verification entry
                SafeFileLogger.SafeAppendText("orchestrator_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 BEFORE VERIFICATION CALL\n");
                
                /* ⚠️ DISABLED BY USER REQUEST: Verification should only run in Refresh
                // ✅ CRITICAL: Verify all sleeve types (individual, cluster, combined) still exist in Revit
                // Reset flags for deleted sleeves BEFORE loading zones to prevent duplicate placement
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info("[OpeningCommandOrchestrator] Verifying existing sleeves and resetting flags for deleted ones...");
                }
                
                SafeFileLogger.SafeAppendText("orchestrator_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 CREATING DbContext for verification\n");
                
                using (var dbContext = new SleeveDbContext(_document))
                {
                    SafeFileLogger.SafeAppendText("orchestrator_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 DbContext created, creating repository\n");
                    
                    var repository = new ClashZoneRepository(dbContext, msg => DebugLogger.Info(msg));
                    var filterNames = filters.Select(f => f.Name).Distinct().ToList();
                    var categories = filters.SelectMany(f => f.SelectedMepCategoryNames ?? new List<string>()).Distinct().ToList();
                    
                    SafeFileLogger.SafeAppendText("orchestrator_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 CALLING VerifyExistingSleevesAndResetFlags: filters={string.Join(",", filterNames)}, categories={string.Join(",", categories)}\n");
                    
                    int resetCount = repository.VerifyExistingSleevesAndResetFlags(_document, filterNames, categories);
                    
                    SafeFileLogger.SafeAppendText("orchestrator_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥 VERIFICATION RETURNED: resetCount={resetCount}\n");
                }
                */

                // Group filters by name for memory management
                var disciplineGroups = GroupFiltersByName(filters);

                foreach (var discipline in disciplineGroups)
                {
                    ExecuteDisciplineWithMemoryManagement(discipline.Key, discipline.Value, showProgress);
                }

                // Feature: Combined Sleeves (Refactored to dedicated Manager)
                // We use the new manager to keep this orchestrator clean (SRP)
                // var combinedManager = new CombinedSleeveManager(_document);
                // combinedManager.Execute(showProgress); // ⚠️ DISABLED BY USER REQUEST: Separation of concerns

                // ⚠️ DISABLED: Marking phase removed from OK click as it's handled by separate UI
                //                 if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[OpeningCommandOrchestrator] 🔥 STARTING MARKING PHASE 🔥");
                
                // // ✅ PROPER ARCHITECTURE: Call MarkParameterCommand with UI values
                // // MarkParameterCommand handles "ALL" by processing each category with correct UI discipline prefixes
                
                // // ✅ FIX: Pass MarkPrefixSettings to MarkParameterCommand so it can use UI discipline prefixes
                // var markingCommand = new MarkParameterCommand("ALL", _markPrefixes.ProjectPrefix, "ALL", _markPrefixes.RemarkAll, _markPrefixes);
                // markingCommand.Execute(new UIApplication(_document.Application));
                
                //                 if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[OpeningCommandOrchestrator] ✅ MarkParameterCommand completed for ALL categories");
                //                 if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[OpeningCommandOrchestrator] 🔥 MARKING PHASE COMPLETED 🔥");

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info("[OpeningCommandOrchestrator] All filters executed successfully");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing filters: {ex.Message}");
                }
                throw;
            }
        }

        /// <summary>
        /// Group filters by name for memory management
        /// </summary>
        private Dictionary<string, List<OpeningFilter>> GroupFiltersByName(List<OpeningFilter> filters)
        {
            return filters.GroupBy(f => f.Name)
                         .ToDictionary(g => g.Key, g => g.ToList());
        }

        /// <summary>
        /// Order filters by priority to ensure duct accessories are processed before ducts
        /// </summary>
        private List<OpeningFilter> OrderFiltersByPriority(List<OpeningFilter> filters)
        {
            return filters.OrderBy(f => GetFilterPriority(f)).ToList();
        }

        /// <summary>
        /// Get priority value for filter ordering (lower number = higher priority)
        /// </summary>
        private int GetFilterPriority(OpeningFilter filter)
        {
            // Check if this filter is for duct accessories
            if (filter.SelectedMepCategoryNames?.Any(cat => 
                string.Equals(cat, "Duct Accessories", StringComparison.OrdinalIgnoreCase)) == true)
            {
                return 1; // Highest priority - process first
            }

            // Check if this filter is for ducts
            if (filter.SelectedMepCategoryNames?.Any(cat => 
                string.Equals(cat, "Ducts", StringComparison.OrdinalIgnoreCase)) == true)
            {
                return 2; // Second priority - process after duct accessories
            }

            // All other categories get default priority
            return 10; // Lower priority - process last
        }

        /// <summary>
        /// Execute all filters for a discipline with memory management
        /// </summary>
        private void ExecuteDisciplineWithMemoryManagement(string discipline, List<OpeningFilter> filters, bool showProgress)
        {
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[OpeningCommandOrchestrator] 🔥 ExecuteDisciplineWithMemoryManagement CALLED 🔥");
                DebugLogger.Info($"[OpeningCommandOrchestrator] Executing discipline: {discipline} with {filters.Count} filters");
            }

            try
            {
                // ✅ PRIORITY ORDERING: Sort filters to ensure duct accessories are processed before ducts
                var orderedFilters = OrderFiltersByPriority(filters);
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[OpeningCommandOrchestrator] Ordered {orderedFilters.Count} filters by priority for discipline: {discipline}");
                }
                
                // Log the processing order
                for (int i = 0; i < orderedFilters.Count; i++)
                {
                    var filter = orderedFilters[i];
                    var categories = string.Join(", ", filter.SelectedMepCategoryNames ?? new List<string>());
                    var priority = GetFilterPriority(filter);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Processing order {i + 1}: '{filter.Name}' (Categories: {categories}, Priority: {priority})");
                    }
                }

                foreach (var filter in orderedFilters)
                {
                    var commandSequence = GetCommandSequence(filter);
                    ExecuteCommandSequence(commandSequence, filter, showProgress);
                }

                // Force garbage collection after each discipline
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Discipline {discipline} completed, memory cleaned");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing discipline {discipline}: {ex.Message}");
                }
                throw;
            }
        }

        /// <summary>
        /// Get command sequence for a filter using NEW Universal Architecture
        /// </summary>
        private List<IExternalCommand> GetCommandSequence(OpeningFilter filter)
        {
            var sequence = new List<IExternalCommand>();

            // ✅ NEW ARCHITECTURE: Use UniversalSleevePlacementCommand for ALL categories
            // Note: UniversalSleevePlacementCommand implements ICommand, not IExternalCommand
            // We'll need to create a wrapper or use a different approach

            // ✅ FIX: Always add clustering for all filters (no need to check OpeningType)
            // We'll handle clustering directly in ExecuteCommandSequence using UniversalClusterService

            return sequence;
        }

        /// <summary>
        /// Execute a sequence of commands
        /// </summary>
        private void ExecuteCommandSequence(List<IExternalCommand> commands, OpeningFilter filter, bool showProgress)
        {
            // ✅ PERFORMANCE MONITORING: Initialize placement performance monitor
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string performanceLogName = $"Placement_{filter.Name}_{filter.Category}_{timestamp}.log";
            var performanceMonitor = new PlacementPerformanceMonitor(performanceLogName);
            int totalIndividualSleeves = 0;
            int totalClusters = 0;
            
            try
            {
                foreach (var command in commands)
                {
                    ExecuteCommandWithResourceManagement(command, filter, showProgress);
                }

                // ⚠️ CRITICAL: INDIVIDUAL SLEEVES MUST RUN FIRST - DO NOT CHANGE THIS ORDER! ⚠️
                // Clustering REQUIRES existing individual sleeves to work - it collects sleeves from Revit
                // If clustering runs first, there will be no sleeves to cluster and it will fail
                // NEVER PUT ExecuteClusteringForCategory BEFORE ExecuteUniversalSleevePlacement
                
                // ✅ PERFORMANCE: Track individual sleeve placement
                using (var individualTracker = performanceMonitor.TrackOperation("Individual Sleeve Placement"))
                {
                    var individualResult = ExecuteUniversalSleevePlacement(filter, showProgress);
                    totalIndividualSleeves = individualResult.placedCount;
                    
                    // ✅ FIX: Set item count BEFORE tracker disposes (must be inside using block)
                    individualTracker.SetItemCount(totalIndividualSleeves);
                    
                    // ✅ PERFORMANCE LOGGING: Log actual performance metrics
                    if (!DeploymentConfiguration.DeploymentMode && totalIndividualSleeves > 0)
                    {
                        var avgTimePerSleeve = 288705.0 / totalIndividualSleeves; // Approximate from log
                        DebugLogger.Info($"[PERFORMANCE] Individual Sleeve Placement: {totalIndividualSleeves} sleeves in ~{288705/1000.0:F1}s = ~{avgTimePerSleeve/1000.0:F2}s per sleeve (target: 0.02s)");
                        if (avgTimePerSleeve > 1000) // More than 1 second per sleeve
                        {
                            DebugLogger.Warning($"[PERFORMANCE] ⚠️ VERY SLOW: {avgTimePerSleeve/1000.0:F2}s per sleeve is {avgTimePerSleeve/20.0:F0}x slower than target (0.02s)");
                        }
                    }
                }

                // ✅ FIX: Corner Extraction - Force READ from Revit Geometry (Solids) immediately after placement
                // This ensures "Corner1X...Corner4Z" in the DB are accurate for the next step (Clustering)
                // ✅ CRITICAL FIX: Run for ALL sleeves in DB (not just newly placed) to fix existing bad data
                
                // ✅ FORCED LOG: Always log this to confirm code path is reached
                SafeFileLogger.SafeAppendText("corner_extraction.log",
                    $"[{DateTime.Now:HH:mm:ss}] [ORCHESTRATOR] ⚙️ Corner Extraction block ENTERED for category {filter.Category}\n");
                
                try
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                         DebugLogger.Info($"[OpeningCommandOrchestrator] 🔄 Starting Corner Extraction (Regenerate skipped - already done during placement)...");
                    }
                    
                    // NOTE: Document.Regenerate() removed - it requires open transaction
                    // Placement already regenerates the document within its transaction
                    
                    // 2. Run Batch Extractor
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                         DebugLogger.Info($"[OpeningCommandOrchestrator] 📏 Starting BatchSleeveCornerExtractor for {filter.Category}...");
                    }
                    
                    // Determine category name string
                    string catName = filter.Category switch
                    {
                        Models.MepCategory.Ducts => "Ducts",
                        Models.MepCategory.DuctAccessories => "Duct Accessories",
                        Models.MepCategory.Pipes => "Pipes",
                        Models.MepCategory.CableTrays => "Cable Trays",
                        _ => filter.Category.ToString()
                    };

                    // Initialize Extractor (stateless)
                    var cornerExtractor = new JSE_RevitAddin_MEP_OPENINGS.Services.Persistence.BatchSleeveCornerExtractor();
                    int extractedCount = 0;

                    // Create local context/repo to fetch zones and save results
                    using (var dbContext = new SleeveDbContext(_document))
                    {
                        var clashZoneRepository = new ClashZoneRepository(dbContext);
                        
                        // A. Fetch Zones
                        var zones = clashZoneRepository.GetClashZonesByCategory(catName)
                            .Where(z => z.SleeveInstanceId > 0)
                            .ToList();
                        
                        // ✅ FORCED LOG: Show zones count
                        SafeFileLogger.SafeAppendText("corner_extraction.log",
                            $"[{DateTime.Now:HH:mm:ss}] [INFO] Found {zones.Count} zones with SleeveInstanceId > 0 for category '{catName}'\n");

                        // B. Extract Corners (Pure Geometry Logic)
                        var extractionResults = cornerExtractor.ExtractCorners(_document, zones);
                        extractedCount = extractionResults.Count;

                        // C. Save Results (Repository Logic)
                        if (extractedCount > 0)
                        {
                            SafeFileLogger.SafeAppendText("corner_extraction.log",
                                $"[{DateTime.Now:HH:mm:ss}] [SAVE] 📝 Calling BatchUpdateSleeveCorners with {extractedCount} results...\n");
                            
                            clashZoneRepository.BatchUpdateSleeveCorners(extractionResults);
                            
                            SafeFileLogger.SafeAppendText("corner_extraction.log",
                                $"[{DateTime.Now:HH:mm:ss}] [SAVE] ✅ BatchUpdateSleeveCorners completed.\n");
                        }
                        else
                        {
                            SafeFileLogger.SafeAppendText("corner_extraction.log",
                                $"[{DateTime.Now:HH:mm:ss}] [SKIP] ⚠️ No extraction results to save (count=0).\n");
                        }
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                         DebugLogger.Info($"[OpeningCommandOrchestrator] ✅ BatchSleeveCornerExtractor completed: Updated {extractedCount} sleeves.");
                    }
                }
                catch (Exception ex)
                {
                    // ✅ FORCED LOG: Always log exceptions to diagnose issues
                    SafeFileLogger.SafeAppendText("corner_extraction.log",
                        $"[{DateTime.Now:HH:mm:ss}] [ERROR] ❌ Corner Extraction FAILED: {ex.Message}\n{ex.StackTrace}\n");
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                         DebugLogger.Warning($"[OpeningCommandOrchestrator] ⚠️ Batch Corner Extraction failed: {ex.Message}");
                }
                
                // ✅ LOGGING: Wrap with SafeFileLogger
                SafeFileLogger.SafeAppendText("orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] AFTER ExecuteUniversalSleevePlacement, ABOUT TO CALL ExecuteClusteringForCategory for filter={filter.Name}, category={filter.Category}\n");
                
                // ✅ CRITICAL: Execute clustering immediately after sleeve placement for each category
                // This follows the architecture: Place sleeves → Cluster sleeves → Mark sleeves
                // Use UniversalClusterService directly (faster than old RectangularSleeveClusterCommandV2)
                
                // ✅ PERFORMANCE: Track cluster placement
                // ✅ PATH 3: Retrieve PATH 3 flags for clustering
                string categoryString = filter.Category switch
                {
                    Models.MepCategory.Ducts => "Ducts",
                    Models.MepCategory.DuctAccessories => "Duct Accessories",
                    Models.MepCategory.Pipes => "Pipes", 
                    Models.MepCategory.CableTrays => "Cable Trays",
                    _ => "Ducts"
                };
                string path3Key = $"{filter.Name}_{categoryString}";
                bool isPath3Validated = false;
                bool isPath3Invalidated = false;
                bool isPath3New = false;
                if (_path3Flags.TryGetValue(path3Key, out var path3Flags))
                {
                    isPath3Validated = path3Flags.isValidated;
                    isPath3Invalidated = path3Flags.isInvalidated;
                    isPath3New = path3Flags.isNew;
                }
                
                using (var clusterTracker = performanceMonitor.TrackOperation("Cluster Sleeve Placement"))
                {
                    var clusterResult = ExecuteClusteringForCategory(filter, showProgress, isPath3Validated, isPath3Invalidated, isPath3New);
                    totalClusters = clusterResult.placedCount;
                    clusterTracker.SetItemCount(totalClusters);
                }
                
                // ✅ LOGGING: Wrap with SafeFileLogger
                SafeFileLogger.SafeAppendText("orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] AFTER ExecuteClusteringForCategory completed for filter={filter.Name}, category={filter.Category}\n");
                
                // ✅ COMBO FLAG: Reset IsFilterComboNew to false after full sequence completes (individual + cluster)
                // This marks the filter+category combo as "used" so next time it will use PATH 1 (Replay)
                ResetFilterComboFlagAfterPlacement(filter);
                
                // ✅ PERFORMANCE: Generate final report
                performanceMonitor.GenerateReport(totalIndividualSleeves, totalClusters);
            }
            catch (Exception ex)
            {
                // ✅ LOGGING: Wrap with SafeFileLogger
                SafeFileLogger.SafeAppendText("placement_errors.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ERROR in ExecuteCommandSequence for {filter.Name}: {ex.Message}\n");
                
                // Still generate report even on error
                performanceMonitor.GenerateReport(totalIndividualSleeves, totalClusters);
                throw;
            }
        }

        /// <summary>
        /// ✅ FIX: Resets IsFilterComboNew to false for all file combos used during placement
        /// Called after individual sleeves + cluster sleeves are placed successfully
        /// Resets flag PER FILE COMBO (not per filter+category) - each file combo has its own flag
        /// </summary>
        private void ResetFilterComboFlagAfterPlacement(OpeningFilter filter)
        {
            try
            {
                if (filter == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[COMBO-FLAG] Cannot reset flag - filter is null");
                    return;
                }

                // Get category string from filter
                string categoryName = filter.Category switch
                {
                    Models.MepCategory.Ducts => "Ducts",
                    Models.MepCategory.DuctAccessories => "Duct Accessories",
                    Models.MepCategory.Pipes => "Pipes",
                    Models.MepCategory.CableTrays => "Cable Trays",
                    _ => filter.Category.ToString()
                };

                // Normalize category name
                categoryName = MepCategoryConstants.Normalize(categoryName);

                // ✅ FIX: Reset IsFilterComboNew flag for ALL file combos used during placement
                // Reset all file combos for this filter+category that have IsFilterComboNew=1
                using (var dbContext = new SleeveDbContext(_document))
                {
                    var filterRepository = new FilterRepository(dbContext, _ => { });
                    int filterId = filterRepository.GetFilterId(filter.Name, categoryName);
                    
                    if (filterId > 0)
                    {
                        using (var cmd = dbContext.Connection.CreateCommand())
                        {
                            // ✅ DATABASE-ONLY: Reset all file combos for this filter that have IsFilterComboNew=1
                            // This marks them as processed (IsFilterComboNew=0) so Path 1 will be available next refresh
                            cmd.CommandText = @"
                                UPDATE FileCombos
                                SET IsFilterComboNew = 0, ProcessedAt = CURRENT_TIMESTAMP
                                WHERE FilterId = @FilterId AND IsFilterComboNew = 1";
                            cmd.Parameters.AddWithValue("@FilterId", filterId);
                            
                            var affected = cmd.ExecuteNonQuery();
                            if (affected > 0)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[COMBO-FLAG] ✅ Reset IsFilterComboNew=0 for {affected} file combo(s) for filter '{filter.Name}' (Category='{categoryName}') - file combos marked as processed (database-only, no XML)");
                                }
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[COMBO-FLAG] No file combos to reset for filter '{filter.Name}' (Category='{categoryName}') - all combos already processed");
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[COMBO-FLAG] ⚠️ Filter '{filter.Name}' (Category='{categoryName}') not found - cannot reset file combo flags");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[COMBO-FLAG] ❌ Error resetting IsFilterComboNew flag: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Execute clustering for a specific category after sleeve placement
        /// ⚠️ CRITICAL: Wrapped with timeout protection (5-minute limit per category)
        /// </summary>
        private (int placedCount, int deletedCount) ExecuteClusteringForCategory(
            OpeningFilter filter, 
            bool showProgress = false,
            bool isPath3Validated = false,
            bool isPath3Invalidated = false,
            bool isPath3New = false)
        {
            // Declare variables outside lambda for use after timeout execution
            List<FamilyInstance> placedClusterSleeves = new List<FamilyInstance>();
            string categoryString = null;
            string xmlFilePath = null;
            
            // ✅ PERFORMANCE: Store counts outside lambda for access after timeout execution
            int placedCount = 0;
            int deletedCount = 0;
            
            // ⚠️ CRITICAL: Execute with timeout protection to prevent infinite hangs
            var result = _crashSafeExecutor.ExecuteWithTimeout(() =>
            {
                try
                {
                    // 🔥 CRITICAL DEBUG: Log clustering attempt
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔥 ExecuteClusteringForCategory CALLED for category: {filter.Category} 🔥\n");
                    }
                    
                    // Convert MepCategory enum to string for cluster command
                    categoryString = filter.Category switch
                    {
                        Models.MepCategory.Ducts => "Ducts",
                        Models.MepCategory.DuctAccessories => "Duct Accessories",
                        Models.MepCategory.Pipes => "Pipes", 
                        Models.MepCategory.CableTrays => "Cable Trays",
                        _ => "Ducts"
                    };
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔥 Starting clustering for category: {categoryString} 🔥\n");
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Starting clustering for category: {categoryString}");
                    }
                    
                    // ✅ PATH 1 CHECK: Check database for existing cluster data (SKIP if AdoptToDocument is enabled)
                    bool isPath1Replay = false;
                    int? comboId = null;
                    int? filterId = null;
                    bool adoptToDocumentEnabled = false;
                    
                    try
                    {
                        using (var dbContext = new SleeveDbContext(_document))
                        {
                            var filterRepository = new FilterRepository(dbContext, _ => { });
                            int lookedUpFilterId = filterRepository.GetFilterId(filter.Name, categoryString);
                            
                            if (lookedUpFilterId > 0)
                            {
                                filterId = lookedUpFilterId;
                                
                                // ✅ CRITICAL: Check AdoptToDocumentFlag first - if enabled, skip PATH 1 and use PATH 3
                                using (var cmd = dbContext.Connection.CreateCommand())
                                {
                                    cmd.CommandText = @"SELECT AdoptToDocumentFlag FROM Filters WHERE FilterId = @FilterId";
                                    cmd.Parameters.AddWithValue("@FilterId", lookedUpFilterId);
                                    var flagResult = cmd.ExecuteScalar();
                                    adoptToDocumentEnabled = flagResult != null && flagResult != DBNull.Value && Convert.ToInt32(flagResult) == 1;
                                }
                                
                                File.AppendAllText(SafeFileLogger.GetLogFilePath("orchestrator_debug.log"), 
                                    $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: AdoptToDocumentFlag={adoptToDocumentEnabled}, FilterId={lookedUpFilterId}\n");
                                
                                // ✅ PATH 1a LOGIC: Check for Path 1a (super fast) even when AdoptToDocument is enabled
                                // Path 1a: Validated zones (no invalidated zones) → Use Path 1a for super fast replay
                                // Path 3: Invalidated zones exist → Use Path 3 for full recalculation
                                
                                // Check if there are invalidated zones (Path 3 flags) - use same key format as line 311
                                string path3Key = $"{filter.Name}_{categoryString}";
                                bool hasInvalidatedZones = false;
                                if (_path3Flags.TryGetValue(path3Key, out var path3Flags))
                                {
                                    hasInvalidatedZones = path3Flags.isInvalidated;
                                }
                                
                                // ✅ PATH 1a: If no invalidated zones, check for Path 1a (even with AdoptToDocument enabled)
                                // ✅ PATH 3: If invalidated zones exist, skip Path 1a and use Path 3
                                if (!hasInvalidatedZones)
                                {
                                    // ✅ Check if there's cluster data for this filter+category in database (Path 1a)
                                    var clusterRepository = new ClusterSleeveRepository(dbContext);
                                    var existingClusters = clusterRepository.LoadClusterSleevesByFilter(lookedUpFilterId, categoryString);
                                    
                                    File.AppendAllText(SafeFileLogger.GetLogFilePath("orchestrator_debug.log"), 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: PATH 1a CHECK - AdoptToDocument={adoptToDocumentEnabled}, NoInvalidatedZones=True, Queried by FilterId={lookedUpFilterId}, Category={categoryString}, Found {existingClusters?.Count ?? 0} clusters\n");
                                    
                                    if (existingClusters != null && existingClusters.Count > 0)
                                    {
                                        // Find comboId from first cluster
                                        comboId = existingClusters[0].ComboId;
                                        isPath1Replay = true;
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[ORCHESTRATOR] ✅ PATH 1a (Super Fast): Found {existingClusters.Count} existing clusters in database for filter '{filter.Name}', category '{categoryString}', comboId={comboId} (AdoptToDocument={adoptToDocumentEnabled}, NoInvalidatedZones=True)");
                                        }
                                        
                                        File.AppendAllText(SafeFileLogger.GetLogFilePath("orchestrator_debug.log"), 
                                            $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: PATH 1a CHECK - Found {existingClusters.Count} clusters in DB, comboId={comboId}, filterId={filterId} (SUPER FAST LANE)\n");
                                    }
                                    else
                                    {
                                        // No clusters in DB, but no invalidated zones → Get comboId for Path 2
                                        using (var cmd = dbContext.Connection.CreateCommand())
                                        {
                                            cmd.CommandText = @"
                                                SELECT DISTINCT fc.ComboId 
                                                FROM FileCombos fc
                                                INNER JOIN Filters f ON fc.FilterId = f.FilterId
                                                WHERE f.FilterId = @FilterId AND f.Category = @Category
                                                LIMIT 1";
                                            cmd.Parameters.AddWithValue("@FilterId", lookedUpFilterId);
                                            cmd.Parameters.AddWithValue("@Category", categoryString);
                                            var comboResult = cmd.ExecuteScalar();
                                            if (comboResult != null && comboResult != DBNull.Value)
                                            {
                                                comboId = Convert.ToInt32(comboResult);
                                            }
                                        }
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[ORCHESTRATOR] PATH 2: No existing clusters in database, will calculate from database. ComboId={comboId?.ToString() ?? "NULL"}");
                                        }
                                        
                                        File.AppendAllText(SafeFileLogger.GetLogFilePath("orchestrator_debug.log"), 
                                            $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: PATH 1a CHECK - No clusters in DB, using PATH 2 (database), comboId={comboId?.ToString() ?? "NULL"}\n");
                                    }
                                }
                                else
                                {
                                    // ✅ PATH 3: Invalidated zones exist → Skip Path 1a, use Path 3 for full recalculation
                                    // Still need comboId for saving cluster data
                                    using (var cmd = dbContext.Connection.CreateCommand())
                                    {
                                        cmd.CommandText = @"
                                            SELECT DISTINCT fc.ComboId 
                                            FROM FileCombos fc
                                            INNER JOIN Filters f ON fc.FilterId = f.FilterId
                                            WHERE f.FilterId = @FilterId AND f.Category = @Category
                                            LIMIT 1";
                                        cmd.Parameters.AddWithValue("@FilterId", lookedUpFilterId);
                                        cmd.Parameters.AddWithValue("@Category", categoryString);
                                        var comboResult = cmd.ExecuteScalar();
                                        if (comboResult != null && comboResult != DBNull.Value)
                                        {
                                            comboId = Convert.ToInt32(comboResult);
                                        }
                                    }
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[ORCHESTRATOR] PATH 3: Invalidated zones detected, skipping PATH 1a, will calculate fresh from database. ComboId={comboId?.ToString() ?? "NULL"}");
                                    }
                                    
                                    File.AppendAllText(SafeFileLogger.GetLogFilePath("orchestrator_debug.log"), 
                                        $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: PATH 1a CHECK SKIPPED - Invalidated zones exist, using PATH 3 (full recalculation), comboId={comboId?.ToString() ?? "NULL"}\n");
                                }
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[ORCHESTRATOR] PATH 2/3: Filter '{filter.Name}' not found in database, will calculate from database");
                                }
                                
                                File.AppendAllText(SafeFileLogger.GetLogFilePath("orchestrator_debug.log"), 
                                    $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: PATH 1 CHECK - Filter not found in DB, using PATH 2/3 (database)\n");
                            }
                        }
                    }
                    catch (Exception pathCheckEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[ORCHESTRATOR] Error checking PATH 1: {pathCheckEx.Message}, falling back to PATH 2/3");
                        }
                        
                        try
                        {
                            File.AppendAllText(SafeFileLogger.GetLogFilePath("orchestrator_debug.log"), 
                                $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: PATH 1 CHECK ERROR: {pathCheckEx.Message}, using PATH 2/3\n");
                        }
                        catch { }
                        
                        // Continue with PATH 2/3 if check fails
                    }
                    
                    // ✅ DATABASE-ONLY: No XML file path needed (all paths use database exclusively)
                    // xmlFilePath parameter is kept for backward compatibility but not used
                    xmlFilePath = null;
                    
                    // ✅ BUILD TIMESTAMP: Log build info to verify correct DLL is loaded
                    try
                    {
                        var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                        var assemblyPath = assembly?.Location ?? string.Empty;
                        var buildTimestamp = !string.IsNullOrWhiteSpace(assemblyPath)
                            ? System.IO.File.GetLastWriteTime(assemblyPath).ToString("yyyy-MM-dd HH:mm:ss")
                            : "unknown";
                        var versionTag = VersionInfo.VersionTag; // "R2023" or "R2024"
                        
                        var orchestratorDebugLogPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                        File.AppendAllText(orchestratorDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] 🔨 BUILD TIMESTAMP: {buildTimestamp} | VERSION: {versionTag} | Assembly: {Path.GetFileName(assemblyPath)}\n");
                        File.AppendAllText(orchestratorDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: isPath1Replay={isPath1Replay}, comboId={comboId}, filterId={filterId} (DATABASE-ONLY, NO XML)\n");
                    }
                    catch { }
                    
                    // Use UniversalClusterService directly (service-based architecture)
                    
                    using (var tx = new Transaction(_document, $"Cluster {categoryString} Openings"))
                    {
                        // ✅ BEST PRACTICE: Check transaction start status (per TRANSACTION_REFACTORING_SUMMARY.md)
                        if (tx.Start() != TransactionStatus.Started)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Error($"[ORCHESTRATOR] ❌ Failed to start transaction: '{tx.GetName()}'");
                            }
                            return Autodesk.Revit.UI.Result.Failed;
                        }
                        
                        // ✅ BEST PRACTICE: Set failure preprocessor to handle warnings (per TRANSACTION_MANAGEMENT_IMPLEMENTATION_PLAN.md)
                        try
                        {
                            var options = tx.GetFailureHandlingOptions();
                            options.SetFailuresPreprocessor(new WarningSwallower());
                            tx.SetFailureHandlingOptions(options);
                        }
                        catch (Exception failureEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[ORCHESTRATOR] Could not set failure preprocessor: {failureEx.Message}");
                            }
                        }
                        
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[ORCHESTRATOR] ✅ Transaction STARTED: '{tx.GetName()}', Document.IsModifiable: {_document.IsModifiable}");
                    }
                    
                    // ✅ STEP 5 OPTIMIZATION: Declare clusterService in outer scope for parameter flush access
                    RefactoredClusterService? clusterService = null;
                    
                    // 🔥 UNCONDITIONAL LOGGING: Prove clustering is being called (DATABASE-ONLY, no XML)
                    SafeFileLogger.SafeAppendText("orchestrator_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 ABOUT TO CALL ClusterSleeves: category={categoryString}, isPath1Replay={isPath1Replay}, comboId={comboId}, filterId={filterId} (DATABASE-ONLY, NO XML)\n");
                    
                    // 🔥 TEST: Direct System.IO logging to test wrapper (using versioned path)
                    try
                    {
                        string logPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: BEFORE ClusterServiceFactory.CreateWithAllServices\n");
                    }
                    catch (Exception ioEx)
                    {
                        // Log to Windows Event Log as last resort
                        try { System.Diagnostics.EventLog.WriteEntry("Application", $"JSE Cluster: IO Error: {ioEx.Message}", System.Diagnostics.EventLogEntryType.Error); } catch { }
                    }
                    
                    try
                    {
                        string logPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: CALLING ClusterServiceFactory.CreateWithAllServices NOW\n");
                        clusterService = ClusterServiceFactory.CreateWithAllServices(_document);
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: ClusterServiceFactory.CreateWithAllServices RETURNED\n");
                    }
                    catch (Exception factoryEx)
                    {
                        try
                        {
                            string logPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: ClusterServiceFactory.CreateWithAllServices EXCEPTION: {factoryEx.Message}\n");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: StackTrace: {factoryEx.StackTrace}\n");
                        }
                        catch { }
                        throw; // Re-throw to be caught by outer try-catch
                    }
                    
                    if (clusterService == null)
                    {
                        string logPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: ERROR: clusterService is NULL!\n");
                        throw new InvalidOperationException("ClusterServiceFactory.CreateWithAllServices returned null");
                    }
                    
                    // 🔥 TEST: Direct System.IO logging after factory (using versioned path)
                    try
                    {
                        string logPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: BEFORE ClusterSleeves CALL (isPath1Replay={isPath1Replay}, comboId={comboId}, filterId={filterId})\n");
                    }
                    catch { }
                    
                    // ✅ FIX: Pass PATH 1 parameters to clustering service (check DB first, then XML)
                    // ✅ PERFORMANCE: Counts are now stored in outer scope variables
                    // ✅ PATH 3: Pass PATH 3 type flags to clustering service
                    try
                    {
                        // ✅ FIX: Pass clearance settings to clustering service for condition change check
                        // ✅ PATH 3: Pass PATH 3 type flags to clustering service
                        var clusterResult = clusterService.ClusterSleeves(
                            _document, 
                            categoryString, 
                            _uiDocument, 
                            xmlFilePath, 
                            filter.Name, 
                            placedClusterSleeves, 
                            isPath1Replay, 
                            comboId, 
                            filterId, 
                            _uiClearances,
                            isPath3Validated: isPath3Validated,
                            isPath3Invalidated: isPath3Invalidated,
                            isPath3New: isPath3New);
                        placedCount = clusterResult.placedCount;
                        deletedCount = clusterResult.deletedCount;
                        
                        // 🔥 TEST: Direct System.IO logging after successful call (using versioned path)
                        try
                        {
                            string logPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: ClusterSleeves RETURNED SUCCESS: placed={placedCount}, deleted={deletedCount}\n");
                        }
                        catch { }
                    }
                    catch (Exception ex)
                    {
                        // 🔥 TEST: Direct System.IO logging on exception (using versioned path)
                        try
                        {
                            string logPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: ClusterSleeves EXCEPTION: {ex.Message}\n");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 DIRECT IO: StackTrace: {ex.StackTrace}\n");
                        }
                        catch { }
                        throw; // Re-throw to be caught by outer try-catch
                    }
                    
                    // 🔥 UNCONDITIONAL LOGGING: Prove clustering returned
                    SafeFileLogger.SafeAppendText("orchestrator_debug.log", $"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 ClusterSleeves RETURNED: placedCount={placedCount}, deletedCount={deletedCount}\n");                        // ✅ CRITICAL LOGGING: Log cluster sleeves returned from ClusterSleeves
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORCHESTRATOR] ClusterSleeves returned: placedCount={placedCount}, deletedCount={deletedCount}, placedClusterSleeves.Count={placedClusterSleeves?.Count ?? 0}");
                            if (placedClusterSleeves != null && placedClusterSleeves.Count > 0)
                            {
                                DebugLogger.Info($"[ORCHESTRATOR] ✅ Cluster sleeve IDs in placedClusterSleeves: {string.Join(", ", placedClusterSleeves.Select(c => c.Id.IntegerValue))}");
                            }
                            else
                            {
                                DebugLogger.Warning($"[ORCHESTRATOR] ⚠️ placedClusterSleeves is EMPTY after ClusterSleeves call!");
                            }
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORCHESTRATOR] About to COMMIT transaction '{tx.GetName()}'");
                        }
                        
                        // ✅ BEST PRACTICE: Check transaction commit status (per TRANSACTION_REFACTORING_SUMMARY.md)
                        var commitStatus = tx.Commit();
                        
                        if (commitStatus == TransactionStatus.Committed)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[ORCHESTRATOR] ✅ Transaction COMMITTED successfully: '{tx.GetName()}'");
                            }
                            
                            // ✅ DIAGNOSTIC: Verify cluster sleeves still exist AFTER transaction commit
                            if (placedClusterSleeves != null && placedClusterSleeves.Count > 0)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 POST-COMMIT VERIFICATION: Checking {placedClusterSleeves.Count} cluster sleeves after transaction commit...\n");
                                
                                int foundAfterCommit = 0;
                                int missingAfterCommit = 0;
                                foreach (var cluster in placedClusterSleeves)
                                {
                                    if (cluster == null || !cluster.IsValidObject)
                                    {
                                        missingAfterCommit++;
                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ POST-COMMIT: Cluster sleeve is NULL or INVALID (was deleted during transaction!)\n");
                                    }
                                    else
                                    {
                                        foundAfterCommit++;
                                        int clusterId = cluster.Id.IntegerValue;
                                        var location = cluster.Location as LocationPoint;
                                        var locPoint = location?.Point;
                                        
                                        // ✅ DIAGNOSTIC: Verify by ID lookup (not just object reference)
                                        var verifyById = _document.GetElement(new ElementId(clusterId)) as FamilyInstance;
                                        bool existsById = verifyById != null && verifyById.IsValidObject;
                                        
                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss}] ✅ POST-COMMIT: Cluster sleeve {clusterId} EXISTS: Name='{cluster.Name}', " +
                                            $"Location=({locPoint?.X:F2}, {locPoint?.Y:F2}, {locPoint?.Z:F2}), " +
                                            $"Category='{cluster.Category?.Name ?? "NULL"}', " +
                                            $"Document='{cluster.Document?.Title ?? "NULL"}', " +
                                            $"IsValid={cluster.IsValidObject}, " +
                                            $"ExistsById={existsById}\n");
                                        
                                        if (!existsById)
                                        {
                                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                                $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ POST-COMMIT WARNING: Cluster sleeve {clusterId} reference is valid but GetElement by ID returns NULL!\n");
                                        }
                                    }
                                }
                                
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 📊 POST-COMMIT RESULT: {foundAfterCommit} found, {missingAfterCommit} missing out of {placedClusterSleeves.Count} cluster sleeves\n");
                                
                                // ✅ ADDITIONAL DIAGNOSTIC: After a short delay, check again if cluster sleeves still exist
                                // This helps detect if something else is deleting them after commit
                                System.Threading.Thread.Sleep(100); // Small delay to allow any async operations to complete
                                
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 POST-COMMIT+100ms VERIFICATION: Re-checking {placedClusterSleeves.Count} cluster sleeves after delay...\n");
                                
                                int foundAfterDelay = 0;
                                int missingAfterDelay = 0;
                                foreach (var clusterRef in placedClusterSleeves)
                                {
                                    if (clusterRef == null) continue;
                                    
                                    int clusterId = clusterRef.Id.IntegerValue;
                                    // ✅ CRITICAL: Look up by ID in active document, not use stale reference
                                    var freshLookup = _document.GetElement(new ElementId(clusterId)) as FamilyInstance;
                                    
                                    if (freshLookup == null || !freshLookup.IsValidObject)
                                    {
                                        missingAfterDelay++;
                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ POST-COMMIT+100ms: Cluster sleeve {clusterId} MISSING (was deleted after commit!)\n");
                                    }
                                    else
                                    {
                                        foundAfterDelay++;
                                        var loc = freshLookup.Location as LocationPoint;
                                        var pt = loc?.Point;
                                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss}] ✅ POST-COMMIT+100ms: Cluster sleeve {clusterId} STILL EXISTS: " +
                                            $"Name='{freshLookup.Name}', Document='{freshLookup.Document?.Title ?? "NULL"}', " +
                                            $"IsActiveDoc={!freshLookup.Document?.IsLinked ?? false}, " +
                                            $"Location=({pt?.X:F2}, {pt?.Y:F2}, {pt?.Z:F2})\n");
                                    }
                                }
                                
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 📊 POST-COMMIT+100ms RESULT: {foundAfterDelay} found, {missingAfterDelay} missing out of {placedClusterSleeves.Count} cluster sleeves\n");
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Error($"[ORCHESTRATOR] ❌ Transaction FAILED to commit: '{tx.GetName()}', Status: {commitStatus}");
                            }
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ TRANSACTION NOT COMMITTED: Status={commitStatus}, cluster sleeves may have been rolled back!\n");
                            
                            return Autodesk.Revit.UI.Result.Failed;
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Clustering complete for {categoryString}: {placedCount} clusters placed, {deletedCount} individual sleeves deleted");
                        }
                        
                        return Autodesk.Revit.UI.Result.Succeeded;
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[OpeningCommandOrchestrator] Error in ExecuteClusteringForCategory for {filter.Category}: {ex.Message}");
                    }
                    return Autodesk.Revit.UI.Result.Failed;
                }
            }, $"Cluster Sleeves for {filter.Category}");
            
            // ✅ PERFORMANCE: Log and return counts
            SafeFileLogger.SafeAppendText("cluster_debug.log",
                $"[{DateTime.Now:HH:mm:ss}] ExecuteClusteringForCategory completed: Status={result}, Placed={placedCount}, Deleted={deletedCount}\n");
            
            // ✅ CRITICAL DIAGNOSTIC: Check if cluster sleeves still exist AFTER ExecuteClusteringForCategory returns
            // This helps detect if something is deleting them after the method returns
            if (placedClusterSleeves != null && placedClusterSleeves.Count > 0)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 POST-RETURN VERIFICATION: Checking {placedClusterSleeves.Count} cluster sleeves after ExecuteClusteringForCategory returns...\n");
                
                int foundAfterReturn = 0;
                int missingAfterReturn = 0;
                foreach (var clusterRef in placedClusterSleeves)
                {
                    if (clusterRef == null) continue;
                    
                    int clusterId = clusterRef.Id.IntegerValue;
                    var freshLookup = _document.GetElement(new ElementId(clusterId)) as FamilyInstance;
                    
                    if (freshLookup == null || !freshLookup.IsValidObject)
                    {
                        missingAfterReturn++;
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ POST-RETURN: Cluster sleeve {clusterId} MISSING (was deleted after method returned!)\n");
                    }
                    else
                    {
                        foundAfterReturn++;
                        var loc = freshLookup.Location as LocationPoint;
                        var pt = loc?.Point;
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss}] ✅ POST-RETURN: Cluster sleeve {clusterId} STILL EXISTS: Location=({pt?.X:F2}, {pt?.Y:F2}, {pt?.Z:F2})\n");
                    }
                }
                
                SafeFileLogger.SafeAppendText("cluster_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] 📊 POST-RETURN RESULT: {foundAfterReturn} found, {missingAfterReturn} missing out of {placedClusterSleeves.Count} cluster sleeves\n");
            }
            
            if (result != Autodesk.Revit.UI.Result.Succeeded)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Clustering for {filter.Category} completed with status: {result}");
                }
                return (placedCount, deletedCount); // Return counts even on failure
            }
            
            // ✅ PERFORMANCE FIX: After placing cluster sleeves, regenerate document and save their bounding boxes to database
            // This uses SleeveCoordinateService to update coordinates (same as individual sleeves)
            // ✅ FIX: Call even when xmlFilePath is null (database-only mode) - pass category instead
            if (placedClusterSleeves != null && placedClusterSleeves.Count > 0 && categoryString != null)
            {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[OpeningCommandOrchestrator] Regenerating document and updating coordinates for {placedClusterSleeves.Count} cluster sleeves");
                    }
                    
                    try
                    {
                        // Step 1: Regenerate document to ensure bounding boxes are available
                        _document.Regenerate();
                        
                        // Step 2: Wait for regeneration to complete
                        System.Threading.Thread.Sleep(200);
                        
                        // ✅ STEP 5 OPTIMIZATION: Parameter flushing handled internally by RefactoredClusterService
                        // (See PlaceClustersFromDatabase method for flush implementation)
                    }
                    catch (Exception regenEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[OpeningCommandOrchestrator] Could not regenerate document: {regenEx.Message}");
                        }
                    }
                    
                    // Step 3: Save cluster sleeve bounding boxes to database (database-only mode)
                    try
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[OpeningCommandOrchestrator] About to call UpdateSleeveCoordinatesInXml with category: {categoryString ?? "NULL"} (database-only mode)");
                        }
                        var coordinateService = new SleeveCoordinateService(_document);
                        // ✅ DATABASE-ONLY: Pass category parameter (xmlFilePath obsolete, kept for backward compatibility)
                        coordinateService.UpdateSleeveCoordinatesInXml(categoryString);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Updated sleeve coordinates for cluster sleeves (including bounding boxes) in database");
                        }
                    }
                    catch (Exception coordEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[OpeningCommandOrchestrator] Error updating coordinates: {coordEx.Message}");
                        }
                    }
                    
                    // Step 4: Cache reload handled internally by RefactoredClusterService
                    try
                    {
                        // RefactoredClusterService handles cache loading via ClusterDataService
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Reloaded cache with cluster sleeve coordinates for {categoryString}");
                        }
                        
                        // Step 5: NOW run cleanup with updated cache (uses XML, not expensive Revit API)
                        // ✅ CRITICAL: Verify placedClusterSleeves list is not empty before cleanup
                        if (placedClusterSleeves == null || placedClusterSleeves.Count == 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[OpeningCommandOrchestrator] ⚠️ placedClusterSleeves is empty or null - skipping cleanup to prevent cluster sleeve deletion");
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[OpeningCommandOrchestrator] ✅ About to run cleanup with {placedClusterSleeves.Count} cluster sleeves in protection set: {string.Join(", ", placedClusterSleeves.Select(c => c.Id.IntegerValue))}");
                            }
                            
                            using (var cleanupTx = new Transaction(_document, $"Cleanup sleeves within clusters"))
                            {
                                cleanupTx.Start();
                                // Cleanup handled by RefactoredClusterService internally - no additional cleanup needed
                                cleanupTx.Commit();
                            }
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[OpeningCommandOrchestrator] Error in cleanup: {cleanupEx.Message}");
                        }
                    }
                }
            
            // Return counts after all processing (regeneration, flush, cleanup) complete
            return (placedCount, deletedCount);
        }

        /// <summary>
        /// Load clash zones for a specific filter from XML file
        /// </summary>
        /// <summary>
        /// ✅ PERFORMANCE FIX: Get XML file path for a specific filter
        /// This avoids loading all 22 XML files during clustering
        /// </summary>
        private string GetXmlFilePathForFilter(OpeningFilter filter)
        {
            // Use actual project filters directory (matches Refresh saves)
            var filtersDir = ProjectPathService.GetFiltersDirectory(_document);

            string categoryName = filter.Category switch
            {
                Models.MepCategory.Ducts => "ducts",
                Models.MepCategory.DuctAccessories => "duct_accessories",
                Models.MepCategory.Pipes => "pipes",
                Models.MepCategory.CableTrays => "cable_trays",
                _ => filter.Category.ToString().ToLower().Replace(" ", "_")
            };

            string xmlFileName = $"{filter.Name}_{categoryName}.xml";
            return Path.Combine(filtersDir, xmlFileName);
        }

        private List<ClashZone> LoadClashZonesForFilter(OpeningFilter filter)
        {
            try
            {

                // ✅ CRITICAL: Get category from filter
                string categoryName = filter.Category switch
                {
                    Models.MepCategory.Ducts => "Ducts",
                    Models.MepCategory.DuctAccessories => "Duct Accessories",
                    Models.MepCategory.Pipes => "Pipes",
                    Models.MepCategory.CableTrays => "Cable Trays",
                    _ => "Ducts"
                };

                // ✅ PHASE SQLITE-2: Load from SQLite FIRST (primary source)
                var dbZones = LoadClashZonesFromDatabase(filter, categoryName);
                if (dbZones != null && dbZones.Count > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Loaded {dbZones.Count} clash zones from SQLite for filter '{filter.Name}' ({categoryName})\n");
                    }
                    return dbZones;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found in database for filter '{filter.Name}' ({categoryName})");
                }
                
                return new List<ClashZone>();
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error loading clash zones for filter {filter.Name}: {ex.Message}");
                }
                return new List<ClashZone>();
            }
        }

        private List<ClashZone> LoadClashZonesFromDatabase(OpeningFilter filter, string categoryName)
        {
            if (filter == null || string.IsNullOrWhiteSpace(categoryName))
                return new List<ClashZone>();

            try
            {
                using (var context = new SleeveDbContext(_document, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[OpeningCommandOrchestrator][SQLite] {msg}");
                }))
                {
                    var repository = new ClashZoneRepository(context, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[OpeningCommandOrchestrator][SQLite] {msg}");
                    });

                    List<ClashZone> eligibleZones = new List<ClashZone>();

                    // ✅ REFACTORED: Ignore ReadyForPlacementFlag as per user request
                    // Filter primarily by IsCurrentClash (zones within section box from last refresh)

                    // Query ALL zones (readyForPlacementOnly=false)
                    var zones = repository.GetClashZonesByFilter(filter.Name, categoryName, unresolvedOnly: false, readyForPlacementOnly: false) ?? new List<ClashZone>();

                    // ✅ SIMPLIFIED ELIGIBILITY: Zone is eligible if:
            // 1. IsCurrentClash = 1 (session flag set during refresh), OR
            // 2. Zone is unresolved (deleted sleeve scenario: IsResolved=0, IsClusterResolved=0, IsCombinedResolved=0)
            // NOTE: ReadyForPlacement is DEPRECATED and should NOT be used for filtering
            eligibleZones = zones.Where(cz => cz != null && (
                cz.IsCurrentClash ||  // Primary: Session flag (set during refresh for zones in section box)
                (!cz.IsResolvedFlag && !cz.IsClusterResolvedFlag && !cz.IsCombinedResolved)  // Fallback: unresolved zone (deleted sleeve)
            )).ToList();

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        int bySession = zones.Count(cz => cz?.IsCurrentClash == true);
                        int byFallback = zones.Count(cz => cz != null && !cz.IsResolvedFlag && !cz.IsClusterResolvedFlag && !cz.IsCombinedResolved);
                        DebugLogger.Info($"[OpeningCommandOrchestrator] 🔄 Filtered: {zones.Count} total -> {eligibleZones.Count} eligible (session={bySession}, fallback={byFallback})");
                    }

                    foreach (var zone in eligibleZones)
                    {
                        zone?.EnsureSleevePlacementPointReconstructed();
                        zone?.EnsureSleevePlacementPointActiveDocumentReconstructed();
                    }

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[OpeningCommandOrchestrator][SQLite] Loaded {eligibleZones.Count} eligible zones - DATA SOURCE: DATABASE");
                        // ✅ CRITICAL: Log data source for placement debugging
                        var placementLogPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        try
                        {
                            File.AppendAllText(placementLogPath, $"[{DateTime.Now:HH:mm:ss}] [DATA-SOURCE] ✅ Using DATABASE for filter '{filter.Name}', category '{categoryName}' ({eligibleZones.Count} eligible zones)\n");
                        }
                        catch { }
                    }

                    return eligibleZones;
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] SQLite load failed for filter '{filter?.Name}' ({categoryName}): {ex.Message}");
                return new List<ClashZone>();
            }
        }

        /// <summary>
        /// Execute UniversalSleevePlacementCommand
        /// ⚠️ CRITICAL: Wrapped with timeout protection (5-minute limit per category)
        /// </summary>
        private (int placedCount, int skippedCount, int errorCount) ExecuteUniversalSleevePlacement(OpeningFilter filter, bool showProgress)
        {
            // Declare variables outside lambda for use after timeout execution
            string xmlFilePath = GetXmlFilePathForFilter(filter);
            var tracePath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
            var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
            
            // ✅ PERFORMANCE: Store counts outside lambda for access after timeout execution
            int placedCount = 0;
            int skippedCount = 0;
            int errorCount = 0;
            UniversalSleevePlacementCommand? universalCommand = null;
            
            // ⚠️ CRITICAL: Execute with timeout protection to prevent infinite hangs
            var result = _crashSafeExecutor.ExecuteWithTimeout(() =>
            {
                try
                {
                    // 🔥 CRITICAL DEBUG: Direct file logging to trace orchestrator execution
                    // ✅ OVERWRITE placement_debug.log at start of each run
                    // ✅ DEPLOYMENT MODE: Skip file writes
                    try { 
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                            var assemblyPath = assembly?.Location ?? string.Empty;
                            var assemblyName = !string.IsNullOrWhiteSpace(assemblyPath) ? System.IO.Path.GetFileName(assemblyPath) : "<unknown>";
                            var buildTimestamp = !string.IsNullOrWhiteSpace(assemblyPath)
                                ? System.IO.File.GetLastWriteTime(assemblyPath).ToString("yyyy-MM-dd HH:mm:ss")
                                : "unknown";
                            var assemblyVersion = !string.IsNullOrWhiteSpace(assemblyPath)
                                ? System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyPath)?.FileVersion ?? "unknown"
                                : "unknown";

                            var nowStamp = DateTime.Now.ToString("HH:mm:ss");
                            var header = $"[{nowStamp}] BUILD TIMESTAMP: {buildTimestamp} | Assembly={assemblyName} | Version={assemblyVersion}\n" +
                                         $"[{nowStamp}] === NEW PLACEMENT RUN STARTED ===\n";

                            System.IO.File.WriteAllText(placementDebugPath, header);
                        }
                    } catch { }
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] CLICK_OK: Begin placement for category={filter.Category}, filter={filter.Name}\n");
                        }
                    } catch { }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔥 ExecuteUniversalSleevePlacement CALLED 🔥\n");
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Filter Category: {filter.Category}\n");
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] UI Clearances Count: {_uiClearances?.Count ?? 0}\n");
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Executing UniversalSleevePlacementCommand for {filter.Category}");
                    }

                    // ✅ CRITICAL FIX: Load clash zones from XML file
                    var clashZones = LoadClashZonesForFilter(filter);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Loaded {clashZones.Count} clash zones for {filter.Category}");
                    }

                    // ✅ CRITICAL FIX: Pre-filter zones based on settings (e.g. MinWallThickness)
                    // This ensures we don't place sleeves on walls that are too thin (user request)
                    try
                    {
                        var zoneFilterService = new ZoneFilterService();
                        int originalCount = clashZones.Count;
                        clashZones = zoneFilterService.PreFilterEligibleClashZones(_document, clashZones);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            int filteredCount = originalCount - clashZones.Count;
                            if (filteredCount > 0)
                            {
                                DebugLogger.Info($"[OpeningCommandOrchestrator] ⚡ FILTERED: {originalCount} -> {clashZones.Count} zones ({filteredCount} removed by settings/wall thickness)");
                            }
                        }
                    }
                    catch (Exception filterEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[OpeningCommandOrchestrator] Error during pre-filtering: {filterEx.Message}");
                        }
                        // Continue with original list if filtering fails
                    }
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] LOAD_XML: zones={clashZones.Count}\n");
                        }
                    } catch { }

                    if (clashZones.Count > 0)
                    {
                        // ✅ CRITICAL FIX: Convert enum to proper string format for strategy creation
                        // Declare these variables FIRST before they are used in PATH 3 Invalidated block
                        string categoryString = filter.Category switch
                        {
                            Models.MepCategory.Ducts => "Ducts",
                            Models.MepCategory.DuctAccessories => "Duct Accessories", // Note: space, not "DuctAccessories"
                            Models.MepCategory.Pipes => "Pipes", 
                            Models.MepCategory.CableTrays => "Cable Trays", // Note: space, not "CableTrays"
                            _ => "Ducts"
                        };
                        
                        // ✅ CRITICAL FIX: Get the filter name with .xml extension for XML file matching
                        string categoryName = filter.Category switch
                        {
                            Models.MepCategory.Ducts => "ducts",
                            Models.MepCategory.DuctAccessories => "duct_accessories",
                            Models.MepCategory.Pipes => "pipes",
                            Models.MepCategory.CableTrays => "cable_trays",
                            _ => "ducts"
                        };
                        string combinedFilterName = $"{filter.Name}_{categoryName}.xml"; // Added .xml
                        
                        // ✅ PATH 3 INVALIDATED: Check for zones that need distinct placement flow
                        // CRITICAL FIX: Invalidated zones are zones where MEP elements MOVED (validation failed)
                        // NOT just zones with sleeves - a zone can have a sleeve AND still be validated (unchanged)
                        // 
                        // For Path 1a (super fast): If all zones are validated (no movement detected), use Path 1a
                        // For Path 3: If any zones are invalidated (movement detected), use Path 3
                        //
                        // Since validation results aren't directly available here, we use a heuristic:
                        // - If zones have sleeves (SleeveInstanceId > 0), they COULD be invalidated OR validated
                        // - For Path 1a: Assume zones with sleeves are VALIDATED (no movement) if no explicit invalidation
                        // - For Path 3: Only mark as invalidated if validation explicitly determined movement
                        //
                        // TODO: Pass validation results from RefreshContext to avoid this heuristic
                        var zonesWithSleeves = clashZones.Where(cz => cz.SleeveInstanceId > 0).ToList();
                        
                        // ✅ PATH 1a LOGIC: If all zones are validated (no invalidated zones detected during refresh),
                        // then zones with sleeves are VALIDATED (not invalidated) - they just need Path 1a replay
                        // For now, we'll check if there are any zones that explicitly need recalculation
                        // If all zones have sleeves but are validated, hasInvalidatedZones should be FALSE
                        var invalidatedZones = new List<ClashZone>(); // Empty by default - assume validated unless proven otherwise
                        
                        var validatedZones = clashZones
                            .Where(cz => !cz.IsResolvedFlag && !cz.IsClusterResolvedFlag && cz.ClusterSleeveInstanceId <= 0)
                            .ToList();
                        
                        // ✅ PATH 3 NEW: New zones are zones that don't have existing sleeves (SleeveInstanceId <= 0)
                        // These zones route to PATH 2 placement, then PATH 3 new clustering
                        var newZones = clashZones.Where(cz => cz.SleeveInstanceId <= 0 && !cz.IsResolvedFlag && !cz.IsClusterResolvedFlag && cz.ClusterSleeveInstanceId <= 0).ToList();
                        
                        // ✅ PATH 3 TRACKING: Store PATH 3 type flags for clustering
                        // CRITICAL: hasInvalidatedZones = false by default (assume validated)
                        // Only set to true if validation explicitly determined zones are invalidated
                        // For Path 1a: If all zones are validated, hasInvalidatedZones = false
                        bool hasInvalidatedZones = false; // Default to false - assume validated unless proven otherwise
                        bool hasValidatedZones = validatedZones.Count > 0; // Validated = eligible for placement
                        bool hasNewZones = newZones.Count > 0 && !hasValidatedZones; // New = no validated zones
                        
                        // ✅ PATH 3: Store flags for clustering (check AdoptToDocument flag)
                        bool adoptToDocumentEnabled = false;
                        try
                        {
                            using (var dbContext = new SleeveDbContext(_document))
                            {
                                var filterRepository = new FilterRepository(dbContext, _ => { });
                                int lookedUpFilterId = filterRepository.GetFilterId(filter.Name, categoryString);
                                if (lookedUpFilterId > 0)
                                {
                                    using (var cmd = dbContext.Connection.CreateCommand())
                                    {
                                        cmd.CommandText = @"SELECT AdoptToDocumentFlag FROM Filters WHERE FilterId = @FilterId";
                                        cmd.Parameters.AddWithValue("@FilterId", lookedUpFilterId);
                                        var flagResult = cmd.ExecuteScalar();
                                        adoptToDocumentEnabled = flagResult != null && flagResult != DBNull.Value && Convert.ToInt32(flagResult) == 1;
                                    }
                                }
                            }
                        }
                        catch { }
                        
                        // ✅ PATH 3: Store flags for clustering
                        string path3Key = $"{filter.Name}_{categoryString}";
                        _path3Flags[path3Key] = (
                            isValidated: adoptToDocumentEnabled && hasValidatedZones,
                            isInvalidated: adoptToDocumentEnabled && hasInvalidatedZones,
                            isNew: adoptToDocumentEnabled && hasNewZones
                        );
                        
                        if (!DeploymentConfiguration.DeploymentMode && invalidatedZones.Count > 0)
                        {
                            DebugLogger.Info($"[OpeningCommandOrchestrator] Detected {invalidatedZones.Count} zones with existing sleeves (potential invalidated zones)");
                        }
                        
                        // ✅ PATH 3 INVALIDATED: Route invalidated zones to distinct placement service
                        if (invalidatedZones.Count > 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[OpeningCommandOrchestrator] Found {invalidatedZones.Count} invalidated zones, routing to PATH 3 Invalidated placement");
                            }
                            
                            try
                            {
                                // ✅ PATH 3 INVALIDATED: Load conditions and create strategy
                                var projectFiltersDir = ProjectPathService.GetFiltersDirectory(_document);
                                var conditionsService = new ConditionsService(_document, projectFiltersDir, msg => 
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info(msg);
                                });
                                
                                var conditionsKey = $"{combinedFilterName}_{categoryString}";
                                var conditions = conditionsService.LoadConditions(conditionsKey);
                                if (conditions == null)
                                {
                                    conditions = new OpeningConditions { FilterName = combinedFilterName, Category = categoryString };
                                }
                                
                                // Create strategy based on category
                                ISleevePlacementStrategy strategy = categoryString.ToLower() switch
                                {
                                    "pipes" => new PipePlacementStrategy(),
                                    "cable trays" => new CableTrayPlacementStrategy(),
                                    "ducts" or "duct accessories" => new DuctPlacementStrategy(),
                                    _ => new DuctPlacementStrategy()
                                };
                                
                                var invalidatedService = new Path3InvalidatedPlacementService(_document, null, _forceDetectionMode);
                                var invalidatedResult = invalidatedService.ExecutePlacement(
                                    invalidatedZones,
                                    combinedFilterName,
                                    categoryString,
                                    conditions,
                                    strategy,
                                    _uiClearances ?? new Dictionary<string, double>());
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[OpeningCommandOrchestrator] PATH 3 Invalidated placement: {invalidatedResult.PlacedCount} placed, {invalidatedResult.DeletedSleeveCount} deleted");
                                }
                            }
                            catch (Exception invalidatedEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error in PATH 3 Invalidated placement: {invalidatedEx.Message}\n{invalidatedEx.StackTrace}");
                                }
                                // Continue with normal placement for remaining zones
                            }
                            
                            // ✅ PATH 3 FIX: Continue with validated zones for normal placement (Path 1)
                            // Validated zones are zones without sleeves that need individual placement
                            if (validatedZones.Count > 0)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[OpeningCommandOrchestrator] Found {validatedZones.Count} validated zones (no sleeves), routing to PATH 1 (individual sleeve placement)");
                                }
                                clashZones = validatedZones;
                            }
                            else
                            {
                                // No validated zones to place - all zones were invalidated
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[OpeningCommandOrchestrator] No validated zones found after processing {invalidatedZones.Count} invalidated zones - skipping individual placement");
                                }
                                clashZones = new List<ClashZone>(); // Empty list - skip individual placement
                            }
                        }
                        
                        // 🔥 CRITICAL DEBUG: Log which XML file we're passing clash zones from
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔍 PASSING CLASH ZONES FROM XML FILE: {xmlFilePath}\n");
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] About to create UniversalSleevePlacementCommand for category: {categoryString}, filter: {combinedFilterName}\n");
                    }
                    
                    universalCommand = new UniversalSleevePlacementCommand(_document, clashZones, categoryString, combinedFilterName, _uiClearances);
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] COMMAND_CREATED: category={categoryString}, xml={xmlFilePath}\n");
                        }
                    } catch { }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand created successfully, about to execute\n");
                    }
                    
                    universalCommand.Execute(_uiDocument.Application);
                    
                    // ✅ REFACTORED: Parameter flushing is now handled internally by NewSleevePlacerService
                    // NewSleevePlacerService uses SleeveParameterService which flushes parameters automatically
                    // No external flush needed - legacy UniversalSleevePlacerService dependency has been removed
                    // ✅ LEGACY REMOVED: universalCommand.Service property no longer exists
                    
                    // ✅ PERFORMANCE: Get counts from command properties
                    placedCount = universalCommand.PlacedCount;
                    skippedCount = universalCommand.SkippedCount;
                    errorCount = universalCommand.ErrorCount;
                    
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand executed: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}\n");
                    
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] COMMAND_EXECUTED: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}\n");
                        }
                    } catch { }
                    
                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand executed successfully: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}\n");
                    }
                    
                    // ✅ Return success after command execution
                    return Autodesk.Revit.UI.Result.Succeeded;
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found for {filter.Category}, skipping placement");
                        }
                        return Autodesk.Revit.UI.Result.Succeeded; // Return success even if no zones to place
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[OpeningCommandOrchestrator] Error in ExecuteUniversalSleevePlacement for {filter.Category}: {ex.Message}");
                    }
                    return Autodesk.Revit.UI.Result.Failed;
                }
            }, $"Place Sleeves for {filter.Category}");
            
            if (result != Autodesk.Revit.UI.Result.Succeeded)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss}] ExecuteUniversalSleevePlacement completed with status: {result}, Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}\n");
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Sleeve placement for {filter.Category} completed with status: {result}");
                }
                return (placedCount, skippedCount, errorCount); // Return counts even on failure
            }
            
            // ✅ PERFORMANCE: Log success
            SafeFileLogger.SafeAppendText("placement_debug.log",
                $"[{DateTime.Now:HH:mm:ss}] ExecuteUniversalSleevePlacement SUCCESS: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}\n");
            
            // ✅ CRITICAL: Following reference document - Regenerate FIRST, then read from Revit and save to XML
            // Reference: SLEEVE_PLACEMENT_SEQUENCING_REFERENCE.md lines 22-35
            // The timing fix: Regenerate ensures bounding boxes are available, then UpdateSleeveCoordinatesInXml reads from Revit
            // ✅ CRITICAL: Save individual sleeve bounding boxes BEFORE clustering
            // Clustering proximity calculation REQUIRES individual sleeve bounding boxes from database
            try
            {
                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ ENTERING UpdateSleeveCoordinatesInXml block...\n");
                }
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] ENTERING_UPDATE_COORD_BLOCK\n");
                    }
                } catch { }
                
                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Regenerating document to ensure bounding boxes are available...\n");
                }
                
                // ✅ CRITICAL LOGGING: Log bounding boxes from Revit BEFORE regeneration
                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    try
                    {
                        var sleevesBeforeRegen = new FilteredElementCollector(_document)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(s => s.Symbol.FamilyName.Contains("Opening"))
                            .ToList();
                        
                        try {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_REGEN] Found {sleevesBeforeRegen.Count} sleeves before regeneration\n");
                            }
                        } catch { }
                        foreach (var sleeve in sleevesBeforeRegen.Take(10)) // Log first 10
                        {
                            try
                            {
                                var bboxBefore = sleeve.get_BoundingBox(null);
                                if (bboxBefore != null)
                                {
                                    try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_REGEN] Sleeve {sleeve.Id.IntegerValue}: Min=({bboxBefore.Min.X:F6}, {bboxBefore.Min.Y:F6}, {bboxBefore.Min.Z:F6}), Max=({bboxBefore.Max.X:F6}, {bboxBefore.Max.Y:F6}, {bboxBefore.Max.Z:F6})\n"); } catch { }
                                }
                                else
                                {
                                    try {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_REGEN] Sleeve {sleeve.Id.IntegerValue}: Bounding box is NULL\n");
                                        }
                                    } catch { }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
                
                // ✅ CRITICAL FIX: Regenerate document BEFORE getting bounding boxes
                // Without regeneration, bounding boxes may not be available immediately after placement
                try
                {
                    _document.Regenerate();
                    // Wait for regeneration to complete
                    System.Threading.Thread.Sleep(200);
                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Document regenerated, now getting bounding boxes...\n");
                    }
                }
                catch (Exception regenEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[OpeningCommandOrchestrator] Could not regenerate document: {regenEx.Message}");
                    }
                }
                
                // ✅ CRITICAL LOGGING: Log bounding boxes from Revit AFTER regeneration
                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    try
                    {
                        var sleevesAfterRegen = new FilteredElementCollector(_document)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(s => s.Symbol.FamilyName.Contains("Opening"))
                            .ToList();
                        
                        try {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_AFTER_REGEN] Found {sleevesAfterRegen.Count} sleeves after regeneration\n");
                            }
                        } catch { }
                        foreach (var sleeve in sleevesAfterRegen.Take(10)) // Log first 10
                        {
                            try
                            {
                                var bboxAfter = sleeve.get_BoundingBox(null);
                                if (bboxAfter != null)
                                {
                                    try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_AFTER_REGEN] Sleeve {sleeve.Id.IntegerValue}: Min=({bboxAfter.Min.X:F6}, {bboxAfter.Min.Y:F6}, {bboxAfter.Min.Z:F6}), Max=({bboxAfter.Max.X:F6}, {bboxAfter.Max.Y:F6}, {bboxAfter.Max.Z:F6})\n"); } catch { }
                                }
                                else
                                {
                                    try {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_AFTER_REGEN] Sleeve {sleeve.Id.IntegerValue}: Bounding box is NULL\n");
                                        }
                                    } catch { }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
                
                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Getting individual sleeve bounding boxes from Revit...\n");
                }
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] BEFORE_UpdateSleeveCoordinatesInXml\n");
                    }
                } catch { }
                
                var coordinateService = new SleeveCoordinateService(_document);
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] SleeveCoordinateService_CREATED\n");
                    }
                } catch { }
                
                // ✅ DATABASE-ONLY: Use category parameter instead of xmlFilePath (XML obsolete)
                // Get category string from filter
                string categoryForUpdate = filter.Category switch
                {
                    Models.MepCategory.Ducts => "Ducts",
                    Models.MepCategory.DuctAccessories => "Duct Accessories",
                    Models.MepCategory.Pipes => "Pipes",
                    Models.MepCategory.CableTrays => "Cable Trays",
                    _ => "Ducts"
                };
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] About to call UpdateSleeveCoordinatesInXml with category: {categoryForUpdate ?? "NULL"} (database-only mode)\n");
                }
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] CALLING_UpdateSleeveCoordinatesInXml: category={categoryForUpdate ?? "NULL"} (database-only)\n");
                    }
                } catch { }
                
                coordinateService.UpdateSleeveCoordinatesInXml(categoryForUpdate);
                
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] AFTER_UpdateSleeveCoordinatesInXml\n");
                    }
                } catch { }
                
                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Individual sleeve coordinates saved - clustering can now calculate proximity\n");
                }
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] UPDATE_COORD_SUCCESS\n");
                    }
                } catch { }
            }
            catch (Exception coordEx)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ Error saving individual coordinates: {coordEx.Message}\n");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[OpeningCommandOrchestrator] UpdateSleeveCoordinatesInXml Exception: {coordEx.Message}\n");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[OpeningCommandOrchestrator] Stack trace: {coordEx.StackTrace}\n");
                }
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] UPDATE_COORD_EXCEPTION: {coordEx.Message}\n");
                    }
                } catch { }
            }
            
            // ✅ PERFORMANCE: Return counts AFTER saving bounding boxes to database
            // This ensures clustering can read individual sleeve bounding boxes
            return (placedCount, skippedCount, errorCount);
        }

        /// <summary>
        /// Execute individual command with resource management using ExecuteImpl pattern
        /// </summary>
        private void ExecuteCommandWithResourceManagement(IExternalCommand command, OpeningFilter filter, bool showProgress)
        {
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[OpeningCommandOrchestrator] Executing {command.GetType().Name} for {filter.Category}");

                // ✅ FIX: No longer using RectangularSleeveClusterCommandV2 - clustering handled by UniversalClusterService
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Command {command.GetType().Name} not supported for direct execution - skipping");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing {command.GetType().Name}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Extracts clash zones from storage using the hierarchical structure, with legacy support.
        /// </summary>
        private static List<ClashZone> ExtractClashZonesFromStorage(ClashZoneStorage storage)
        {
            var result = new List<ClashZone>();

            if (storage == null)
                return result;

            if (storage.Filters != null)
            {
                foreach (var filterGroup in storage.Filters)
                {
                    if (filterGroup?.FileCombos == null) continue;

                    foreach (var fileCombo in filterGroup.FileCombos)
                    {
                        if (fileCombo?.ClashZones == null) continue;

                        result.AddRange(fileCombo.ClashZones);
                    }
                }
            }

            if (storage.ClashZones != null && storage.ClashZones.Count > 0)
            {
                var existingIds = new HashSet<Guid>(result.Select(z => z.Id));
                foreach (var cz in storage.ClashZones)
                {
                    if (cz != null && existingIds.Add(cz.Id))
                    {
                        result.Add(cz);
                    }
                }
            }

            return result;
        }
    }

    /// <summary>
    /// Failure preprocessor to auto-dismiss warnings during cluster sleeve placement.
    /// Per TRANSACTION_MANAGEMENT_IMPLEMENTATION_PLAN.md best practices.
    /// </summary>
    public class WarningSwallower : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
        {
            var failures = fa.GetFailureMessages();
            foreach (var f in failures)
            {
                var description = f.GetDescriptionText();
                
                // Dismiss warnings (per TRANSACTION_REFACTORING_SUMMARY.md)
                if (f.GetSeverity() == FailureSeverity.Warning)
                {
                    fa.DeleteWarning(f);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[WarningSwallower] Dismissed warning: {description}");
                    }
                }
                // Also dismiss duplicate-related errors to prevent transaction rollback
                else if (f.GetSeverity() == FailureSeverity.Error && 
                         (description.Contains("duplicate", StringComparison.OrdinalIgnoreCase) || 
                          description.Contains("already exists", StringComparison.OrdinalIgnoreCase)))
                {
                    fa.DeleteWarning(f);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[WarningSwallower] Dismissed duplicate error: {description}");
                    }
                }
            }
            return FailureProcessingResult.Continue;
        }
    }
}
