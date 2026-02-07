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
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Calculation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Workflow;
using JSE_RevitAddin_MEP_OPENINGS.Services.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
// using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration;

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

        // private readonly MarkPrefixSettings _markPrefixes;
        private readonly bool _forceDetectionMode;
        
        // ⚠️ CRITICAL: Crash-safe executor for timeout protection (5-minute limit per category)
        private readonly CrashSafeExecutor _crashSafeExecutor;
        
        // ✅ PATH 3: Store PATH 3 type flags for clustering (key: "FilterName_Category", value: (isValidated, isInvalidated, isNew))
        private JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor _performanceMonitor;
        
        // ✅ PATH 3: Store PATH 3 type flags for clustering (key: "FilterName_Category", value: (isValidated, isInvalidated, isNew))
        private Dictionary<string, (bool isValidated, bool isInvalidated, bool isNew)> _path3Flags 
            = new Dictionary<string, (bool isValidated, bool isInvalidated, bool isNew)>();



        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument, Dictionary<string, double> uiClearances = null, object markPrefixes = null, bool forceDetectionMode = false, IPerformanceMonitor performanceMonitor = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _uiClearances = uiClearances ?? new Dictionary<string, double>();
            // _markPrefixes = markPrefixes ?? new MarkPrefixSettings();
            _forceDetectionMode = forceDetectionMode;
            _performanceMonitor = performanceMonitor; // ✅ Use provided monitor or null (will create new one if needed)
            
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
                var buildTimestamp = System.IO.File.GetLastWriteTime(assembly.Location).ToString("yyyy-MM-dd HH:mm:ss");
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔨 BUILD TIMESTAMP: {buildTimestamp}\n");
            }
            catch { }

            // ✅ PERFORMANCE MONITORING: Use provided monitor or create new one
            if (_performanceMonitor == null)
            {
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                string logName = $"SleevePlacement_Batch_{timestamp}.log";
                _performanceMonitor = new PlacementPerformanceMonitor(logName);
            }
            int totalIndividualPlaced = 0;
            int totalClustersPlaced = 0;


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
                    totalIndividualPlaced += ExecuteDisciplineWithMemoryManagement(discipline.Key, discipline.Value, showProgress);
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

                // ✅ GLOBAL CLUSTERING & CLEANUP PHASE
                // This runs AFTER all individual sleeves from all disciplines/filters have been placed
                // and their REVIT coordinates have been updated in the database.
                totalClustersPlaced = ExecuteGlobalClusteringAndCleanup(filters, showProgress);

                // ✅ GENERATE REPORT AFTER ALL PHASES (Individual + Cluster)
                // This ensures the summary table includes data from the global clustering phase
                if (_performanceMonitor != null)
                {
                    _performanceMonitor.GenerateReport(totalIndividualPlaced, totalClustersPlaced);
                }

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
        /// Returns the category strings selected for this filter (one or many). Only these categories get placement/clustering; if not selected, no process.
        /// </summary>
        private static List<string> GetSelectedCategoryStringsForFilter(OpeningFilter filter)
        {
            if (filter == null) return new List<string>();
            if (filter.SelectedMepCategoryNames != null && filter.SelectedMepCategoryNames.Count > 0)
            {
                return filter.SelectedMepCategoryNames
                    .Select(c => Models.MepCategoryConstants.Normalize(c))
                    .Distinct()
                    .ToList();
            }
            string single = filter.Category switch
            {
                Models.MepCategory.Ducts => "Ducts",
                Models.MepCategory.DuctAccessories => "Duct Accessories",
                Models.MepCategory.Pipes => "Pipes",
                Models.MepCategory.CableTrays => "Cable Trays",
                _ => Models.MepCategoryConstants.Normalize(filter.Category.ToString())
            };
            return new List<string> { single };
        }

        /// <summary>
        /// Execute all filters for a discipline with memory management. Returns total individual sleeves placed.
        /// </summary>
        private int ExecuteDisciplineWithMemoryManagement(string discipline, List<OpeningFilter> filters, bool showProgress)
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

                // ✅ FIX: Deduplicate filters by name to prevent multiple placement calls for the same filter
                // Since LoadClashZonesForFilter loads ALL categories for a filter name,
                // we only need to call it ONCE per unique filter name.
                var uniqueFilters = orderedFilters.GroupBy(f => f.Name).Select(g => g.First()).ToList();

                int totalIndividualSleeves = 0;
                int totalClusters = 0;
                foreach (var filter in uniqueFilters)
                {
                    var commandSequence = GetCommandSequence(filter);
                    var (sleeves, clusters) = ExecuteCommandSequence(commandSequence, filter, showProgress);
                    totalIndividualSleeves += sleeves;
                    totalClusters += clusters;
                }

                // Report generation moved to main ExecuteMultipleFilters method to include global clustering totals

                // Force garbage collection after each discipline
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Discipline {discipline} completed, memory cleaned");
                }
                
                return totalIndividualSleeves;
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
        /// Execute a sequence of commands. Returns (individual sleeves placed, clusters placed) for report aggregation.
        /// </summary>
        private (int totalIndividualSleeves, int totalClusters) ExecuteCommandSequence(List<IExternalCommand> commands, OpeningFilter filter, bool showProgress)
        {
            var performanceMonitor = _performanceMonitor;
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

                // ✅ PERFORMANCE: Track individual sleeve placement (name reflects bulk vs sequential method)
                string placementMethod = OptimizationFlags.UseBulkIndividualSleevePlacement ? "Bulk Individual Sleeve Placement" : "Individual Sleeve Placement (Sequential)";
                using (var individualTracker = performanceMonitor.TrackOperation(placementMethod))
                {
                    var individualResult = ExecuteUniversalSleevePlacement(filter, showProgress, individualTracker);
                    totalIndividualSleeves = individualResult.placedCount;

                    // ✅ FIX: Set item count only when we have it from return; bulk path sets it inside to avoid closure/return lag
                    if (totalIndividualSleeves > 0)
                        individualTracker.SetItemCount(totalIndividualSleeves);

                    // ✅ PERFORMANCE LOGGING: Log actual performance metrics
                    if (!DeploymentConfiguration.DeploymentMode && totalIndividualSleeves > 0)
                    {
                        var avgTimePerSleeve = 288705.0 / totalIndividualSleeves; // Approximate from log
                        string methodLabel = OptimizationFlags.UseBulkIndividualSleevePlacement ? "Bulk Individual" : "Individual Sequential";
                        DebugLogger.Info($"[PERFORMANCE] {methodLabel} Sleeve Placement: {totalIndividualSleeves} sleeves in ~{288705 / 1000.0:F1}s = ~{avgTimePerSleeve / 1000.0:F2}s per sleeve (target: 0.02s)");
                        if (avgTimePerSleeve > 1000) // More than 1 second per sleeve
                        {
                            DebugLogger.Warning($"[PERFORMANCE] ⚠️ VERY SLOW: {avgTimePerSleeve / 1000.0:F2}s per sleeve is {avgTimePerSleeve / 20.0:F0}x slower than target (0.02s)");
                        }
                    }

                    // ✅ LOGGING: Wrap with SafeFileLogger
                    SafeFileLogger.SafeAppendText("orchestrator_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] AFTER ExecuteUniversalSleevePlacement, ABOUT TO CALL ExecuteClusteringForCategory for filter={filter.Name}, category={filter.Category}\n");

                    // ✅ CLUSTERING REMOVED: Now handled globally in ExecuteGlobalClusteringAndCleanup
                    // to ensure cross-category overlaps are correctly identified after all coordinates are saved.

                    string categoryString = filter.Category.ToString();
                    string path3Key = $"{filter.Name}_{categoryString}";
                    bool isPath3Validated = false;
                    bool isPath3Invalidated = false;

                    if (_path3Flags.TryGetValue(path3Key, out var path3Flags))
                    {
                        // Path 3 logic (previously had clustering here, now handled above)
                        // ✅ COMBO FLAG: Reset isNew to false after full sequence completes (individual + cluster)
                        // This marks the filter+category combo as "used" so next time it will use PATH 1 (Replay)
                        _path3Flags[path3Key] = (path3Flags.isValidated, path3Flags.isInvalidated, false);
                        isPath3Validated = true;
                    }
                }
                return (totalIndividualSleeves, totalClusters);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log",
                    $"[{DateTime.Now:HH:mm:ss}] ERROR in ExecuteCommandSequence for {filter.Name}: {ex.Message}\n");
                _performanceMonitor?.GenerateReport(totalIndividualSleeves, totalClusters);
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
                // ✅ BULK PLACEMENT FIX: Load ALL categories for this filter with ReadyForPlacement=1 in ONE DB call
                // ReadyForPlacementFlag=1 is set during refresh (SetReadyForPlacementForUnresolvedZonesInSectionBox) so we can place in one shot
                var categoriesToLoad = GetCategoriesWithReadyForPlacement(filter);

                if (categoriesToLoad.Count == 0)
                {
                    string categoryName = filter.Category switch
                    {
                        Models.MepCategory.Ducts => "Ducts",
                        Models.MepCategory.DuctAccessories => "Duct Accessories",
                        Models.MepCategory.Pipes => "Pipes",
                        Models.MepCategory.CableTrays => "Cable Trays",
                        _ => "Ducts"
                    };
                    categoriesToLoad.Add(categoryName);
                }

                using (var context = new SleeveDbContext(_document, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[OpeningCommandOrchestrator][SQLite] {msg}");
                }))
                {
                    var repository = new ClashZoneRepository(context);
                    List<ClashZone> zones;
                    List<ClashZone> eligibleZones;

                    // Single DB call: load all categories with ReadyForPlacementFlag=1 in one query
                    using (var loadTracker = _performanceMonitor?.TrackOperation("Step 1: LOADING FROM DB"))
                    {
                        zones = repository.GetClashZonesByFilterAllCategories(filter.Name, categoriesToLoad, unresolvedOnly: false, readyForPlacementOnly: true) ?? new List<ClashZone>();
                        eligibleZones = zones.Where(cz => cz != null && (
                            cz.IsCurrentClash ||
                            (!cz.IsResolvedFlag && !cz.IsClusterResolvedFlag && !cz.IsCombinedResolved)
                        )).ToList();
                        loadTracker?.SetItemCount(eligibleZones.Count);
                    }

                    if (!DeploymentConfiguration.DeploymentMode && zones.Count > 0)
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

                    if (eligibleZones.Count == 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found in database for filter '{filter.Name}' (categories: {string.Join(", ", categoriesToLoad)})");
                        return new List<ClashZone>();
                    }

                    // Deduplicate by GUID (same zone can appear in multiple categories/file combos)
                    int originalCount = eligibleZones.Count;
                    var allZones = eligibleZones
                        .Where(z => !string.IsNullOrEmpty(z.ClashZoneGuid))
                        .GroupBy(z => z.ClashZoneGuid.ToUpper())
                        .Select(g => g.First())
                        .ToList();

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        if (allZones.Count < originalCount)
                            DebugLogger.Warning($"[OpeningCommandOrchestrator] ⚠️ DEDUPLICATED: {originalCount} -> {allZones.Count} zones (Removed {originalCount - allZones.Count} duplicates)");
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ TOTAL: Loaded {allZones.Count} clash zones from {categoriesToLoad.Count} category(ies) in ONE DB call for filter '{filter.Name}'\n");
                    }

                    return allZones;
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error loading clash zones for filter {filter.Name}: {ex.Message}");
                return new List<ClashZone>();
            }
        }

        /// <summary>
        /// ✅ BULK PLACEMENT FIX: Get all categories that have ReadyForPlacement=1 for a filter
        /// </summary>
        private List<string> GetCategoriesWithReadyForPlacement(OpeningFilter filter)
        {
            var categories = new List<string>();
            
            try
            {
                using (var context = new SleeveDbContext(_document, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[OpeningCommandOrchestrator][SQLite] {msg}");
                }))
                {
                    var repository = new ClashZoneRepository(context);
                    
                    // Query distinct categories that have ReadyForPlacement=1 for this filter
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        cmd.CommandText = @"
                            SELECT DISTINCT f.Category
                            FROM Filters f
                            INNER JOIN FileCombos fc ON f.FilterId = fc.FilterId
                            INNER JOIN ClashZones cz ON fc.ComboId = cz.ComboId
                            WHERE f.FilterName = @FilterName
                              AND cz.ReadyForPlacementFlag = 1
                              AND cz.IsCombinedResolved = 0";
                        
                        cmd.Parameters.AddWithValue("@FilterName", FilterNameHelper.NormalizeBaseName(filter.Name, filter.Name, ""));
                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (!reader.IsDBNull(0))
                                {
                                    var category = reader.GetString(0);
                                    if (!string.IsNullOrWhiteSpace(category))
                                        categories.Add(category);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Error getting categories with ReadyForPlacement: {ex.Message}");
                }
            }
            
            return categories;
        }

        private List<ClashZone> LoadClashZonesFromDatabase(OpeningFilter filter, string categoryName, bool readyForPlacementOnly = false)
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
                    var repository = new ClashZoneRepository(context);

                    List<ClashZone> eligibleZones = new List<ClashZone>();

                    // Step 1: LOADING FROM DB
                    List<ClashZone> zones = null;
                    using (var loadTracker = _performanceMonitor?.TrackOperation("Step 1: LOADING FROM DB"))
                    {
                        // ✅ BULK PLACEMENT FIX: Use readyForPlacementOnly parameter
                        zones = repository.GetClashZonesByFilter(filter.Name, categoryName, unresolvedOnly: false, readyForPlacementOnly: readyForPlacementOnly) ?? new List<ClashZone>();

                        // ✅ SIMPLIFIED ELIGIBILITY: Zone is eligible if:
                        // 1. IsCurrentClash = 1 (session flag set during refresh), OR
                        // 2. Zone is unresolved (deleted sleeve scenario: IsResolved=0, IsClusterResolved=0, IsCombinedResolved=0)
                        eligibleZones = zones.Where(cz => cz != null && (
                            cz.IsCurrentClash ||  
                            (!cz.IsResolvedFlag && !cz.IsClusterResolvedFlag && !cz.IsCombinedResolved)
                        )).ToList();

                        loadTracker?.SetItemCount(eligibleZones.Count);
                    }

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
        private (int placedCount, int skippedCount, int errorCount) ExecuteUniversalSleevePlacement(OpeningFilter filter, bool showProgress, IOperationTracker parentTracker = null)
        {
            // Declare variables outside lambda for use after timeout execution
            string xmlFilePath = GetXmlFilePathForFilter(filter);
            var tracePath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
            var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
            
            // ✅ PERFORMANCE: Store counts outside lambda for access after timeout execution
            int placedCount = 0;
            int skippedCount = 0;
            int errorCount = 0;
            
            // ✅ FIX: Store placed zones (GUID + ElementId) for Step 5 corner extraction
            // This avoids SQLite WAL visibility issues where GetPlacedClashZones() returns 0 items
            List<(Guid ZoneGuid, int ElementId)> placedZonesForExtraction = new List<(Guid, int)>();
            
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

                    // ✅ PERFORMANCE: Track loading clash zones
                    List<ClashZone> clashZones;
                    using (parentTracker?.TrackSubOperation("Load Clash Zones"))
                    {
                        clashZones = LoadClashZonesForFilter(filter);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[OpeningCommandOrchestrator] Loaded {clashZones.Count} clash zones for {filter.Category}");
                        }
                    }

                    // ✅ PERFORMANCE: Track pre-filtering
                    using (parentTracker?.TrackSubOperation("Pre-Filter Zones (Wall Thickness)"))
                    {
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
                    }
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] LOAD_XML: zones={clashZones.Count}\n");
                        }
                    } catch { }
                    
                    // ✅ FLAG-MANAGER INTEGRATION: Create flag manager for this run
                    var flagManager = JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement.FlagManagerFactory.CreateAdapter(_document);
                    
                    if (clashZones.Count > 0)
                    {
                        // ✅ PERFORMANCE: Track flag syncing
                        using (parentTracker?.TrackSubOperation("Sync Flags from Database"))
                        {
                            // ✅ SYNC FLAGS: Ensure in-memory zones reflect database state before placement
                            try
                            {
                                string categoryStringSync = filter.Category.ToString(); // Or normalize if needed
                                flagManager.SyncFlagsFromGlobal(clashZones, categoryStringSync);
                            }
                            catch (Exception syncEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Error syncing flags from DB: {syncEx.Message}");
                            }
                        }

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
                        
                         // ✅ PATH 3 INVALIDATED: Detect zones where MEP moved
                        // These zones have sleeves (SleeveInstanceId > 0)
                        var invalidatedZones = clashZones.Where(cz => cz.SleeveInstanceId > 0).ToList();
                        
                        // ✅ VALIDATED ZONES: Zones without sleeves that need placement
                        var validatedZones = clashZones.Where(cz => cz.SleeveInstanceId <= 0 && !cz.IsResolvedFlag && !cz.IsClusterResolvedFlag && cz.ClusterSleeveInstanceId <= 0).ToList();
                        
                        // ✅ PERFORMANCE: Track deletion of invalidated sleeves
                        if (invalidatedZones.Count > 0)
                        {
                            using (parentTracker?.TrackSubOperation("Delete Invalidated Sleeves"))
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[OpeningCommandOrchestrator] Deleting old sleeves for {invalidatedZones.Count} invalidated zones...");
                                
                                using (var t = new Transaction(_document, "Delete Invalidated Sleeves"))
                                {
                                    t.Start();
                                    int deletedCount = 0;
                                    foreach (var zone in invalidatedZones)
                                    {
                                        if (zone.SleeveInstanceId > 0)
                                        {
                                            try
                                            {
                                                // Safety: don't delete cluster sleeves if recently placed
                                                if (FlagManagement.FlagManagerProtectionHelper.IsRecentlyPlacedClusterSleeve(zone.SleeveInstanceId)) continue;

                                                var id = new ElementId(zone.SleeveInstanceId);
                                                if (_document.GetElement(id) != null)
                                                {
                                                    _document.Delete(id);
                                                    deletedCount++;
                                                }
                                                // ✅ CRITICAL: Reset the ID on the object so the placer knows it's a new placement
                                                zone.SleeveInstanceId = 0;
                                                zone.IsResolvedFlag = false;
                                            }
                                            catch { /* Ignore */ }
                                        }
                                    }
                                    t.Commit();
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[OpeningCommandOrchestrator] Deleted {deletedCount} old sleeves.");
                                }
                            }
                        }

                        // ✅ CONSOLIDATED LIST: Place both previously-empty and previously-invalidated zones
                        var zonesToPlace = validatedZones.Concat(invalidatedZones).ToList();
                        
                        // ✅ DIAGNOSTIC: Log zones being processed to identify duplicates
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            var zoneIds = zonesToPlace.Select(z => z.Id.ToString()).ToList();
                            var zoneCategories = zonesToPlace.GroupBy(z => z.MepElementCategory ?? "Unknown").ToDictionary(g => g.Key, g => g.Count());
                            SafeFileLogger.SafeAppendText("placement_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] [BULK-PLACEMENT] 📊 Filter '{filter.Name}': Processing {zonesToPlace.Count} zones to place. Categories: {string.Join(", ", zoneCategories.Select(kvp => $"{kvp.Key}={kvp.Value}"))}\n");
                            SafeFileLogger.SafeAppendText("placement_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] [BULK-PLACEMENT] 📊 Zone IDs: {string.Join(", ", zoneIds.Take(10))}{(zoneIds.Count > 10 ? "..." : "")}\n");
                        }
                        
                        if (zonesToPlace.Count > 0)
                        {
                            if (OptimizationFlags.UseBulkIndividualSleevePlacement)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🚀 [BULK-PLACEMENT] Routing to Consolidated Bulk Path (ALL categories together)\n");

                                // ═══════════════════════════════════════════════════════════════════════════════
                                // OPERATION 1: Calculate (pre-plan) and save to DB. No Revit API yet.
                                // ═══════════════════════════════════════════════════════════════════════════════
                                var projectFiltersDir = ProjectPathService.GetFiltersDirectory(_document);
                                var conditionsService = new ConditionsService(_document, projectFiltersDir, msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg); });
                                
                                var allBulkTaskItems = new List<(ClashZone Zone, SleevePlacementPlanningDto Plan)>();
                                
                                var zonesByCategory = zonesToPlace
                                    .GroupBy(z => z.MepElementCategory ?? "Ducts")
                                    .ToList();
                                
                                using (parentTracker?.TrackSubOperation("Operation 1: Calculate and save to DB"))
                                {
                                // Sub-op: plan placement for all categories
                                using (parentTracker?.TrackSubOperation("Plan placement (all categories)"))
                                {
                                    foreach (var categoryGroup in zonesByCategory)
                                    {
                                    var category = categoryGroup.Key;
                                    var categoryZones = categoryGroup.ToList();
                                    
                                    // Get category string for conditions lookup (use different variable names to avoid conflicts)
                                    string catString = category switch
                                    {
                                        "Ducts" => "Ducts",
                                        "Duct Accessories" => "Duct Accessories",
                                        "Pipes" => "Pipes",
                                        "Cable Trays" or "CableTrays" => "Cable Trays",
                                        _ => "Ducts"
                                    };
                                    
                                    string catName = category switch
                                    {
                                        "Ducts" => "ducts",
                                        "Duct Accessories" => "duct_accessories",
                                        "Pipes" => "pipes",
                                        "Cable Trays" or "CableTrays" => "cable_trays",
                                        _ => "ducts"
                                    };
                                    string filterNameForCategory = $"{filter.Name}_{catName}.xml";
                                    
                                    // Load conditions for this category
                                    var conditionsKey = $"{filterNameForCategory}_{catString}";
                                    var categoryConditions = conditionsService.LoadConditions(conditionsKey) ?? new OpeningConditions { FilterName = filterNameForCategory, Category = catString };

                                    // ✅ PATH 1 FAST PATH: Check if zones already have PLACED data in DB
                                    // Only use Path 1 if sleeves were actually placed before (SleeveInstanceId > 0)
                                    // This prevents using Path 1 with fresh DB that only has refresh/clash detection data
                                    var zonesWithData = categoryZones.Where(z => 
                                        z.SleeveInstanceId > 0 && // ✅ CRITICAL: Must have been placed before
                                        (z.CalculatedSleeveWidth > 0 || z.SleeveWidth > 0) &&
                                        (z.CalculatedSleeveHeight > 0 || z.SleeveHeight > 0) &&
                                        (z.SleevePlacementPointX != 0 || z.CalculatedPlacementX != 0) &&
                                        !string.IsNullOrEmpty(z.SleeveFamilyName ?? z.CalculatedFamilyName)
                                    ).ToList();
                                    
                                    if (zonesWithData.Count == categoryZones.Count)
                                    {
                                        // ✅ ALL zones have placement data - skip calculation, use existing data
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[BULK-PLACEMENT] ✅ PATH 1 FAST PATH: All {categoryZones.Count} zones in category '{category}' have placement data - skipping calculation, using existing data");
                                        }
                                        
                                        // Create plans from existing zone data (no calculation needed)
                                        foreach (var zone in categoryZones)
                                        {
                                            double width = zone.CalculatedSleeveWidth > 0 ? zone.CalculatedSleeveWidth : zone.SleeveWidth;
                                            double height = zone.CalculatedSleeveHeight > 0 ? zone.CalculatedSleeveHeight : zone.SleeveHeight;
                                            // ✅ ROUNDING: Apply user rounding (RoundAlwaysUp, RoundingValue) so e.g. 710 → 750
                                            var (roundedW, roundedH) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(width, height);
                                            width = roundedW; height = roundedH;
                                            zone.CalculatedSleeveWidth = width; zone.CalculatedSleeveHeight = height;
                                            zone.SleeveWidth = width; zone.SleeveHeight = height;
                                            
                                            // Depth = structural thickness for all (Floor, Wall, Framing)
                                            double depth = zone.StructuralElementThickness > 0.001 ? zone.StructuralElementThickness : (zone.CalculatedSleeveDepth > 0 ? zone.CalculatedSleeveDepth : zone.SleeveDepth);
                                            zone.CalculatedSleeveDepth = depth > 0 ? depth : zone.CalculatedSleeveDepth;
                                            
                                            double placementX = zone.SleevePlacementPointX != 0 ? zone.SleevePlacementPointX : zone.CalculatedPlacementX;
                                            double placementY = zone.SleevePlacementPointY != 0 ? zone.SleevePlacementPointY : zone.CalculatedPlacementY;
                                            double placementZ = zone.SleevePlacementPointZ != 0 ? zone.SleevePlacementPointZ : zone.CalculatedPlacementZ;
                                            
                                            // ✅ ROTATION FIX: Use SleeveRotationService to properly handle circular elements
                                            // Circular elements on floors should have 0 rotation, even if MepElementRotationAngle is set
                                            var rotationService = new Placement.SleeveRotationService();
                                            double rotationRadians = rotationService.DetermineRotation(zone);
                                            
                                            var plan = new Placement.SleevePlacementPlanningDto(
                                                clashZoneId: zone.Id,
                                                hostType: zone.StructuralElementType ?? "Unknown",
                                                rawSleeveSizeFt: Math.Max(width, height),
                                                insulationThicknessFt: 0,
                                                targetWidthFt: width,
                                                targetHeightFt: height,
                                                clearanceFt: 0,
                                                requiredDepthFt: depth,
                                                rotationRadians: rotationRadians,
                                                risk: Placement.ClearanceRiskClassification.None,
                                                shouldSkip: false,
                                                skipReason: "",
                                                logTrace: "PATH 1 Fast Path - Using existing data",
                                                familyName: zone.SleeveFamilyName ?? zone.CalculatedFamilyName ?? "",
                                                placementPoint: new XYZ(placementX, placementY, placementZ),
                                                isCircular: zone.SleeveDiameter > 0
                                            );
                                            
                                            allBulkTaskItems.Add((zone, plan));
                                        }
                                    }
                                    else
                                    {
                                        // ✅ Some zones missing data - run calculation for all zones
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            DebugLogger.Info($"[BULK-PLACEMENT] ⚠️ PATH 2: {categoryZones.Count - zonesWithData.Count} zones missing placement data - running calculation");
                                        }
                                        
                                        // ✅ NEW: PRE-CALCULATE PLACEMENT DATA (Sizes, Family Names, Skip Rules)
                                        // This ensures BulkPlacementService has all necessary metadata and follows the SAME rules as Standard path
                                        var planner = new ParallelSleevePlacementPlanner(categoryConditions);
                                        var planningResult = planner.Plan(categoryZones);
                                    
                                        // Map planned data back to Task Items (Zone + Plan) for THIS category
                                        var plannedMap = planningResult.Items.ToDictionary(i => i.ClashZoneId);
                                        
                                        foreach (var zone in categoryZones)
                                        {
                                            if (plannedMap.TryGetValue(zone.Id, out var plan))
                                            {
                                                if (plan.ShouldSkip) continue;
                                                
                                                // ✅ CRITICAL FIX: Map calculated values back to ClashZone object for persistence
                                                // The Planner calculates these but doesn't mutate the zone. We must apply them 
                                                // so they are saved to the DB later in Step 6 (BatchUpdateSleevePlacementData).
                                                
                                                // 1. Dimensions
                                                zone.CalculatedSleeveWidth = plan.TargetWidthFt;
                                                zone.CalculatedSleeveHeight = plan.TargetHeightFt;
                                                // Depth = structural thickness for all (Floor, Wall, Framing)
                                                double depthForSleeve = zone.StructuralElementThickness > 0.001 ? zone.StructuralElementThickness : (plan.TargetDepthFt > 0 ? plan.TargetDepthFt : 0);
                                                zone.CalculatedSleeveDepth = depthForSleeve > 0 ? depthForSleeve : zone.StructuralElementThickness;
                                                
                                                // ✅ USER FIX: Also map to "Actual" DB columns so clustering works correctly
                                                // Bulk placement forces these dimensions, so we can trust them as "Actual"
                                                zone.SleeveWidth = plan.TargetWidthFt;
                                                zone.SleeveHeight = plan.TargetHeightFt;
                                                zone.SleeveDiameter = plan.TargetDiameterFt;
                                                
                                                // 2. Family Name
                                                zone.SleeveFamilyName = plan.SleeveFamilyName;
                                                
                                                // 3. Placement Point & Rotation
                                                if (plan.PlacementPoint != null)
                                                {
                                                    zone.SleevePlacementPointX = plan.PlacementPoint.X;
                                                    zone.SleevePlacementPointY = plan.PlacementPoint.Y;
                                                    zone.SleevePlacementPointZ = plan.PlacementPoint.Z;
                                                    
                                                    // Ensure Calculated properties are also set for the Pre-Save
                                                    zone.CalculatedPlacementX = plan.PlacementPoint.X;
                                                    zone.CalculatedPlacementY = plan.PlacementPoint.Y;
                                                    zone.CalculatedPlacementZ = plan.PlacementPoint.Z;
                                                    zone.CalculatedFamilyName = plan.SleeveFamilyName;
                                                    zone.CalculatedRotation = plan.RotationRadians;
                                                }
                                                zone.MepElementRotationAngle = plan.RotationRadians;

                                                // ✅ BULK PLACEMENT FIX: Add to combined list (all categories together)
                                                allBulkTaskItems.Add((zone, plan));
                                            }
                                        }
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[BULK-PLANNING] Category '{category}': Planned {plannedMap.Count} items (Total so far: {allBulkTaskItems.Count} across all categories)");
                                    }
                                }
                                }
                                 
                                // ✅ SAFETY STEP: PRE-SAVE CALCULATED DATA TO DATABASE (ALL categories together)
                                // This ensures that if Revit crashes during the heavy bulk placement (1000+ items),
                                // the dimensions and positions are safely stored in DB and not lost in memory.
                                // ✅ FIX: Always pre-save zones that HAVE calculated data (CalculatedSleeveWidth > 0).
                                // Previous logic used "zonesNeedingSave" = zones MISSING data, which excluded PATH 2 zones
                                // (they had just been populated by planner) - causing CalculatedSleeveWidth to never persist.
                                var zonesToSave = allBulkTaskItems.Select(x => x.Zone).ToList();

                                // ✅ FIX: Include pipe sleeves (diameter-only) and any sleeve
                                // that has *any* meaningful calculated size, not just width+height.
                                // Previously we required BOTH width and height > 0, which excluded pipes.
                                var zonesWithCalculatedData = zonesToSave.Where(z =>
                                    !string.IsNullOrWhiteSpace(z.ClashZoneGuid) &&
                                    (
                                        (z.CalculatedSleeveWidth > 0.001 || z.SleeveWidth > 0.001) ||
                                        (z.CalculatedSleeveHeight > 0.001 || z.SleeveHeight > 0.001) ||
                                        (z.CalculatedSleeveDiameter > 0.001 || z.SleeveDiameter > 0.001)
                                    ) &&
                                    !string.IsNullOrEmpty(z.SleeveFamilyName ?? z.CalculatedFamilyName)
                                ).ToList();

                                // ✅ Diagnostic: log to placement_debug.log so user can verify pre-save input
                                var sampleZ = zonesWithCalculatedData.FirstOrDefault();
                                SafeFileLogger.SafeAppendText("placement_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PRE-SAVE] zonesToSave={zonesToSave.Count}, zonesWithCalculatedData={zonesWithCalculatedData.Count}, sample ClashZoneId={sampleZ?.ClashZoneId}, W={sampleZ?.SleeveWidth}, Guid={sampleZ?.ClashZoneGuid}\n");

                                if (zonesWithCalculatedData.Count > 0)
                                {
                                    using (parentTracker?.TrackSubOperation("Pre-Save Calculated Data to DB"))
                                    {
                                        try
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info($"[BULK-SAFETY] Pre-saving {zonesWithCalculatedData.Count} zones (CalculatedSleeveWidth/Height) to DB...");
                                            
                                            using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(_document))
                                            {
                                                var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClashZoneRepository(dbContext);
                                                repo.BatchUpdateCalculatedData(zonesWithCalculatedData);
                                                if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info($"[BULK-SAFETY] ✅ Pre-save complete.");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            SafeFileLogger.SafeAppendText("placement_debug.log", $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PRE-SAVE] ❌ Pre-save failed: {ex.Message}\n");
                                            if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Error($"[BULK-SAFETY] ❌ Pre-save failed: {ex.Message}");
                                            // Non-fatal, continue with placement
                                        }
                                    }
                                }

                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[BULK-PLANNING] ✅ Planning complete for ALL categories. Total items to place: {allBulkTaskItems.Count}");
                                } // end Operation 1: Calculate and save to DB

                                // ═══════════════════════════════════════════════════════════════════════════════
                                // OPERATION 2: Bulk placement (Revit API) using data saved in DB from Operation 1.
                                // ═══════════════════════════════════════════════════════════════════════════════
                                 var paramService = new SleeveParameterService(_document);
                                 var bulkService = new BulkPlacementService(
                                     _document, 
                                     msg => 
                                     {
                                         if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg);
                                         SafeFileLogger.SafeAppendText("bulk_placement_trace.log", $"{DateTime.Now:HH:mm:ss} {msg}\n");
                                     },
                                     _performanceMonitor,
                                     paramService);
                                BulkPlacementResult bulkResult = null;
                                 SafeFileLogger.SafeAppendText("placement_debug.log", $"\n[{DateTime.Now:HH:mm:ss}] [BULK-PLACEMENT-START] allBulkTaskItems.Count = {allBulkTaskItems.Count}\n");

                                 using (parentTracker?.TrackSubOperation("Check for Duplicate Zones"))
                                 {
                                     // Check for duplicates by ClashZoneGuid
                                     var duplicateZones = allBulkTaskItems
                                         .GroupBy(x => x.Zone.ClashZoneGuid)
                                         .Where(g => !string.IsNullOrEmpty(g.Key) && g.Count() > 1)
                                         .ToList();

                                     if (duplicateZones.Count > 0)
                                     {
                                         SafeFileLogger.SafeAppendText("placement_debug.log", 
                                             $"[{DateTime.Now:HH:mm:ss}] [WARNING] Found {duplicateZones.Count} duplicate zones in allBulkTaskItems!\n");
                                         foreach (var dup in duplicateZones)
                                         {
                                             SafeFileLogger.SafeAppendText("placement_debug.log", 
                                                 $"  - Zone {dup.Key} appears {dup.Count()} times\n");
                                         }
                                     }
                                 }

                                 List<(ClashZone Zone, SleevePlacementPlanningDto Plan)> bulkTaskItemsDeduped;
                                 using (parentTracker?.TrackSubOperation("Dedupe by Placement Location"))
                                 {
                                     // ✅ FIX: Dedupe by placement location to avoid "identical instances in the same place" (Revit warning).
                                     // Two zones with same XYZ (different GUIDs) would create two sleeves → two nested Family2 at same spot.
                                     // Keep first per (family + rounded XYZ) so we only place one instance per location.
                                     const double locationToleranceFt = 0.0328; // ~10mm - catches near-duplicates from floating point / different calc paths
                                     Func<double, double> round = v => Math.Round(v / locationToleranceFt) * locationToleranceFt;
                                     bulkTaskItemsDeduped = allBulkTaskItems
                                         .GroupBy(x =>
                                         {
                                             var pt = x.Plan.PlacementPoint ?? new XYZ(x.Zone.IntersectionPointX, x.Zone.IntersectionPointY, x.Zone.IntersectionPointZ);
                                             return ($"{x.Plan.SleeveFamilyName ?? x.Zone.SleeveFamilyName}", round(pt.X), round(pt.Y), round(pt.Z));
                                         })
                                         .Select(g => g.First())
                                         .ToList();
                                     if (bulkTaskItemsDeduped.Count < allBulkTaskItems.Count)
                                     {
                                         var dropped = allBulkTaskItems.Count - bulkTaskItemsDeduped.Count;
                                         SafeFileLogger.SafeAppendText("placement_debug.log",
                                             $"[{DateTime.Now:HH:mm:ss}] [DEDUPE-LOCATION] Removed {dropped} items with same placement point (kept {bulkTaskItemsDeduped.Count}) to avoid Revit duplicate warning.\n");
                                         if (!DeploymentConfiguration.DeploymentMode)
                                             DebugLogger.Info($"[BULK-PLACEMENT] Deduped by location: {allBulkTaskItems.Count} → {bulkTaskItemsDeduped.Count} (dropped {dropped} to prevent identical instances)");
                                     }
                                 }

                                using (var placeTracker = parentTracker?.TrackSubOperation("Operation 2: Bulk placement (Revit API)"))
                                {
                                    if (OptimizationFlags.UseTwoTransactionBulkPlacement)
                                    {
                                        // ✅ TWO-TRANSACTION FLOW (Monday-style): placement commit first, then parameters. Faster (~25 sleeves/s). Rollback: set UseTwoTransactionBulkPlacement = false.
                                        using (var t = new Transaction(_document, $"Bulk Place {filter.Name}"))
                                        {
                                            t.Start();
                                            
                                            using (placeTracker?.TrackSubOperation("ExecuteBulkPlacement (Revit API)"))
                                            {
                                                bulkResult = bulkService.ExecuteBulkPlacement(_document, bulkTaskItemsDeduped);
                                            }
                                            using (placeTracker?.TrackSubOperation("Transaction Commit (Placement Only)"))
                                            {
                                                t.Commit();
                                            }
                                        }
                                        if (bulkResult != null && bulkResult.OverallSuccess && bulkResult.PlacedCount > 0)
                                        {
                                            // ✅ PERFORMANCE OPTIMIZATION: Transaction required for parameter setting, but optimized
                                            // ISSUE: Parameters revert to default without transaction wrapper
                                            // FIX: Keep transaction but remove tracking overhead from commit (reduces validation time)
                                            try 
                                            {
                                                using (var tParams = new Transaction(_document, "Flush Parameters"))
                                                {
                                                    tParams.Start();
                                                    // Transaction optimized for parameter flushing
                                                    using (placeTracker?.TrackSubOperation("Flush Deferred Parameters"))
                                                    {
                                                        paramService.FlushDeferredParameters(clearList: true, context: "AfterPlacement");
                                                    }
                                                    
                                                    var commitTimer = System.Diagnostics.Stopwatch.StartNew();
                                                    using (placeTracker?.TrackSubOperation("Transaction Commit (Params)"))
                                                    {
                                                        tParams.Commit();
                                                    }
                                                    commitTimer.Stop();
                                                    SafeFileLogger.SafeAppendText("performance.log", $"[{DateTime.Now:HH:mm:ss.fff}] [PERF-TELEMETRY] ParamCommit={commitTimer.ElapsedMilliseconds}ms\n");
                                                }
                                                
                                                var regenTimer = System.Diagnostics.Stopwatch.StartNew();
                                                // ✅ OPTIMIZATION: Regenerate OUTSIDE transaction (was inside before)
                                                using (placeTracker?.TrackSubOperation("Regenerate (Total)"))
                                                {
                                                    _document.Regenerate();
                                                }
                                                regenTimer.Stop();
                                                SafeFileLogger.SafeAppendText("performance.log", $"[{DateTime.Now:HH:mm:ss.fff}] [PERF-TELEMETRY] Regenerate={regenTimer.ElapsedMilliseconds}ms\n");
                                            }
                                            catch (Exception paramEx)
                                            {
                                                SafeFileLogger.SafeAppendText("placement_debug.log", 
                                                    $"[{DateTime.Now:HH:mm:ss.fff}] [BATCH-PARAM-ERROR] ⚠️ Parameter flushing failed, but sleeves are already placed. Proceeding to DB persistence: {paramEx.Message}\n");
                                                if (!DeploymentConfiguration.DeploymentMode) 
                                                    DebugLogger.Warning($"[BATCH-PLACEMENT] Parameter flush failed (non-fatal for persistence): {paramEx.Message}");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // ROLLBACK PATH: Single transaction (placement + flush). Set UseTwoTransactionBulkPlacement = true to re-enable fast path.
                                        using (var t = new Transaction(_document, $"Bulk Place {filter.Name}"))
                                        {
                                            t.Start();
                                            using (placeTracker?.TrackSubOperation("ExecuteBulkPlacement (Revit API)"))
                                            {
                                                bulkResult = bulkService.ExecuteBulkPlacement(_document, bulkTaskItemsDeduped);
                                            }
                                            if (bulkResult != null && bulkResult.OverallSuccess && bulkResult.PlacedCount > 0)
                                            {
                                                using (placeTracker?.TrackSubOperation("Flush Deferred Parameters"))
                                                {
                                                    paramService.FlushDeferredParameters(clearList: true, context: "AfterPlacement");
                                                }
                                            }
                                            using (placeTracker?.TrackSubOperation("Transaction Commit (Placement + Parameters)"))
                                            {
                                                t.Commit();
                                            }
                                        }
                                    }

                                    // ✅ Set count immediately after commit so final log/report show correct Placed even if post-placement throws
                                    if (bulkResult != null && bulkResult.PlacedCount > 0)
                                    {
                                        placedCount = bulkResult.PlacedCount;
                                        parentTracker?.SetItemCount(bulkResult.PlacedCount);
                                    }

                                    // ✅ STEP 2: POST-PLACEMENT PROCESSING (Metadata, Geometry) - no further Revit transactions
                                    if (bulkResult != null && bulkResult.OverallSuccess && bulkResult.PlacedCount > 0)
                                    {
                                        SafeFileLogger.SafeAppendText("placement_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss.fff}] [BULK-POST-PLACEMENT] UpdateZonesFromElements + DB persist (PlacedCount={bulkResult.PlacedCount}, TwoTx={OptimizationFlags.UseTwoTransactionBulkPlacement})\n");
                                        try
                                        {
                                            using (placeTracker?.TrackSubOperation("Update Zones From Elements"))
                                            {
                                                bulkService.UpdateZonesFromElements(_document, bulkResult.PlacedItems);
                                            }
                                        }
                                        catch (Exception updateEx)
                                        {
                                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                                $"[{DateTime.Now:HH:mm:ss.fff}] [BULK-POST-PLACEMENT] UpdateZonesFromElements failed (continuing to persist): {updateEx.Message}\n{updateEx.StackTrace}\n");
                                        }

                                        // 2e. OUTSIDE ALL TRANSACTIONS: Database persistence (always run when PlacedCount>0)
                                        try
                                        {
                                            List<ClashZone> placedZones;
                                            using (placeTracker?.TrackSubOperation("Prepare Placed Zones List"))
                                            {
                                                placedZones = bulkResult.PlacedItems.Select(p => p.Zone).ToList();
                                                SafeFileLogger.SafeAppendText("placement_debug.log",
                                                    $"[{DateTime.Now:HH:mm:ss.fff}] [BULK-PERSIST-START] placedZones={placedZones.Count}, SleeveInstanceId>0={placedZones.Count(z => z.SleeveInstanceId > 0)}\n");
                                            }
                                             
                                             // ✅ Update flags in database after bulk placement
                                             using (placeTracker?.TrackSubOperation("Update Flags After Placement"))
                                             {
                                                 if (placedZones.Any())
                                                 {
                                                     var flagUpdates = bulkResult.PlacedItems
                                                         .Select(item => (item.Zone.Id, item.ElementId.IntegerValue, false))
                                                         .ToList();
                                                     
                                                     flagManager.UpdateFlagsAfterPlacement(flagUpdates);
                                                     // Set in-memory so BatchUpdateSleevePlacementData persists correct flag
                                                     foreach (var z in placedZones) z.IsResolvedFlag = true;
                                                 }
                                             }

                                             using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(_document))
                                            {
                                                var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClashZoneRepository(dbContext);
                                                
                                                int filterId = -1;
                                                using (placeTracker?.TrackSubOperation("Lookup Filter ID"))
                                                {
                                                    string categoryForLookup = filter.Category switch
                                                    {
                                                        Models.MepCategory.Ducts => "Ducts",
                                                        Models.MepCategory.DuctAccessories => "Duct Accessories",
                                                        Models.MepCategory.Pipes => "Pipes",
                                                        Models.MepCategory.CableTrays => "Cable Trays",
                                                        _ => filter.Category.ToString()
                                                    };
                                                    var filterLookupService = new JSE_RevitAddin_MEP_OPENINGS.Services.Filters.FilterLookupService(dbContext);
                                                    filterId = filterLookupService.GetFilterId(filter.Name, categoryForLookup);
                                                }

                                                // Persist sleeve data: corners, bounding box, flags, dimensions, placement
                                                List<ClashZone> zonesWithIds;
                                                using (placeTracker?.TrackSubOperation("Prepare Zones With IDs"))
                                                {
                                                    zonesWithIds = placedZones.Where(z => z.SleeveInstanceId > 0).ToList();
                                                }
                                                using (placeTracker?.TrackSubOperation("BatchUpdateSleevePlacementData"))
                                                {
                                                    repo.BatchUpdateSleevePlacementData(zonesWithIds);
                                                }
                                                SafeFileLogger.SafeAppendText("placement_debug.log",
                                                    $"[{DateTime.Now:HH:mm:ss.fff}] [BULK-PERSIST] BatchUpdateSleevePlacementData completed for {zonesWithIds.Count} zones\n");
                                                
                                                // 3. Save Snapshots
                                                using (placeTracker?.TrackSubOperation("SaveSleeveSnapshotsForPlacedSleeves"))
                                                {
                                                    repo.SaveSleeveSnapshotsForPlacedSleeves(filterId > 0 ? filterId : -1, placedZones);
                                                }
                                                SafeFileLogger.SafeAppendText("placement_debug.log",
                                                    $"[{DateTime.Now:HH:mm:ss.fff}] [BULK-PERSIST-OK] Post-placement DB persistence completed\n");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                                $"[{DateTime.Now:HH:mm:ss.fff}] [BULK-PERSIST] ❌ Failed: {ex.Message}\n{ex.StackTrace}\n");
                                            if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Error($"[BULK-PERSIST] ❌ Database persistence failed: {ex.Message}");
                                        }

                                        // 2f. Set orchestration variables for downstream Corner Extraction (Step 5)
                                        placedZonesForExtraction = bulkResult.PlacedItems
                                            .Select(item => (ZoneGuid: item.Zone.Id, ElementId: item.ElementId.IntegerValue))
                                            .ToList();

                                        placedCount = bulkResult.PlacedCount;
                                        errorCount = bulkResult.FailedCount;
                                        parentTracker?.SetItemCount(bulkResult.PlacedCount);
                                    }
                                    else if (bulkResult != null && !bulkResult.OverallSuccess)
                                    {
                                        SafeFileLogger.SafeAppendText("placement_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss.fff}] [BULK-PERSIST] Skipped (OverallSuccess=false). PlacedCount={bulkResult.PlacedCount}\n");
                                        errorCount = zonesToPlace.Count;
                                        placedCount = bulkResult.PlacedCount;
                                        parentTracker?.SetItemCount(bulkResult.PlacedCount);
                                    }
                                    else if (bulkResult == null || bulkResult.PlacedCount == 0)
                                    {
                                        SafeFileLogger.SafeAppendText("placement_debug.log",
                                            $"[{DateTime.Now:HH:mm:ss.fff}] [BULK-PERSIST] Skipped (bulkResult null or PlacedCount=0). bulkResult!=null={bulkResult != null}, PlacedCount={bulkResult?.PlacedCount ?? -1}\n");
                                    }
                                    // ✅ FIX: Always sync placed count and tracker when we have bulk result (so report shows correct count even when PlacedCount=0)
                                    if (bulkResult != null)
                                    {
                                        placedCount = bulkResult.PlacedCount;
                                        parentTracker?.SetItemCount(bulkResult.PlacedCount);
                                    }
                                }
                            }
                            else
                            {
                                // ✅ STANDARD PLACEMENT PATH: Process each category separately (sequential placement)
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🚀 [STANDARD-PLACEMENT] Routing to Consolidated Standard Path\n");
                                
                                // Group zones by category for standard placement (each category needs its own conditions/strategy)
                                var zonesByCategoryStandard = zonesToPlace
                                    .GroupBy(z => z.MepElementCategory ?? "Ducts")
                                    .ToList();
                                
                                foreach (var categoryGroup in zonesByCategoryStandard)
                                {
                                    var cat = categoryGroup.Key;
                                    var categoryZonesForStandard = categoryGroup.ToList();
                                    
                                    // Get category info for this category group
                                    string catStr = cat switch
                                    {
                                        "Ducts" => "Ducts",
                                        "Duct Accessories" => "Duct Accessories",
                                        "Pipes" => "Pipes",
                                        "Cable Trays" or "CableTrays" => "Cable Trays",
                                        _ => "Ducts"
                                    };
                                    
                                    string catNm = cat switch
                                    {
                                        "Ducts" => "ducts",
                                        "Duct Accessories" => "duct_accessories",
                                        "Pipes" => "pipes",
                                        "Cable Trays" or "CableTrays" => "cable_trays",
                                        _ => "ducts"
                                    };
                                    string filterNameForCat = $"{filter.Name}_{catNm}.xml";
                                    
                                    // Load conditions and strategy for this category
                                    var projectFiltersDir = ProjectPathService.GetFiltersDirectory(_document);
                                    var conditionsService = new ConditionsService(_document, projectFiltersDir, msg => { if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg); });
                                    var conditionsKey = $"{filterNameForCat}_{catStr}";
                                    var categoryConditions = conditionsService.LoadConditions(conditionsKey) ?? new OpeningConditions { FilterName = filterNameForCat, Category = catStr };

                                    ISleevePlacementStrategy categoryStrategy = catStr.ToLower() switch {
                                        "pipes" => new PipePlacementStrategy(),
                                        "cable trays" => new CableTrayPlacementStrategy(),
                                        "duct accessories" or "ductaccessories" => new DamperPlacementStrategy(_document),
                                        "ducts" => new DuctPlacementStrategy(),
                                        _ => new DuctPlacementStrategy()
                                    };
                                
                                    using (var context = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(_document))
                                    {
                                        var repository = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClashZoneRepository(context);
                                        
                                        var placerService = new NewSleevePlacerService(
                                            _document,
                                            categoryConditions,
                                            categoryStrategy,
                                            _uiClearances ?? new Dictionary<string, double>(),
                                            repository, // ISleeveRepository (Use SQLite repo, not XML repo)
                                            null, // IZoneFilterService (not needed here)
                                            new Services.FamilyManager(_document), // IFamilyManager
                                            flagManager, // ✅ Use real flagManager instead of null
                                            isReplayPath: false,
                                            filter.Name);
                                        
                                        var placementOutcome = placerService.PlaceAllSleevesInTransaction(categoryZonesForStandard); // Use category-specific zones
                                        
                                        placedCount += placementOutcome.placed;
                                        skippedCount += placementOutcome.skipped;
                                        errorCount += placementOutcome.errors;
                                        placedZonesForExtraction.AddRange(placementOutcome.placedItems);
                                    }
                                }

                                // ✅ FIX #1: ADD SNAPSHOT SAVING TO STANDARD PLACEMENT PATH (Issue #1 Fix)
                                // Without this, snapshots are only saved in bulk mode, not in standard mode
                                if (placedCount > 0)
                                {
                                    try
                                    {
                                        // Extract zones that were actually placed (have SleeveInstanceId)
                                        var placedZonesStandard = zonesToPlace.Where(z => z.SleeveInstanceId > 0).ToList();
                                        
                                        if (placedZonesStandard.Any())
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode) 
                                                DebugLogger.Info($"[STANDARD-PERSIST] 📸 Saving snapshots for {placedZonesStandard.Count} standard-placed sleeves...");
                                            
                                            using (var dbContext = new SleeveDbContext(_document))
                                            {
                                                var repo = new ClashZoneRepository(dbContext);
                                                
                                                // Get FilterId for standard path
                                                string categoryForSnapshot = filter.Category switch
                                                {
                                                    Models.MepCategory.Ducts => "Ducts",
                                                    Models.MepCategory.DuctAccessories => "Duct Accessories",
                                                    Models.MepCategory.Pipes => "Pipes",
                                                    Models.MepCategory.CableTrays => "Cable Trays",
                                                    _ => filter.Category.ToString()
                                                };
                                                
                                                var filterLookupService = new JSE_RevitAddin_MEP_OPENINGS.Services.Filters.FilterLookupService(dbContext);
                                                int filterIdStandard = filterLookupService.GetFilterId(filter.Name, categoryForSnapshot);
                                                
                                                repo.SaveSleeveSnapshotsForPlacedSleeves(filterIdStandard > 0 ? filterIdStandard : -1, placedZonesStandard);
                                            }
                                            
                                            if (!DeploymentConfiguration.DeploymentMode) 
                                                DebugLogger.Info($"[STANDARD-PERSIST] ✅ Snapshots saved for standard placement path");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode) 
                                            DebugLogger.Warning($"[STANDARD-PERSIST] ⚠️ Failed to save snapshots: {ex.Message}");
                                    }
                                }
                            }
                            
                            // ✅ STEP 6: SAVE PLACED DATA TO DB (Unified for both paths)
                            if (placedZonesForExtraction.Any())
                            {
                                // ✅ PERFORMANCE: Track as sub-operation of individual placement
                                using (var saveTracker = parentTracker?.TrackSubOperation("Step 6: SAVE PLACED DATA TO DB"))
                                {
                                    using (var dbContext = new SleeveDbContext(_document))
                                    {
                                        var repo = new ClashZoneRepository(dbContext);
                                        // ✅ FIXED: Use BatchUpdateSleevePlacementData to update FULL placement geometry + ID
                                        // This ensures calculated dimensions and coordinates are persisted
                                        var successfullyPlacedZones = zonesToPlace.Where(z => z.SleeveInstanceId > 0).ToList();
                                        repo.BatchUpdateSleevePlacementData(successfullyPlacedZones);
                                    }
                                }
                            }
                        }
                        
                        // 🔥 CRITICAL DEBUG: Log which XML file we're passing clash zones from
                    
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss}] Placement executed: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}\n");
                    
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] PLACEMENT_COMPLETED: Placed={placedCount}, Skipped={skippedCount}, Errors={errorCount}\n");
                        }
                    } catch { }
                    
                    // ✅ Return success after processing
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
            
            // ✅ DIAGNOSTIC: Log that we're continuing past the return statement
            SafeFileLogger.SafeAppendText("batch_v2.log",
                $"[{DateTime.Now:HH:mm:ss}] 🔍 DIAGNOSTIC: Continuing after ExecuteUniversalSleevePlacement return, about to enter coordinate update block\n");
            
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
                
                // ✅ RE-ENABLED: Explicit Regenerate() + short sleep so that
                // all sleeve bounding boxes and centers reflect final Revit geometry
                // (constraints, joins, host updates) before we read them for DB persistence.
                try
                {
                    _document.Regenerate();
                    System.Threading.Thread.Sleep(200);

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Document regenerated before AFTER_REGEN + coordinate update.\n");
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[{DateTime.Now:HH:mm:ss}] Regenerate failed (continuing with existing geometry): {ex.Message}\n");
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
            

            // ✅ PHASE 2 & 3: HYBRID BATCH CLUSTERING (Extraction -> Calculation -> Placement/Swap)
            // ==========================================================================================
            // Executing Post-Placement Workflow
            // Step 5: RETRIEVED PLACED DATA FROM MODEL & EXTRACT CORNERS
            // ✅ MOVED OUTSIDE EnableClusteringWorkflow block (User Request: "verify db state")
            // ✅ PERFORMANCE: Track as sub-operation of individual placement
            using (var cornerTracker = parentTracker?.TrackSubOperation("Step 5: RETRIEVED PLACED DATA FROM MODEL"))
            {
                try
                {
                    if (placedZonesForExtraction.Any())
                    {
                       if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Phase 2: Extracting Corners for {placedZonesForExtraction.Count} zones...\n");

                       using (var dbContext = new SleeveDbContext(_document))
                       {
                            var repo = new ClashZoneRepository(dbContext);
                            var extractor = new BatchSleeveCornerExtractor(repo);
                            
                            // ✅ FIX: Fetch full ClashZone objects to get Orientation (User Request)
                            var guids = placedZonesForExtraction.Select(x => x.ZoneGuid).ToList();
                            var fullZones = repo.GetClashZonesByGuids(guids);
                            
                            // Update SleeveInstanceId in objects (since they might be fresh from DB or stale)
                            foreach (var fz in fullZones)
                            {
                                var match = placedZonesForExtraction.FirstOrDefault(p => p.ZoneGuid == fz.Id);
                                if (match.ElementId > 0) fz.SleeveInstanceId = match.ElementId;
                            }

                            // ✅ FIX: Use new signature that accepts List<ClashZone>
                            int extractedCount = extractor.ExtractAndSaveCornersForZones(_document, fullZones);
                            cornerTracker?.SetItemCount(extractedCount);
                       }
                    }

                    // ✅ Batch corner extraction for ALL placed sleeves in DB (always run): saves corners for every placed rectangular sleeve (and bbox-derived for round)
                    using (var dbContext2 = new SleeveDbContext(_document))
                    {
                        var repo2 = new ClashZoneRepository(dbContext2);
                        var extractorAll = new BatchSleeveCornerExtractor(repo2);
                        extractorAll.ExtractAndSaveCorners(_document);
                        // ✅ Cluster sleeves: extract corners from Revit and save to ClusterSleeves + ClusterSleeves_v2
                        extractorAll.ExtractAndSaveCornersForClusters(_document);
                    }
                }
                catch (Exception ex)
                {
                     if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[HybridBatch] Error in Corner Extraction: {ex.Message}");
                }
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

        /// <summary>
        /// ✅ GLOBAL CLUSTERING & CLEANUP: Centralized phase after all individual sleeves are placed.
        /// Iterates all categories, calculates clusters, and performs a single consolidated placement/cleanup.
        /// This ensures cross-category overlaps are correctly identified and cleaned up.
        /// </summary>
        /// <summary>
        /// ✅ GLOBAL CLUSTERING & CLEANUP: Centralized phase after all individual sleeves are placed.
        /// Iterates all categories, calculates clusters, and performs a single consolidated placement/cleanup.
        /// This ensures cross-category overlaps are correctly identified and cleaned up.
        /// Returns total clusters placed.
        /// </summary>
        private int ExecuteGlobalClusteringAndCleanup(List<OpeningFilter> filters, bool showProgress)
        {
            if (!OptimizationFlags.EnableClusteringWorkflow)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⏭️ GLOBAL CLUSTERING SKIPPED: EnableClusteringWorkflow=false\n");
                return 0;
            }

            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"\n[{DateTime.Now:HH:mm:ss}] 🚀 GLOBAL CLUSTERING PHASE: Starting consolidated process for all categories\n");

            try
            {
                var clusterableCategories = new[] { "Ducts", "Duct Accessories", "Pipes", "Cable Trays" };
                var clusterService = ClusterServiceFactory.CreateRefactored(_document, performanceMonitor: _performanceMonitor);
                
                using (var ctx = new SleeveDbContext(_document))
                {
                    var repo = new ClashZoneRepository(ctx);
                    
                    // 1. RECOVERY PHASE: Ensure all unresolved zones have their SleeveInstanceId from snapshots
                    // This is critical for "skip placement" runs and cross-session persistence
                    var allUnresolvedZones = repo.GetAllClashZones().Where(z => !z.IsClusterResolved).ToList();
                    int recoveredCount = 0;
                    foreach (var z in allUnresolvedZones.Where(z => z.SleeveInstanceId <= 0))
                    {
                        int recoveredId = repo.TryGetSleeveInstanceIdFromSnapshot(z.Id);
                        if (recoveredId > 0)
                        {
                            z.SleeveInstanceId = recoveredId;
                            repo.UpdateSleeveInstanceId(z.Id, recoveredId);
                            recoveredCount++;
                        }
                    }

                    if (recoveredCount > 0)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🔄 RECOVERY: Recovered {recoveredCount} sleeve IDs from snapshots before clustering\n");
                        
                        // ✅ CRITICAL: Extract and save corners for recovered sleeves 
                        // Clustering requires saved corners in the DB for proximity calculation
                        var extractor = new BatchSleeveCornerExtractor(repo);
                        extractor.ExtractAndSaveCorners(_document);
                    }

                    // 2. CALCULATION PHASE (Per Category)
                    // We load unresolved zones from DB (which now have their Revit coordinates saved)
                    // ✅ PERFORMANCE: Track calculation phase
                    using (var calcTracker = _performanceMonitor?.TrackOperation("Global Cluster Calculation"))
                    {
                        // ✅ FORCE ENABLE PARALLEL PROCESSING: Ensure optimization is active regardless of config
                        if (!OptimizationFlags.UseParallelProcessing)
                        {
                            OptimizationFlags.UseParallelProcessing = true;
                            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ FORCE-ENABLING Parallel Processing for Global Cluster Calculation.\n");
                        }

                        if (OptimizationFlags.UseParallelProcessing)
                        {
                            // ✅ PARALLEL MODE: Calculate categories concurrently (Memory Only)
                            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🚀 GLOBAL CLUSTERING: Parallel Processing ENABLED for {clusterableCategories.Length} types (Calculation Only)\n");
                            
                            var globalResults = new System.Collections.Concurrent.ConcurrentBag<JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BatchClusterCalculationResult>();

                            System.Threading.Tasks.Parallel.ForEach(clusterableCategories, cat =>
                            {
                                try 
                                {
                                    var catZones = allUnresolvedZones.Where(z => string.Equals(z.MepElementCategory, cat, StringComparison.OrdinalIgnoreCase)).ToList();
                                    if (catZones.Count > 0)
                                    {
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] 🧮 GLOBAL CLUSTERING (Parallel): Calculating for {cat} ({catZones.Count} zones) on Thread {System.Threading.Thread.CurrentThread.ManagedThreadId}\n");
                                        
                                        // ✅ ISOLATION: Create new service instance per thread
                                        var parallelClusterService = ClusterServiceFactory.CreateRefactored(_document, performanceMonitor: null);
                                        
                                        // ✅ PHASE 1: Calculate Only (No DB Writes)
                                        var results = parallelClusterService.ClusterSleevesCalculateOnly(_document, catZones, cat, 0, 0);
                                        
                                        foreach (var r in results)
                                        {
                                            globalResults.Add(r);
                                        }
                                        
                                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] 🧮 GLOBAL CLUSTERING (Parallel): {cat} complete. Yielded {results.Count} clusters.\n");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ❌ GLOBAL CLUSTERING PARALLEL ERROR ({cat}): {ex.Message}\n");
                                }
                            });
                            
                            // ✅ PHASE 2: Single Batch Save (Sequential DB Access)
                            if (globalResults.Count > 0)
                            {
                                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 💾 BATCH SAVE: Saving {globalResults.Count} total clusters to database...\n");
                                
                                // Create a service instance to handle saving
                                var saveService = ClusterServiceFactory.CreateRefactored(_document, performanceMonitor: null);
                                saveService.ClusterSleevesBatchSave(globalResults);
                                saveService.ClusterSleevesBatchUpdateFlags(globalResults);
                                
                                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH SAVE: Complete.\n");
                            }
                        }
                        else
                        {
                            // SEQUENTIAL MODE (Legacy)
                             SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🐢 GLOBAL CLUSTERING: Sequential Processing (Legacy)\n");
                            
                            // Re-use outer instance for sequential
                            var sequentialClusterService = ClusterServiceFactory.CreateRefactored(_document, performanceMonitor: _performanceMonitor);
                            
                            foreach (var cat in clusterableCategories)
                            {
                                var catZones = allUnresolvedZones.Where(z => string.Equals(z.MepElementCategory, cat, StringComparison.OrdinalIgnoreCase)).ToList();
                                if (catZones.Count == 0) continue;

                                SafeFileLogger.SafeAppendText("batch_v2.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🧮 GLOBAL CLUSTERING: Calculating for {cat} ({catZones.Count} zones)\n");

                                // Track sub-operation per category
                                using (var catTracker = calcTracker?.TrackSubOperation($"Calculate {cat}"))
                                {
                                    try
                                    {
                                        // skipPlacement: true ensures we only save the cluster definition to ClusterSleeves_v2
                                        sequentialClusterService.ClusterSleevesV2(_document, catZones, cat, 0, 0, skipPlacement: true);
                                        catTracker?.SetItemCount(catZones.Count);
                                    }
                                    catch (Exception ex)
                                    {
                                        SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ❌ GLOBAL CLUSTERING CALC ERROR ({cat}): {ex.Message}\n");
                                    }
                                }
                            }
                        }
                    }

                    // 2. CONSOLIDATED PLACEMENT & CLEANUP PHASE (Single API Transaction)
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🏗️ GLOBAL CLUSTERING: Starting consolidated placement and cleanup\n");
                    
                    string dbPath = ctx.DatabasePath;
                    var cleanupService = new ClusterCleanupService();
                    var batchPlacementService = new BatchClusterPlacementService(
                        dbPath,
                        repo,
                        new SleeveParameterService(_document),
                        cleanupService,
                        _performanceMonitor);

                    int placed = 0, failed = 0, cleanedUp = 0;
                    using (var bulkTracker = _performanceMonitor?.TrackOperation("Global Bulk Cluster Placement"))
                    {
                        (placed, failed, cleanedUp) = batchPlacementService.PlaceAllCategoriesAndCleanup(
                            _document, 
                            cleanupService, 
                            useSingleTransaction: OptimizationFlags.UseTwoTransactionBulkPlacement);
                        
                        bulkTracker?.SetItemCount(placed);
                    }

                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ✅ GLOBAL CLUSTERING COMPLETE: {placed} clusters placed, {cleanedUp} individual sleeves cleaned up\n");

                    // Report is generated once after all phases (individual + cluster) in the main flow — do not call here to avoid duplicate report
                    return placed;
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ❌ GLOBAL CLUSTERING FAILED: {ex.Message}\n{ex.StackTrace}\n");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[GlobalClustering] Error: {ex.Message}");
                return 0;
            }
            return 0;
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
                         (description.IndexOf("duplicate", StringComparison.OrdinalIgnoreCase) >= 0 || 
                          description.IndexOf("already exists", StringComparison.OrdinalIgnoreCase) >= 0))
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
