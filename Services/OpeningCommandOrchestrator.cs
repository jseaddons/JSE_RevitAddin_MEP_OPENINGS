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
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

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
        
        // ⚠️ CRITICAL: Crash-safe executor for timeout protection (5-minute limit per category)
        private readonly CrashSafeExecutor _crashSafeExecutor;

        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument, Dictionary<string, double> uiClearances = null, MarkPrefixSettings markPrefixes = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _uiClearances = uiClearances ?? new Dictionary<string, double>();
            _markPrefixes = markPrefixes ?? new MarkPrefixSettings();
            
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

            try
            {
                // Group filters by name for memory management
                var disciplineGroups = GroupFiltersByName(filters);

                foreach (var discipline in disciplineGroups)
                {
                    ExecuteDisciplineWithMemoryManagement(discipline.Key, discipline.Value, showProgress);
                }

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
            foreach (var command in commands)
            {
                ExecuteCommandWithResourceManagement(command, filter, showProgress);
            }

            // ⚠️ CRITICAL: INDIVIDUAL SLEEVES MUST RUN FIRST - DO NOT CHANGE THIS ORDER! ⚠️
            // Clustering REQUIRES existing individual sleeves to work - it collects sleeves from Revit
            // If clustering runs first, there will be no sleeves to cluster and it will fail
            // NEVER PUT ExecuteClusteringForCategory BEFORE ExecuteUniversalSleevePlacement
            ExecuteUniversalSleevePlacement(filter, showProgress);
            
            // ✅ CRITICAL: Execute clustering immediately after sleeve placement for each category
            // This follows the architecture: Place sleeves → Cluster sleeves → Mark sleeves
            // Use UniversalClusterService directly (faster than old RectangularSleeveClusterCommandV2)
            ExecuteClusteringForCategory(filter, showProgress);
            
            // ✅ COMBO FLAG: Reset IsFilterComboNew to false after full sequence completes (individual + cluster)
            // This marks the filter+category combo as "used" so next time it will use PATH 1 (Replay)
            ResetFilterComboFlagAfterPlacement(filter);
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
                            // Reset all file combos for this filter that have IsFilterComboNew=1
                            cmd.CommandText = @"
                                UPDATE FileCombos
                                SET IsFilterComboNew = 0
                                WHERE FilterId = @FilterId AND IsFilterComboNew = 1";
                            cmd.Parameters.AddWithValue("@FilterId", filterId);
                            
                            var affected = cmd.ExecuteNonQuery();
                            if (affected > 0)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[COMBO-FLAG] ✅ Reset IsFilterComboNew=false for {affected} file combo(s) for filter '{filter.Name}' (Category='{categoryName}') - file combos marked as used");
                                }
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[COMBO-FLAG] No file combos to reset for filter '{filter.Name}' (Category='{categoryName}') - all already marked as used");
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
        private void ExecuteClusteringForCategory(OpeningFilter filter, bool showProgress)
        {
            // Declare variables outside lambda for use after timeout execution
            List<FamilyInstance> placedClusterSleeves = new List<FamilyInstance>();
            string categoryString = null;
            string xmlFilePath = null;
            
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
                    
                    // ✅ PERFORMANCE FIX: Get XML file path for this category to avoid loading all 22 XML files
                    xmlFilePath = GetXmlFilePathForFilter(filter);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string orchestratorDebugLogPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            File.AppendAllText(orchestratorDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] xmlFilePath = {xmlFilePath ?? "NULL"}\n");
                        }
                    }
                    
                    // Use UniversalClusterService directly (service-based architecture)
                    
                    using (var tx = new Transaction(_document, $"Cluster {categoryString} Openings"))
                    {
                        tx.Start();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORCHESTRATOR] ✅ Transaction STARTED: '{tx.GetName()}', Document.IsModifiable: {_document.IsModifiable}");
                        }
                        
                        var clusterService = new UniversalClusterService();
                        // ✅ FIX: Pass filter name to clustering service so it can set it on cluster sleeves
                        var (placedCount, deletedCount) = clusterService.ClusterSleeves(_document, categoryString, _uiDocument, xmlFilePath, filter.Name, placedClusterSleeves);
                        
                        // ✅ CRITICAL LOGGING: Log cluster sleeves returned from ClusterSleeves
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
                        
                        tx.Commit();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[ORCHESTRATOR] ✅ Transaction COMMITTED successfully: '{tx.GetName()}'");
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
            
            if (result != Autodesk.Revit.UI.Result.Succeeded)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Clustering for {filter.Category} completed with status: {result}");
                }
                return; // Exit early if clustering failed
            }
            
            // ✅ PERFORMANCE FIX: After placing cluster sleeves, regenerate document and save their bounding boxes to XML
            // This uses SleeveCoordinateService to update coordinates (same as individual sleeves)
            if (placedClusterSleeves != null && placedClusterSleeves.Count > 0 && categoryString != null && xmlFilePath != null)
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
                    }
                    catch (Exception regenEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[OpeningCommandOrchestrator] Could not regenerate document: {regenEx.Message}");
                        }
                    }
                    
                    // Step 3: Save cluster sleeve bounding boxes to XML
                    try
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[OpeningCommandOrchestrator] About to call UpdateSleeveCoordinatesInXml with xmlFilePath: {xmlFilePath ?? "NULL"}");
                        }
                        var coordinateService = new SleeveCoordinateService(_document);
                        coordinateService.UpdateSleeveCoordinatesInXml(xmlFilePath);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Updated sleeve coordinates for cluster sleeves");
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
                    
                    // Step 4: Reload cache with updated cluster sleeve coordinates
                    try
                    {
                        var clusterServiceReload = new UniversalClusterService();
                        // Load cache for the specific category being processed
                        clusterServiceReload.LoadClashZoneCacheForCleanup(xmlFilePath, categoryString, _document, filter.Name);
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
                            // Renamed method: XmlSave -> DbSave (XML fallback still supported via xmlFilePath parameter)
                            var additionalDeleted = clusterServiceReload.CleanupSleevesWithinClustersAfterDbSave(_document, placedClusterSleeves, xmlFilePath);
                            cleanupTx.Commit();
                            
                            if (additionalDeleted > 0 && !DeploymentConfiguration.DeploymentMode)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Cleaned up {additionalDeleted} additional sleeves within cluster bounding boxes");
                                }
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
                // ✅ PERFORMANCE FIX: Use the helper method to get XML file path
                string xmlFilePath = GetXmlFilePathForFilter(filter);

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
                var dbClashZones = LoadClashZonesFromDatabase(filter, categoryName);
                if (DeploymentConfiguration.UseSqliteAsPrimary)
                {
                    // Phase 2: SQLite is primary
                    if (dbClashZones != null && dbClashZones.Count > 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ PHASE 2: Loaded {dbClashZones.Count} clash zones from SQLite (PRIMARY) for filter '{filter.Name}' ({categoryName})\n");
                        }
                        return dbClashZones;
                    }
                    else
                    {
                        // SQLite has no zones - fallback to XML
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[OpeningCommandOrchestrator] PHASE 2: SQLite has no zones, falling back to XML: {xmlFilePath}");
                        }
                    }
                }
                else
                {
                    // Legacy mode: XML is primary, SQLite is fallback
                    if (dbClashZones != null && dbClashZones.Count > 0)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Loaded {dbClashZones.Count} clash zones from SQLite (fallback) for filter '{filter.Name}' ({categoryName})\n");
                        }
                        return dbClashZones;
                    }
                }

                                        if (!DeploymentConfiguration.DeploymentMode)
                {
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Looking for clash zones in: {xmlFilePath}");
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔍 GUID-GUIDED LOAD: Category='{categoryName}', Filter='{filter.Name}'\n");
                }

                // ✅ STEP 1: Load unresolved entries from Global XML (deterministic GUIDs)
                // Global XML is the single source of truth for which zones need placement
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, categoryName);
                var allGlobalEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                
                // Get unresolved GUIDs (zones that need placement)
                var unresolvedGuids = allGlobalEntries
                    .Where(e => !e.IsResolved && !e.IsClusterResolved)
                    .Select(e => Guid.Parse(e.Id))
                    .ToHashSet();
                
                                        if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[GUID-GUIDED-LOAD] Found {unresolvedGuids.Count} unresolved GUIDs in Global XML for category '{categoryName}'");
                }

                if (!File.Exists(xmlFilePath))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ XML file not found: {xmlFilePath}\n");
                            DebugLogger.Warning($"[OpeningCommandOrchestrator] XML file not found: {xmlFilePath}");
                    }
                    return new List<ClashZone>();
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ XML file found: {xmlFilePath}\n");
                }

                // ✅ STEP 2: Load Filter XML and filter by unresolved GUIDs
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                OpeningFilter loadedFilter;
                
                using (var reader = new StreamReader(xmlFilePath))
                {
                    loadedFilter = (OpeningFilter)serializer.Deserialize(reader);
                }

                // Extract clash zones from the loaded filter (tree-aware)
                var allClashZones = ExtractClashZonesFromStorage(loadedFilter?.ClashZoneStorage);

                if (allClashZones.Count == 0)
                {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                    {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ No clash zones found in XML file: {xmlFilePath}\n");
                            DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found in XML file: {xmlFilePath}");
                    }
                    return new List<ClashZone>();
                }

                // ✅ CRITICAL FIX: Reconstruct SleevePlacementPoint and IntersectionPoint from XML-serializable properties
                foreach (var cz in allClashZones)
                {
                    cz.EnsureSleevePlacementPointReconstructed();
                    if (cz.IntersectionPoint == null && (Math.Abs(cz.IntersectionPointX) > 1e-9 || Math.Abs(cz.IntersectionPointY) > 1e-9 || Math.Abs(cz.IntersectionPointZ) > 1e-9))
                    {
                        cz.IntersectionPoint = new XYZ(cz.IntersectionPointX, cz.IntersectionPointY, cz.IntersectionPointZ);
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Successfully loaded {allClashZones.Count} total clash zones from {xmlFilePath}\n");
                }

                // ✅ STEP 3: Filter clash zones by unresolved GUIDs from Global XML
                // Only load clash zones that match unresolved GUIDs (deterministic GUID guides us to correct Filter XML data)
                var clashZones = new List<ClashZone>();
                
                if (unresolvedGuids.Count > 0)
                {
                    // Filter by GUID - only include zones that match unresolved GUIDs from Global XML
                    clashZones = allClashZones
                        .Where(cz => unresolvedGuids.Contains(cz.Id))
                        .ToList();
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[GUID-GUIDED-LOAD] ✅ Filtered {allClashZones.Count} total zones → {clashZones.Count} zones matching unresolved GUIDs from Global XML");
                        
                        // Log which GUIDs were found/not found
                        var foundGuids = clashZones.Select(cz => cz.Id).ToHashSet();
                        var missingGuids = unresolvedGuids.Except(foundGuids).ToList();
                        if (missingGuids.Count > 0)
                        {
                            DebugLogger.Warning($"[GUID-GUIDED-LOAD] ⚠️ {missingGuids.Count} unresolved GUIDs from Global XML not found in Filter XML: {string.Join(", ", missingGuids.Take(5))}{(missingGuids.Count > 5 ? "..." : "")}");
                        }
                    }
                }
                else
                {
                    // No unresolved GUIDs - return empty list (all zones are resolved)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[GUID-GUIDED-LOAD] ✅ No unresolved GUIDs in Global XML - all zones are resolved, returning empty list");
                    }
                    return new List<ClashZone>();
                }

                // ✅ STEP 4: Sync flags from Global XML (already loaded above)
                // Since we're GUID-guided, we can sync flags directly by GUID match using FlagManager
                if (clashZones.Count > 0)
                {
                    try
                    {
                        // Use FlagManager for efficient flag syncing by GUID
                        var flagManager = new FlagManager(_document);
                        flagManager.SyncFlagsFromGlobal(clashZones, categoryName);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            int clusterResolvedAfterSync = clashZones.Count(cz => cz.IsClusterResolved);
                            int individualResolvedAfterSync = clashZones.Count(cz => cz.IsResolved);
                            DebugLogger.Info($"[GUID-GUIDED-LOAD] ✅ Synced flags from Global XML: {clashZones.Count} zones, {clusterResolvedAfterSync} cluster-resolved, {individualResolvedAfterSync} individual-resolved");
                        }
                    }
                    catch (Exception syncEx)
                    {
                        // Log error but continue - don't fail placement if sync fails
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Warning($"[GUID-GUIDED-LOAD] Error syncing flags from Global XML: {syncEx.Message}");
                        }
                    }
                }

                return clashZones;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
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

                    // ✅ CRITICAL FIX: Load ALL zones from DB (unresolvedOnly=false), then filter by flags in memory
                    // This ensures we get zones even after flag reset (IsResolved=false, SleeveInstanceId=-1)
                    // DB query filters by SleeveState which might not match reset flags correctly
                    var zones = repository.GetClashZonesByFilter(filter.Name, categoryName, unresolvedOnly: false) ?? new List<ClashZone>();

                    foreach (var zone in zones)
                    {
                        zone?.EnsureSleevePlacementPointReconstructed();
                        zone?.EnsureSleevePlacementPointActiveDocumentReconstructed();
                    }

                    // ✅ FILTER IN MEMORY: Only return zones that need placement (not resolved, not cluster resolved)
                    // This ensures placement uses DB zones correctly, filtering by actual flag state
                    var eligibleZones = zones
                        .Where(cz => cz != null && !cz.IsResolved && !cz.IsClusterResolved)
                        .ToList();
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[OpeningCommandOrchestrator][SQLite] Loaded {zones.Count} total zones, {eligibleZones.Count} eligible for placement (filtered by flags in memory) - DATA SOURCE: DATABASE");
                        // ✅ CRITICAL: Log data source for placement debugging
                        var placementLogPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        try
                        {
                            File.AppendAllText(placementLogPath, $"[{DateTime.Now:HH:mm:ss}] [DATA-SOURCE] ✅ Using DATABASE for filter '{filter.Name}', category '{categoryName}' ({eligibleZones.Count} eligible zones from {zones.Count} total)\n");
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
        private void ExecuteUniversalSleevePlacement(OpeningFilter filter, bool showProgress)
        {
            // Declare variables outside lambda for use after timeout execution
            string xmlFilePath = GetXmlFilePathForFilter(filter);
            var tracePath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
            var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
            
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
                        // Invalidated zones are zones that have existing sleeves (SleeveInstanceId > 0)
                        // These zones need to be deleted and re-placed at new intersection points
                        // Note: This is a simplified check - in production, invalidated zones should be
                        // marked during refresh and stored in RefreshContext.InvalidatedZones
                        var invalidatedZones = clashZones.Where(cz => cz.SleeveInstanceId > 0).ToList();
                        
                        var validatedZones = clashZones.Except(invalidatedZones).ToList();
                        
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
                                
                                var invalidatedService = new Path3InvalidatedPlacementService(_document);
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
                            
                            // Continue with validated zones for normal placement
                            clashZones = validatedZones;
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
                    
                    var universalCommand = new UniversalSleevePlacementCommand(_document, clashZones, categoryString, combinedFilterName, _uiClearances);
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] COMMAND_CREATED: category={categoryString}, xml={xmlFilePath}\n");
                        }
                    } catch { }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand created successfully, about to execute\n");
                    }
                    
                    universalCommand.Execute(_uiDocument.Application);
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] COMMAND_EXECUTED\n");
                        }
                    } catch { }
                    
                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand executed successfully\n");
                    }
                    try {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] AFTER_EXECUTE_LOG\n");
                        }
                    } catch { }
                    
                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[OpeningCommandOrchestrator] UniversalSleevePlacementCommand completed successfully");
                    }
                    
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
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Sleeve placement for {filter.Category} completed with status: {result}");
                }
                return; // Exit early if placement failed
            }
            
            // ✅ CRITICAL: Following reference document - Regenerate FIRST, then read from Revit and save to XML
            // Reference: SLEEVE_PLACEMENT_SEQUENCING_REFERENCE.md lines 22-35
            // The timing fix: Regenerate ensures bounding boxes are available, then UpdateSleeveCoordinatesInXml reads from Revit
            // ✅ CRITICAL: Save individual sleeve bounding boxes BEFORE clustering
            // Clustering proximity calculation REQUIRES individual sleeve bounding boxes from XML
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
                
                // ✅ DEPLOYMENT: Wrapped in deployment mode check
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ About to call UpdateSleeveCoordinatesInXml with xmlFilePath: {xmlFilePath ?? "NULL"}\n");
                }
                try {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] CALLING_UpdateSleeveCoordinatesInXml: {xmlFilePath ?? "NULL"}\n");
                    }
                } catch { }
                
                coordinateService.UpdateSleeveCoordinatesInXml(xmlFilePath);
                
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
}
