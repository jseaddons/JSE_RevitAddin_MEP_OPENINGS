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

        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument, Dictionary<string, double> uiClearances = null, MarkPrefixSettings markPrefixes = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _uiClearances = uiClearances ?? new Dictionary<string, double>();
            _markPrefixes = markPrefixes ?? new MarkPrefixSettings();
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
                // DebugLogger.Info("[OpeningCommandOrchestrator] 🔥 STARTING MARKING PHASE 🔥");
                
                // // ✅ PROPER ARCHITECTURE: Call MarkParameterCommand with UI values
                // // MarkParameterCommand handles "ALL" by processing each category with correct UI discipline prefixes
                
                // // ✅ FIX: Pass MarkPrefixSettings to MarkParameterCommand so it can use UI discipline prefixes
                // var markingCommand = new MarkParameterCommand("ALL", _markPrefixes.ProjectPrefix, "ALL", _markPrefixes.RemarkAll, _markPrefixes);
                // markingCommand.Execute(new UIApplication(_document.Application));
                
                // DebugLogger.Info("[OpeningCommandOrchestrator] ✅ MarkParameterCommand completed for ALL categories");
                // DebugLogger.Info("[OpeningCommandOrchestrator] 🔥 MARKING PHASE COMPLETED 🔥");

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
        }

        /// <summary>
        /// Execute clustering for a specific category after sleeve placement
        /// </summary>
        private void ExecuteClusteringForCategory(OpeningFilter filter, bool showProgress)
        {
            try
            {
                // 🔥 CRITICAL DEBUG: Log clustering attempt
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔥 ExecuteClusteringForCategory CALLED for category: {filter.Category} 🔥\n");
                }
                
                // Convert MepCategory enum to string for cluster command
                string categoryString = filter.Category switch
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
                string xmlFilePath = GetXmlFilePathForFilter(filter);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    string orchestratorDebugLogPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                    System.IO.File.AppendAllText(orchestratorDebugLogPath, $"[{DateTime.Now:HH:mm:ss}] xmlFilePath = {xmlFilePath ?? "NULL"}\n");
                }
                
                // Use UniversalClusterService directly (service-based architecture)
                List<FamilyInstance> placedClusterSleeves = new List<FamilyInstance>();
                
                using (var tx = new Transaction(_document, $"Cluster {categoryString} Openings"))
                {
                    tx.Start();
                    
                    var clusterService = new UniversalClusterService();
                    // ✅ FIX: Pass filter name to clustering service so it can set it on cluster sleeves
                    var (placedCount, deletedCount) = clusterService.ClusterSleeves(_document, categoryString, _uiDocument, xmlFilePath, filter.Name, placedClusterSleeves);
                    
                    tx.Commit();
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Clustering complete for {categoryString}: {placedCount} clusters placed, {deletedCount} individual sleeves deleted");
                    }
                }
                
                // ✅ PERFORMANCE FIX: After placing cluster sleeves, regenerate document and save their bounding boxes to XML
                // This uses SleeveCoordinateService to update coordinates (same as individual sleeves)
                if (placedClusterSleeves.Count > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
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
                            DebugLogger.Warning($"[OpeningCommandOrchestrator] Could not regenerate document: {regenEx.Message}");
                        }
                    }
                    
                    // Step 3: Save cluster sleeve bounding boxes to XML
                    try
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[OpeningCommandOrchestrator] About to call UpdateSleeveCoordinatesInXml with xmlFilePath: {xmlFilePath ?? "NULL"}");
                        }
                        var coordinateService = new SleeveCoordinateService(_document);
                        coordinateService.UpdateSleeveCoordinatesInXml(xmlFilePath);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Updated sleeve coordinates for cluster sleeves");
                        }
                    }
                    catch (Exception coordEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
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
                            DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Reloaded cache with cluster sleeve coordinates for {categoryString}");
                        }
                        
                        // Step 5: NOW run cleanup with updated cache (uses XML, not expensive Revit API)
                        using (var cleanupTx = new Transaction(_document, $"Cleanup sleeves within clusters"))
                        {
                            cleanupTx.Start();
                            var additionalDeleted = clusterServiceReload.CleanupSleevesWithinClustersAfterXmlSave(_document, placedClusterSleeves, xmlFilePath);
                            cleanupTx.Commit();
                            
                            if (additionalDeleted > 0 && !DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Cleaned up {additionalDeleted} additional sleeves within cluster bounding boxes");
                            }
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[OpeningCommandOrchestrator] Error in cleanup: {cleanupEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ ERROR in ExecuteClusteringForCategory: {ex.Message}\n");
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error clustering {filter.Category}: {ex.Message}");
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Stack trace: {ex.StackTrace}");
                }
                // Don't throw - clustering failure shouldn't stop the entire process
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

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[OpeningCommandOrchestrator] Looking for clash zones in: {xmlFilePath}");
                    // 🔥 CRITICAL DEBUG: Force direct file logging to trace orchestrator execution
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔍 INDIVIDUAL SLEEVE SERVICE READING FROM: {xmlFilePath}\n");
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

                // Load clash zones from XML
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                OpeningFilter loadedFilter;
                
                // 🔥 CRITICAL DEBUG: Log raw XML content before deserialization
                string rawXmlContent = File.ReadAllText(xmlFilePath);
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔍 RAW XML CONTENT (first 1000 chars): {rawXmlContent.Substring(0, Math.Min(1000, rawXmlContent.Length))}\n");
                }
                
                using (var reader = new StreamReader(xmlFilePath))
                {
                    loadedFilter = (OpeningFilter)serializer.Deserialize(reader);
                }

                // Extract clash zones from the loaded filter
                var clashZones = new List<ClashZone>();
                if (loadedFilter?.ClashZoneStorage?.ClashZones != null)
                {
                    clashZones = loadedFilter.ClashZoneStorage.ClashZones;
                    // ✅ CRITICAL FIX: Reconstruct SleevePlacementPoint and IntersectionPoint from XML-serializable properties
                    foreach (var cz in clashZones)
                    {
                        cz.EnsureSleevePlacementPointReconstructed();
                        // Also reconstruct IntersectionPoint if needed
                        if (cz.IntersectionPoint == null && (Math.Abs(cz.IntersectionPointX) > 1e-9 || Math.Abs(cz.IntersectionPointY) > 1e-9 || Math.Abs(cz.IntersectionPointZ) > 1e-9))
                        {
                            cz.IntersectionPoint = new XYZ(cz.IntersectionPointX, cz.IntersectionPointY, cz.IntersectionPointZ);
                        }
                    }
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Successfully loaded {clashZones.Count} clash zones from {xmlFilePath}\n");
                        
                        // 🔥 CRITICAL DEBUG: Check flag values immediately after deserialization
                        int clusterResolvedCount = clashZones.Count(cz => cz.IsClusterResolved);
                        int individualResolvedCount = clashZones.Count(cz => cz.IsResolved);
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 📊 FLAGS AFTER XML DESERIALIZATION: IsClusterResolved=True: {clusterResolvedCount}, IsResolved=True: {individualResolvedCount}\n");
                        
                        // 🔥 CRITICAL DEBUG: Log individual clash zone flag values to identify the issue
                        for (int i = 0; i < clashZones.Count && i < 5; i++) // Log first 5 clash zones
                        {
                            var cz = clashZones[i];
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔍 ClashZone {i}: IsResolved={cz.IsResolved}, IsClusterResolved={cz.IsClusterResolved}, ClusterSleeveInstanceId={cz.ClusterSleeveInstanceId}\n");
                        }
                        
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Successfully loaded {clashZones.Count} clash zones from {xmlFilePath}");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ No clash zones found in XML file: {xmlFilePath}\n");
                        DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found in XML file: {xmlFilePath}");
                    }
                }

                return clashZones;
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

        /// <summary>
        /// Execute UniversalSleevePlacementCommand
        /// </summary>
        private void ExecuteUniversalSleevePlacement(OpeningFilter filter, bool showProgress)
        {
            try
            {
                // 🔥 CRITICAL DEBUG: Direct file logging to trace orchestrator execution
                var tracePath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                var placementDebugPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                // ✅ OVERWRITE placement_debug.log at start of each run
                try { System.IO.File.WriteAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] === NEW PLACEMENT RUN STARTED ===\n"); } catch { }
                try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] CLICK_OK: Begin placement for category={filter.Category}, filter={filter.Name}\n"); } catch { }
                
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
                try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] LOAD_XML: zones={clashZones.Count}\n"); } catch { }

                if (clashZones.Count > 0)
                {
                    // 🔥 CRITICAL DEBUG: Log which XML file we're passing clash zones from
                    string xmlFilePath = GetXmlFilePathForFilter(filter);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔍 PASSING CLASH ZONES FROM XML FILE: {xmlFilePath}\n");
                    }
                    
                    // ✅ CRITICAL FIX: Convert enum to proper string format for strategy creation
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
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] About to create UniversalSleevePlacementCommand for category: {categoryString}, filter: {combinedFilterName}\n");
                    }
                    
                    var universalCommand = new UniversalSleevePlacementCommand(_document, clashZones, categoryString, combinedFilterName, _uiClearances);
                    try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] COMMAND_CREATED: category={categoryString}, xml={xmlFilePath}\n"); } catch { }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand created successfully, about to execute\n");
                    }
                    
                    universalCommand.Execute(_uiDocument.Application);
                    try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] COMMAND_EXECUTED\n"); } catch { }
                    
                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand executed successfully\n");
                    }
                    try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] AFTER_EXECUTE_LOG\n"); } catch { }
                    
                    // ✅ DEPLOYMENT: Wrapped in deployment mode check
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[OpeningCommandOrchestrator] UniversalSleevePlacementCommand completed successfully");
                    }
                    try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] BEFORE_UPDATE_COORDINATES\n"); } catch { }
                    
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
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ ENTERING UpdateSleeveCoordinatesInXml block...\n");
                        }
                        try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] ENTERING_UPDATE_COORD_BLOCK\n"); } catch { }
                        
                        // ✅ DEPLOYMENT: Wrapped in deployment mode check
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
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
                                
                                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_REGEN] Found {sleevesBeforeRegen.Count} sleeves before regeneration\n"); } catch { }
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
                                            try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_BEFORE_REGEN] Sleeve {sleeve.Id.IntegerValue}: Bounding box is NULL\n"); } catch { }
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
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Document regenerated, now getting bounding boxes...\n");
                            }
                        }
                        catch (Exception regenEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
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
                                
                                try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_AFTER_REGEN] Found {sleevesAfterRegen.Count} sleeves after regeneration\n"); } catch { }
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
                                            try { System.IO.File.AppendAllText(placementDebugPath, $"[{DateTime.Now:HH:mm:ss}] [BOUNDING_BOX_AFTER_REGEN] Sleeve {sleeve.Id.IntegerValue}: Bounding box is NULL\n"); } catch { }
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
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Getting individual sleeve bounding boxes from Revit...\n");
                        }
                        try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] BEFORE_UpdateSleeveCoordinatesInXml\n"); } catch { }
                        
                        var coordinateService = new SleeveCoordinateService(_document);
                        try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] SleeveCoordinateService_CREATED\n"); } catch { }
                        
                        // ✅ DEPLOYMENT: Wrapped in deployment mode check
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ About to call UpdateSleeveCoordinatesInXml with xmlFilePath: {xmlFilePath ?? "NULL"}\n");
                        }
                        try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] CALLING_UpdateSleeveCoordinatesInXml: {xmlFilePath ?? "NULL"}\n"); } catch { }
                        
                        coordinateService.UpdateSleeveCoordinatesInXml(xmlFilePath);
                        
                        try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] AFTER_UpdateSleeveCoordinatesInXml\n"); } catch { }
                        
                        // ✅ DEPLOYMENT: Wrapped in deployment mode check
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Individual sleeve coordinates saved - clustering can now calculate proximity\n");
                        }
                        try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] UPDATE_COORD_SUCCESS\n"); } catch { }
                    }
                    catch (Exception coordEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ⚠️ Error saving individual coordinates: {coordEx.Message}\n");
                            DebugLogger.Error($"[OpeningCommandOrchestrator] UpdateSleeveCoordinatesInXml Exception: {coordEx.Message}\n");
                            DebugLogger.Error($"[OpeningCommandOrchestrator] Stack trace: {coordEx.StackTrace}\n");
                        }
                        try { System.IO.File.AppendAllText(tracePath, $"[{DateTime.Now:HH:mm:ss}] UPDATE_COORD_EXCEPTION: {coordEx.Message}\n"); } catch { }
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found for {filter.Category}, skipping placement");
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ❌ EXCEPTION in ExecuteUniversalSleevePlacement: {ex.Message}\n");
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Stack trace: {ex.StackTrace}\n");
                    DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing UniversalSleevePlacementCommand: {ex.Message}");
                }
                throw;
            }
        }

        /// <summary>
        /// Execute individual command with resource management using ExecuteImpl pattern
        /// </summary>
        private void ExecuteCommandWithResourceManagement(IExternalCommand command, OpeningFilter filter, bool showProgress)
        {
            try
            {
                DebugLogger.Info($"[OpeningCommandOrchestrator] Executing {command.GetType().Name} for {filter.Category}");

                // ✅ FIX: No longer using RectangularSleeveClusterCommandV2 - clustering handled by UniversalClusterService
                DebugLogger.Warning($"[OpeningCommandOrchestrator] Command {command.GetType().Name} not supported for direct execution - skipping");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing {command.GetType().Name}: {ex.Message}");
                throw;
    }
}

    }
}
