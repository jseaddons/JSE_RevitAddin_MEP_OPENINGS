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
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔥 ExecuteMultipleFilters CALLED 🔥\n");
            
            DebugLogger.Info("[OpeningCommandOrchestrator] 🔥 ExecuteMultipleFilters CALLED 🔥");
            
            if (filters == null || filters.Count == 0)
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ No filters provided\n");
                DebugLogger.Warning("[OpeningCommandOrchestrator] No filters provided");
                return;
            }

            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] Starting execution of {filters.Count} filters\n");
            DebugLogger.Info($"[OpeningCommandOrchestrator] Starting execution of {filters.Count} filters");

            try
            {
                // Group filters by name for memory management
                var disciplineGroups = GroupFiltersByName(filters);

                foreach (var discipline in disciplineGroups)
                {
                    ExecuteDisciplineWithMemoryManagement(discipline.Key, discipline.Value, showProgress);
                }

                // Execute marking for all disciplines at the end
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 STARTING MARKING PHASE 🔥\n");
                DebugLogger.Info("[OpeningCommandOrchestrator] 🔥 STARTING MARKING PHASE 🔥");
                
                // ✅ PROPER ARCHITECTURE: Call MarkParameterCommand with UI values
                // MarkParameterCommand handles "ALL" by processing each category with correct UI discipline prefixes
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 Calling MarkParameterCommand for ALL categories with UI values 🔥\n");
                
                // ✅ FIX: Pass MarkPrefixSettings to MarkParameterCommand so it can use UI discipline prefixes
                var markingCommand = new MarkParameterCommand("ALL", _markPrefixes.ProjectPrefix, "ALL", _markPrefixes.RemarkAll, _markPrefixes);
                markingCommand.Execute(new UIApplication(_document.Application));
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ✅ MarkParameterCommand completed for ALL categories\n");
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 MARKING PHASE COMPLETED 🔥\n");
                DebugLogger.Info("[OpeningCommandOrchestrator] 🔥 MARKING PHASE COMPLETED 🔥");

                DebugLogger.Info("[OpeningCommandOrchestrator] All filters executed successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing filters: {ex.Message}");
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
            DebugLogger.Info($"[OpeningCommandOrchestrator] 🔥 ExecuteDisciplineWithMemoryManagement CALLED 🔥");
            DebugLogger.Info($"[OpeningCommandOrchestrator] Executing discipline: {discipline} with {filters.Count} filters");

            try
            {
                // ✅ PRIORITY ORDERING: Sort filters to ensure duct accessories are processed before ducts
                var orderedFilters = OrderFiltersByPriority(filters);
                DebugLogger.Info($"[OpeningCommandOrchestrator] Ordered {orderedFilters.Count} filters by priority for discipline: {discipline}");
                
                // Log the processing order
                for (int i = 0; i < orderedFilters.Count; i++)
                {
                    var filter = orderedFilters[i];
                    var categories = string.Join(", ", filter.SelectedMepCategoryNames ?? new List<string>());
                    var priority = GetFilterPriority(filter);
                    DebugLogger.Info($"[OpeningCommandOrchestrator] Processing order {i + 1}: '{filter.Name}' (Categories: {categories}, Priority: {priority})");
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

                DebugLogger.Info($"[OpeningCommandOrchestrator] Discipline {discipline} completed, memory cleaned");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing discipline {discipline}: {ex.Message}");
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
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 ExecuteClusteringForCategory CALLED for category: {filter.Category} 🔥\n");
                
                // Convert MepCategory enum to string for cluster command
                string categoryString = filter.Category switch
                {
                    Models.MepCategory.Ducts => "Ducts",
                    Models.MepCategory.DuctAccessories => "Duct Accessories",
                    Models.MepCategory.Pipes => "Pipes", 
                    Models.MepCategory.CableTrays => "Cable Trays",
                    _ => "Ducts"
                };
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 Starting clustering for category: {categoryString} 🔥\n");
                
                DebugLogger.Info($"[OpeningCommandOrchestrator] Starting clustering for category: {categoryString}");
                
                // ✅ PERFORMANCE FIX: Get XML file path for this category to avoid loading all 22 XML files
                string xmlFilePath = GetXmlFilePathForFilter(filter);
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] xmlFilePath = {xmlFilePath ?? "NULL"}\n");
                
                // Use UniversalClusterService directly (service-based architecture)
                List<FamilyInstance> placedClusterSleeves = new List<FamilyInstance>();
                
                using (var tx = new Transaction(_document, $"Cluster {categoryString} Openings"))
                {
                    tx.Start();
                    
                    var clusterService = new UniversalClusterService();
                    // ✅ FIX: Pass filter name to clustering service so it can set it on cluster sleeves
                    var (placedCount, deletedCount) = clusterService.ClusterSleeves(_document, categoryString, _uiDocument, xmlFilePath, filter.Name, placedClusterSleeves);
                    
                    tx.Commit();
                    
                    DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Clustering complete for {categoryString}: {placedCount} clusters placed, {deletedCount} individual sleeves deleted");
                }
                
                // ✅ PERFORMANCE FIX: After placing cluster sleeves, regenerate document and save their bounding boxes to XML
                // This uses SleeveCoordinateService to update coordinates (same as individual sleeves)
                if (placedClusterSleeves.Count > 0)
                {
                    DebugLogger.Info($"[OpeningCommandOrchestrator] Regenerating document and updating coordinates for {placedClusterSleeves.Count} cluster sleeves");
                    
                    try
                    {
                        // Step 1: Regenerate document to ensure bounding boxes are available
                        _document.Regenerate();
                        
                        // Step 2: Wait for regeneration to complete
                        System.Threading.Thread.Sleep(200);
                    }
                    catch (Exception regenEx)
                    {
                        DebugLogger.Warning($"[OpeningCommandOrchestrator] Could not regenerate document: {regenEx.Message}");
                    }
                    
                    // Step 3: Save cluster sleeve bounding boxes to XML
                    try
                    {
                        DebugLogger.Info($"[OpeningCommandOrchestrator] About to call UpdateSleeveCoordinatesInXml with xmlFilePath: {xmlFilePath ?? "NULL"}");
                        var coordinateService = new SleeveCoordinateService(_document);
                        coordinateService.UpdateSleeveCoordinatesInXml(xmlFilePath);
                        DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Updated sleeve coordinates for cluster sleeves");
                    }
                    catch (Exception coordEx)
                    {
                        DebugLogger.Error($"[OpeningCommandOrchestrator] Error updating coordinates: {coordEx.Message}");
                    }
                    
                    // Step 4: Reload cache with updated cluster sleeve coordinates
                    try
                    {
                        var clusterServiceReload = new UniversalClusterService();
                        clusterServiceReload.LoadClashZoneCacheForCleanup(xmlFilePath, categoryString, _document, filter.Name);
                        DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Reloaded cache with cluster sleeve coordinates");
                        
                        // Step 5: NOW run cleanup with updated cache (uses XML, not expensive Revit API)
                        using (var cleanupTx = new Transaction(_document, $"Cleanup sleeves within clusters"))
                        {
                            cleanupTx.Start();
                            var additionalDeleted = clusterServiceReload.CleanupSleevesWithinClustersAfterXmlSave(_document, placedClusterSleeves);
                            cleanupTx.Commit();
                            
                            if (additionalDeleted > 0)
                            {
                                DebugLogger.Info($"[OpeningCommandOrchestrator] ✓ Cleaned up {additionalDeleted} additional sleeves within cluster bounding boxes");
                            }
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        DebugLogger.Error($"[OpeningCommandOrchestrator] Error in cleanup: {cleanupEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ ERROR in ExecuteClusteringForCategory: {ex.Message}\n");
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error clustering {filter.Category}: {ex.Message}");
                DebugLogger.Error($"[OpeningCommandOrchestrator] Stack trace: {ex.StackTrace}");
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
            // Convert MepCategory enum back to string for file naming
            string categoryName = filter.Category switch
            {
                Models.MepCategory.Ducts => "Ducts",
                Models.MepCategory.DuctAccessories => "Duct Accessories", 
                Models.MepCategory.Pipes => "Pipes",
                Models.MepCategory.CableTrays => "Cable Trays",
                _ => "Ducts"
            };

            // Construct XML file path (same logic as LoadClashZonesForFilter)
            string filterName = filter.Name;
            string xmlFileName = $"{filterName}_{categoryName.Replace(" ", "_").ToLower()}.xml";
            string xmlFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters", xmlFileName);
            
            return xmlFilePath;
        }

        private List<ClashZone> LoadClashZonesForFilter(OpeningFilter filter)
        {
            try
            {
                // ✅ PERFORMANCE FIX: Use the helper method to get XML file path
                string xmlFilePath = GetXmlFilePathForFilter(filter);

                DebugLogger.Info($"[OpeningCommandOrchestrator] Looking for clash zones in: {xmlFilePath}");
                
                // 🔥 CRITICAL DEBUG: Force direct file logging to trace orchestrator execution
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 INDIVIDUAL SLEEVE SERVICE READING FROM: {xmlFilePath}\n");

                if (!File.Exists(xmlFilePath))
                {
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ❌ XML file not found: {xmlFilePath}\n");
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] XML file not found: {xmlFilePath}");
                    return new List<ClashZone>();
                }
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ✅ XML file found: {xmlFilePath}\n");

                // Load clash zones from XML
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                OpeningFilter loadedFilter;
                
                // 🔥 CRITICAL DEBUG: Log raw XML content before deserialization
                string rawXmlContent = File.ReadAllText(xmlFilePath);
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔍 RAW XML CONTENT (first 1000 chars): {rawXmlContent.Substring(0, Math.Min(1000, rawXmlContent.Length))}\n");
                
                using (var reader = new StreamReader(xmlFilePath))
                {
                    loadedFilter = (OpeningFilter)serializer.Deserialize(reader);
                }

                // Extract clash zones from the loaded filter
                var clashZones = new List<ClashZone>();
                if (loadedFilter?.ClashZoneStorage?.ClashZones != null)
                {
                    clashZones = loadedFilter.ClashZoneStorage.ClashZones;
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ✅ Successfully loaded {clashZones.Count} clash zones from {xmlFilePath}\n");
                    
                // 🔥 CRITICAL DEBUG: Check flag values immediately after deserialization
                int clusterResolvedCount = clashZones.Count(cz => cz.IsClusterResolved);
                int individualResolvedCount = clashZones.Count(cz => cz.IsResolved);
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📊 FLAGS AFTER XML DESERIALIZATION: IsClusterResolved=True: {clusterResolvedCount}, IsResolved=True: {individualResolvedCount}\n");
                
                // 🔥 CRITICAL DEBUG: Log individual clash zone flag values to identify the issue
                for (int i = 0; i < clashZones.Count && i < 5; i++) // Log first 5 clash zones
                {
                    var cz = clashZones[i];
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 ClashZone {i}: IsResolved={cz.IsResolved}, IsClusterResolved={cz.IsClusterResolved}, ClusterSleeveInstanceId={cz.ClusterSleeveInstanceId}\n");
                }
                    
                    DebugLogger.Info($"[OpeningCommandOrchestrator] Successfully loaded {clashZones.Count} clash zones from {xmlFilePath}");
                }
                else
                {
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ❌ No clash zones found in XML file: {xmlFilePath}\n");
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found in XML file: {xmlFilePath}");
                }

                return clashZones;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error loading clash zones for filter {filter.Name}: {ex.Message}");
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
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 ExecuteUniversalSleevePlacement CALLED 🔥\n");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] Filter Category: {filter.Category}\n");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] UI Clearances Count: {_uiClearances?.Count ?? 0}\n");
                
                DebugLogger.Info($"[OpeningCommandOrchestrator] Executing UniversalSleevePlacementCommand for {filter.Category}");

                // ✅ CRITICAL FIX: Load clash zones from XML file
                var clashZones = LoadClashZonesForFilter(filter);
                DebugLogger.Info($"[OpeningCommandOrchestrator] Loaded {clashZones.Count} clash zones for {filter.Category}");

                if (clashZones.Count > 0)
                {
                    // 🔥 CRITICAL DEBUG: Log which XML file we're passing clash zones from
                    string xmlFilePath = GetXmlFilePathForFilter(filter);
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 PASSING CLASH ZONES FROM XML FILE: {xmlFilePath}\n");
                    
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
                    
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] About to create UniversalSleevePlacementCommand for category: {categoryString}, filter: {combinedFilterName}\n");
                    
                    var universalCommand = new UniversalSleevePlacementCommand(_document, clashZones, categoryString, combinedFilterName, _uiClearances);
                    
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand created successfully, about to execute\n");
                    
                    universalCommand.Execute(_uiDocument.Application);
                    
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] UniversalSleevePlacementCommand executed successfully\n");
                    
                    DebugLogger.Info($"[OpeningCommandOrchestrator] UniversalSleevePlacementCommand completed successfully");
                    
                    // ✅ CRITICAL: Save individual sleeve bounding boxes BEFORE clustering
                    // Clustering proximity calculation REQUIRES individual sleeve bounding boxes from XML
                    try
                    {
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] Getting individual sleeve bounding boxes from Revit...\n");
                        
                        var coordinateService = new SleeveCoordinateService(_document);
                        coordinateService.UpdateSleeveCoordinatesInXml(xmlFilePath);
                        
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ Individual sleeve coordinates saved - clustering can now calculate proximity\n");
                    }
                    catch (Exception coordEx)
                    {
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error saving individual coordinates: {coordEx.Message}\n");
                    }
                }
                else
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found for {filter.Category}, skipping placement");
                }
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ EXCEPTION in ExecuteUniversalSleevePlacement: {ex.Message}\n");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] Stack trace: {ex.StackTrace}\n");
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing UniversalSleevePlacementCommand: {ex.Message}");
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
