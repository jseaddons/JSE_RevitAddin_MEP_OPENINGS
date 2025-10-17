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

        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            _uiClearances = new Dictionary<string, double>();
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
            if (filters == null || filters.Count == 0)
            {
                DebugLogger.Warning("[OpeningCommandOrchestrator] No filters provided");
                return;
            }

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
                ExecuteMarkingForAllDisciplines(filters, showProgress);

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

            // Add clustering if needed
            if (filter.OpeningType == OpeningType.RectangularClusters)
            {
                sequence.Add(new RectangularSleeveClusterCommandV2());
            }

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

            // Execute UniversalSleevePlacementCommand separately since it implements ICommand
            ExecuteUniversalSleevePlacement(filter, showProgress);
        }

        /// <summary>
        /// Load clash zones for a specific filter from XML file
        /// </summary>
        private List<ClashZone> LoadClashZonesForFilter(OpeningFilter filter)
        {
            try
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

                // Construct XML file path (CORRECTED: Use same path as refresh process)
                string filterName = filter.Name;
                string xmlFileName = $"{filterName}_{categoryName.Replace(" ", "_").ToLower()}.xml";
                string xmlFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters", xmlFileName);

                DebugLogger.Info($"[OpeningCommandOrchestrator] Looking for clash zones in: {xmlFilePath}");
                
                // 🔥 CRITICAL DEBUG: Force direct file logging to trace orchestrator execution
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] Looking for clash zones in: {xmlFilePath}\n");

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
                    // ✅ CRITICAL FIX: Convert enum to proper string format for strategy creation
                    string categoryString = filter.Category switch
                    {
                        Models.MepCategory.Ducts => "Ducts",
                        Models.MepCategory.DuctAccessories => "Duct Accessories", // Note: space, not "DuctAccessories"
                        Models.MepCategory.Pipes => "Pipes", 
                        Models.MepCategory.CableTrays => "Cable Trays", // Note: space, not "CableTrays"
                        _ => "Ducts"
                    };
                    
                    var universalCommand = new UniversalSleevePlacementCommand(_document, clashZones, categoryString, _uiClearances);
                    universalCommand.Execute(_uiDocument.Application);
                    DebugLogger.Info($"[OpeningCommandOrchestrator] UniversalSleevePlacementCommand completed successfully");
                }
                else
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] No clash zones found for {filter.Category}, skipping placement");
                }
            }
            catch (Exception ex)
            {
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

                // Use ExecuteImpl pattern for direct execution (no ExternalCommandData construction)
                if (command is RectangularSleeveClusterCommandV2 clusterCommand)
                {
                    // For now, skip clustering commands that don't have ExecuteImpl
                    // They will be handled through the Revit command system when needed
                    DebugLogger.Info($"[OpeningCommandOrchestrator] Skipping {command.GetType().Name} - will be handled through Revit command system");
                    return;
                }
                else
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Command {command.GetType().Name} not supported for direct execution - skipping");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing {command.GetType().Name}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Execute marking for all disciplines at the end using ExecuteImpl pattern
        /// </summary>
        private void ExecuteMarkingForAllDisciplines(List<OpeningFilter> filters, bool showProgress)
        {
            // For now, always execute marking since OpeningFilter doesn't have IncludeMarking property
            try
            {
                DebugLogger.Info("[OpeningCommandOrchestrator] Executing marking for all disciplines");

                var markingCommand = new MarkParameterAddValue();
                
                // Use ExecuteImpl pattern for direct execution (no ExternalCommandData construction)
                var result = markingCommand.ExecuteImpl(_document, _uiDocument, "[MARK_PARAMETER]");

                if (!result.Success)
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Marking completed with warnings: {result.ErrorMessage}");
                }
                else
                {
                    DebugLogger.Info("[OpeningCommandOrchestrator] Marking completed successfully");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing marking: {ex.Message}");
                throw;
            }
        }
    }
}
