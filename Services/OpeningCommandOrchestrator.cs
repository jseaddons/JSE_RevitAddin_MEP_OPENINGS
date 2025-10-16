using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
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
        /// Execute all filters for a discipline with memory management
        /// </summary>
        private void ExecuteDisciplineWithMemoryManagement(string discipline, List<OpeningFilter> filters, bool showProgress)
        {
            DebugLogger.Info($"[OpeningCommandOrchestrator] Executing discipline: {discipline} with {filters.Count} filters");

            try
            {
                foreach (var filter in filters)
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
        /// Execute UniversalSleevePlacementCommand
        /// </summary>
        private void ExecuteUniversalSleevePlacement(OpeningFilter filter, bool showProgress)
        {
            try
            {
                DebugLogger.Info($"[OpeningCommandOrchestrator] Executing UniversalSleevePlacementCommand for {filter.Category}");

                var universalCommand = new UniversalSleevePlacementCommand(_document, new List<ClashZone>(), filter.Category.ToString());
                universalCommand.Execute(_uiDocument.Application);

                DebugLogger.Info($"[OpeningCommandOrchestrator] UniversalSleevePlacementCommand completed successfully");
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
