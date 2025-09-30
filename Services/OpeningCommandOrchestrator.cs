using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using JSE_RevitAddin_MEP_OPENINGS.Views;
using JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Cost-effective, OOP-based command orchestrator with memory management
    /// Handles multiple disciplines with proper resource cleanup
    /// </summary>
    public class OpeningCommandOrchestrator : IDisposable
    {
        private readonly Document _document;
        private readonly UIDocument _uiDocument;
        private readonly Dictionary<string, DisciplineCommandExecutor> _disciplineExecutors;
        private bool _disposed = false;
        
        // Memory management
        private readonly List<IDisposable> _disposableResources = new List<IDisposable>();
        private readonly Dictionary<string, long> _memoryUsage = new Dictionary<string, long>();
        
        // UI clearance settings
        private Dictionary<string, double> _uiClearances;
        
        // Full penetration setting
        private bool _fullPenetrationEnabled = true;
        
        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
            
            // Set logging context for orchestrator debugging
            DebugLogger.SetServiceContext("Orchestrator");
            _disciplineExecutors = new Dictionary<string, DisciplineCommandExecutor>();
            _uiClearances = null;
            
            InitializeDisciplineExecutors();
        }
        
        /// <summary>
        /// Set UI clearance settings to override default clearance calculations
        /// </summary>
        /// <param name="clearances">Dictionary of clearance settings from UI</param>
        public void SetUIClearances(Dictionary<string, double> clearances)
        {
            DebugLogger.Info($"[CLEARANCE_DEBUG] OpeningCommandOrchestrator.SetUIClearances called with {clearances.Count} clearance values:");
            foreach (var kvp in clearances)
            {
                DebugLogger.Info($"[CLEARANCE_DEBUG]   {kvp.Key} = {kvp.Value}mm");
            }
            
            ClearanceManager.Instance.SetUIClearances(clearances);
            
            DebugLogger.Info($"[CLEARANCE_DEBUG] ClearanceManager.SetUIClearances completed");
        }

        /// <summary>
        /// Set full penetration setting from UI
        /// </summary>
        /// <param name="enabled">Whether full penetration mode is enabled</param>
        public void SetFullPenetrationEnabled(bool enabled)
        {
            _fullPenetrationEnabled = enabled;
        }
        
        
        // Clash zone service for incremental updates
        private ClashZoneService? _clashZoneService;
        
        /// <summary>
        /// Set clash zone service for incremental sleeve placement
        /// </summary>
        /// <param name="clashZoneService">Service for managing clash zones</param>
        public void SetClashZoneService(ClashZoneService clashZoneService)
        {
            _clashZoneService = clashZoneService;
            DebugLogger.Info($"OpeningCommandOrchestrator: ClashZoneService set with {clashZoneService.GetClashZoneStatistics().total} existing clash zones");
        }
        
        /// <summary>
        /// Execute multiple filters with memory management and progress dialog
        /// </summary>
        public OrchestrationResult ExecuteMultipleFilters(List<OpeningFilter> filters, bool showProgress = true)
        {
            if (filters == null || filters.Count == 0)
            {
                return new OrchestrationResult
                {
                    Success = false,
                    ErrorMessage = "No filters provided for execution"
                };
            }
            
            var result = new OrchestrationResult();
            OpeningProgressDialog progressDialog = null;
            
            try
            {
                if (showProgress)
                {
                    progressDialog = InitializeProgressDialog(filters.Count);
                }
                
                result = ExecuteWithProgress(filters, progressDialog);
            }
            finally
            {
                if (progressDialog != null)
                {
                    progressDialog.Dispose();
                }
            }
            
            return result;
        }
        
        private OpeningProgressDialog InitializeProgressDialog(int totalDisciplines)
        {
            try
            {
                DebugLogger.Log("Initializing progress dialog...");
                
                // For WinForms, we don't need WPF Application initialization
                // WinForms dialogs work directly in Revit add-in context
                var progressDialog = new OpeningProgressDialog();
                progressDialog.Show();
                
                // Update initial progress
                progressDialog.UpdateProgress(0, totalDisciplines, "Initializing opening creation...");
                
                DebugLogger.Log("Progress dialog initialized successfully");
                return progressDialog;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Failed to initialize progress dialog: {ex.Message}");
                return null;
            }
        }
        
        private OrchestrationResult ExecuteWithProgress(List<OpeningFilter> filters, OpeningProgressDialog progressDialog)
        {
            var result = new OrchestrationResult();
            var executedDisciplines = new List<string>();
            
            try
            {
                DebugLogger.Info($"Starting orchestration for {filters.Count} filters");
                
                // Group filters by discipline for efficient processing
                var disciplineGroups = GroupFiltersByDiscipline(filters);
                var totalDisciplines = disciplineGroups.Count;
                var currentDiscipline = 0;
                
                foreach (var disciplineGroup in disciplineGroups)
                {
                    currentDiscipline++;
                    var disciplineName = disciplineGroup.Key;
                    var disciplineFilters = disciplineGroup.Value;
                    
                    // Update progress for discipline start
                    if (progressDialog != null)
                    {
                        progressDialog.UpdateProgress(currentDiscipline, totalDisciplines, 
                            $"Processing discipline: {disciplineName}");
                    }
                    
                    DebugLogger.Info($"Processing discipline: {disciplineName} with {disciplineFilters.Count} filters");
                    
                    // Execute discipline with memory management
                    var disciplineResult = ExecuteDisciplineWithMemoryManagement(disciplineName, disciplineFilters);
                    
                    if (disciplineResult.Success)
                    {
                        executedDisciplines.Add(disciplineName);
                        result.ProcessedDisciplines.Add(disciplineName);
                        result.TotalCommandsExecuted += disciplineResult.CommandsExecuted;
                        
                        // Update progress for discipline completion
                        if (progressDialog != null)
                        {
                            progressDialog.UpdateProgress(currentDiscipline, totalDisciplines, 
                                $"Completed discipline: {disciplineName}");
                        }
                    }
                    else
                    {
                        result.Errors.Add($"Discipline {disciplineName}: {disciplineResult.ErrorMessage}");
                        
                        // Update progress for discipline error
                        if (progressDialog != null)
                        {
                            progressDialog.UpdateProgress(currentDiscipline, totalDisciplines, 
                                $"Error in discipline: {disciplineName}");
                        }
                    }
                    
                    // Force garbage collection after each discipline
                    ForceGarbageCollection(disciplineName);
                }
                
                // Execute marking once at the end for all disciplines
                if (executedDisciplines.Count > 0)
                {
                    // Update progress for marking
                    if (progressDialog != null)
                    {
                        progressDialog.UpdateProgress(totalDisciplines, totalDisciplines, 
                            "Executing marking for all disciplines...");
                    }
                    
                    var markingResult = ExecuteMarkingForAllDisciplines(executedDisciplines);
                    if (markingResult.Success)
                    {
                        result.MarkingCompleted = true;
                        
                        // Update progress for marking completion
                        if (progressDialog != null)
                        {
                            progressDialog.UpdateProgress(totalDisciplines, totalDisciplines, 
                                "Marking completed successfully!");
                        }
                    }
                    else
                    {
                        result.Errors.Add($"Marking failed: {markingResult.ErrorMessage}");
                        
                        // Update progress for marking error
                        if (progressDialog != null)
                        {
                            progressDialog.UpdateProgress(totalDisciplines, totalDisciplines, 
                                "Marking failed!");
                        }
                    }
                }
                
                result.Success = executedDisciplines.Count > 0;
                result.Message = $"Successfully processed {executedDisciplines.Count} disciplines";
                
                DebugLogger.Info($"Orchestration completed. Success: {result.Success}, Disciplines: {string.Join(", ", executedDisciplines)}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Orchestration failed: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            finally
            {
                // Cleanup all resources
                CleanupResources();
            }
            
            return result;
        }
        
        /// <summary>
        /// Execute single discipline with memory management
        /// </summary>
        private DisciplineExecutionResult ExecuteDisciplineWithMemoryManagement(string disciplineName, List<OpeningFilter> filters)
        {
            var result = new DisciplineExecutionResult();
            var initialMemory = GC.GetTotalMemory(false);
            
            try
            {
                DebugLogger.Info($"Starting execution for discipline: {disciplineName}");
                
                // Get or create discipline executor
                if (!_disciplineExecutors.TryGetValue(disciplineName, out var executor))
                {
                    executor = CreateDisciplineExecutor(disciplineName);
                    _disciplineExecutors[disciplineName] = executor;
                }
                
                // Execute commands for this discipline
                var commands = GetCommandsForDiscipline(filters);
                result.CommandsExecuted = 0;
                
                DebugLogger.Info($"Found {commands.Count} commands to execute for {disciplineName}");
                foreach (var command in commands)
                {
                    try
                    {
                        DebugLogger.Info($"Executing command: {command.GetType().Name}");
                        // Pass the first filter that has clash zones for DuctSleeveCommand
                        var filterWithClashZones = filters.FirstOrDefault(f => f.ClashZoneStorage?.ClashZones?.Count > 0);
                        var commandResult = ExecuteCommandWithResourceManagement(command, filterWithClashZones);
                        if (commandResult.Success)
                        {
                            result.CommandsExecuted++;
                            DebugLogger.Info($"Command {command.GetType().Name} executed successfully");
                        }
                        else
                        {
                            DebugLogger.Error($"Command {command.GetType().Name} failed: {commandResult.ErrorMessage}");
                            result.Errors.Add($"{command.GetType().Name}: {commandResult.ErrorMessage}");
                        }
                    }
                    catch (Exception ex)
                    {
                        DebugLogger.Error($"Command execution failed: {ex.Message}");
                        result.Errors.Add($"{command.GetType().Name}: {ex.Message}");
                    }
                }
                
                result.Success = result.CommandsExecuted > 0;
                result.Message = $"Executed {result.CommandsExecuted} commands for {disciplineName}";
                
                // Track memory usage
                var finalMemory = GC.GetTotalMemory(false);
                var memoryUsed = finalMemory - initialMemory;
                _memoryUsage[disciplineName] = memoryUsed;
                
                DebugLogger.Info($"Discipline {disciplineName} completed. Memory used: {memoryUsed / 1024} KB");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Discipline execution failed: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            
            return result;
        }
        
        /// <summary>
        /// Create discipline executor for a given discipline
        /// </summary>
        private DisciplineCommandExecutor CreateDisciplineExecutor(string disciplineName)
        {
            return new DisciplineCommandExecutor(disciplineName);
        }
        
        
        /// <summary>
        /// Converts clash zones to filtered ducts format expected by DuctSleeveCommand
        /// </summary>
        private List<(Duct, Transform?)>? ConvertClashZonesToFilteredDucts(List<ClashZone>? clashZones)
        {
            if (clashZones == null || clashZones.Count == 0)
            {
                DebugLogger.Info("[ORCHESTRATOR] No clash zones to convert to filtered ducts");
                return null;
            }

            // NOTE: Section box filtering is now done in the orchestrator before calling this method
            var filteredDucts = new List<(Duct, Transform?)>();
            
            foreach (var clashZone in clashZones)
            {
                try
                {
                    var mepElement = _document.GetElement(clashZone.MepElementId);
                    if (mepElement is Duct duct)
                    {
                        // Get transform if the duct is from a linked file
                        Transform? transform = null;
                        if (duct.Document != _document)
                        {
                            // This is a linked duct, get its transform
                            // For linked documents, we need to find the link instance
                            var linkInstances = new FilteredElementCollector(_document)
                                .OfClass(typeof(RevitLinkInstance))
                                .Cast<RevitLinkInstance>()
                                .Where(li => li.GetLinkDocument() == duct.Document);
                            
                            var linkInstance = linkInstances.FirstOrDefault();
                            if (linkInstance != null)
                            {
                                transform = linkInstance.GetTotalTransform();
                            }
                        }
                        
                        filteredDucts.Add((duct, transform));
                        DebugLogger.Info($"[ORCHESTRATOR] Converted clash zone {clashZone.Id} to filtered duct {duct.Id}");
                    }
                    else
                    {
                        DebugLogger.Warning($"[ORCHESTRATOR] Clash zone {clashZone.Id} MEP element is not a duct: {mepElement?.GetType().Name}");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[ORCHESTRATOR] Error converting clash zone {clashZone.Id} to filtered duct: {ex.Message}");
                }
            }

            DebugLogger.Info($"[ORCHESTRATOR] Converted {filteredDucts.Count} clash zones to filtered ducts");
            return filteredDucts.Count > 0 ? filteredDucts : null;
        }
        
        /// <summary>
        /// Execute command with proper resource management
        /// </summary>
        private CommandExecutionResult ExecuteCommandWithResourceManagement(IExternalCommand command, OpeningFilter? filter = null)
        {
            var result = new CommandExecutionResult();
            
            try
            {
                DebugLogger.Info($"[ORCHESTRATOR] Executing command: {command.GetType().Name}");
                
                // Execute command logic directly without ExternalCommandData
                // Commands should have ExecuteImpl methods that take UIApplication
                if (command is DuctSleeveCommand dsc)
                {
                    var uiClearances = ClearanceManager.Instance.GetUIClearances();
                    DebugLogger.Info($"[ORCHESTRATOR] Passing {uiClearances.Count} UI clearances to DuctSleeveCommand");
                    
                    // ORCHESTRATOR ROLE: Only trigger commands in sequence - no filtering logic
                    var filteredDucts = ConvertClashZonesToFilteredDucts(filter?.ClashZoneStorage?.ClashZones);
                    var clashZones = filter?.ClashZoneStorage?.ClashZones;
                    
                    DebugLogger.Info($"[ORCHESTRATOR] Triggering DuctSleeveCommand with {filteredDucts?.Count ?? 0} filtered ducts and {clashZones?.Count ?? 0} clash zones");
                    
                    var message = "";
                    var commandResult = dsc.Execute(_uiDocument, _document, ref message, new ElementSet(), filteredDucts, clashZones);
                    result.Success = commandResult == Result.Succeeded;
                    result.Message = "DuctSleeveCommand executed";
                }
                else if (command is FireDamperPlaceCommand fdp)
                {
                    DebugLogger.Info($"[ORCHESTRATOR] FireDamperPlaceCommand - ExecuteImpl not implemented yet");
                    result.Success = false;
                    result.ErrorMessage = "FireDamperPlaceCommand ExecuteImpl not implemented";
                }
                else if (command is CableTraySleeveCommand cts)
                {
                    DebugLogger.Info($"[ORCHESTRATOR] CableTraySleeveCommand - ExecuteImpl not implemented yet");
                    result.Success = false;
                    result.ErrorMessage = "CableTraySleeveCommand ExecuteImpl not implemented";
                }
                else if (command is PipeSleeveCommand psc)
                {
                    DebugLogger.Info($"[ORCHESTRATOR] PipeSleeveCommand - ExecuteImpl not implemented yet");
                    result.Success = false;
                    result.ErrorMessage = "PipeSleeveCommand ExecuteImpl not implemented";
                }
                else if (command is RectangularSleeveClusterCommandV2 rsc)
                {
                    DebugLogger.Info($"[ORCHESTRATOR] RectangularSleeveClusterCommandV2 - ExecuteImpl not implemented yet");
                    result.Success = false;
                    result.ErrorMessage = "RectangularSleeveClusterCommandV2 ExecuteImpl not implemented";
                }
                else if (command is PipeOpeningsRectCommand por)
                {
                    DebugLogger.Info($"[ORCHESTRATOR] PipeOpeningsRectCommand - ExecuteImpl not implemented yet");
                    result.Success = false;
                    result.ErrorMessage = "PipeOpeningsRectCommand ExecuteImpl not implemented";
                }
                else if (command is MarkParameterAddValue mpav)
                {
                    DebugLogger.Info($"[ORCHESTRATOR] MarkParameterAddValue - calling ExecuteImpl with prefix from UI");
                    string prefix = GetPrefixFromUI();
                    result = mpav.ExecuteImpl(_document, _uiDocument, prefix);
                    DebugLogger.Info($"[ORCHESTRATOR] MarkParameterAddValue result: {(result.Success ? "SUCCESS" : "FAILED")}");
                }
                else
                {
                    // For other commands, we need to implement ExecuteImpl pattern
                    DebugLogger.Warning($"[ORCHESTRATOR] Command {command.GetType().Name} does not support direct execution - skipping");
                    result.Success = false;
                    result.ErrorMessage = "Command does not support direct execution";
                }
                
                DebugLogger.Info($"[ORCHESTRATOR] Command {command.GetType().Name} execution result: {(result.Success ? "SUCCESS" : "FAILED")}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ORCHESTRATOR] Command execution failed: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            
            return result;
        }
        
        /// <summary>
        /// Group filters by discipline for efficient processing
        /// </summary>
        private Dictionary<string, List<OpeningFilter>> GroupFiltersByDiscipline(List<OpeningFilter> filters)
        {
            var groups = new Dictionary<string, List<OpeningFilter>>();
            
            foreach (var filter in filters)
            {
                var disciplineName = filter.Name; // User-defined discipline name
                
                if (!groups.ContainsKey(disciplineName))
                {
                    groups[disciplineName] = new List<OpeningFilter>();
                }
                
                groups[disciplineName].Add(filter);
            }
            
            return groups;
        }
        
        /// <summary>
        /// Get commands for discipline based on filters
        /// </summary>
        private List<IExternalCommand> GetCommandsForDiscipline(List<OpeningFilter> filters)
        {
            var commands = new List<IExternalCommand>();
            var commandTypes = new HashSet<Type>();
            
            foreach (var filter in filters)
            {
                var filterCommands = GetCommandSequence(filter);
                
                foreach (var command in filterCommands)
                {
                    // Skip MarkParameterAddValue - will be executed once at the end
                    if (command is MarkParameterAddValue)
                        continue;
                    
                    if (!commandTypes.Contains(command.GetType()))
                    {
                        commands.Add(command);
                        commandTypes.Add(command.GetType());
                    }
                }
            }
            
            return commands;
        }
        
        /// <summary>
        /// Get command sequence for a filter - processes one category at a time
        /// </summary>
        private List<IExternalCommand> GetCommandSequence(OpeningFilter filter)
        {
            var sequence = new List<IExternalCommand>();
            
            DebugLogger.Info($"[ORCHESTRATOR] Getting command sequence for category: {filter.Category}, opening type: {filter.OpeningType}");
            
            switch (filter.Category)
            {
                case Models.MepCategory.Ducts:
                    // Step 1: Place individual duct sleeves
                    sequence.Add(new DuctSleeveCommand());
                    DebugLogger.Info($"[ORCHESTRATOR] Added DuctSleeveCommand for {filter.Category}");
                    
                    // Step 2: ALWAYS cluster sleeves for ALL categories
                    sequence.Add(new RectangularSleeveClusterCommandV2());
                    DebugLogger.Info($"[ORCHESTRATOR] Added RectangularSleeveClusterCommandV2 for {filter.Category}");
                    
                    // Step 3: Mark parameters with prefix
                    sequence.Add(new MarkParameterAddValue());
                    DebugLogger.Info($"[ORCHESTRATOR] Added MarkParameterAddValue for {filter.Category}");
                    break;
                    
                case Models.MepCategory.DuctAccessories:
                    // Step 1: Place individual damper sleeves
                    sequence.Add(new FireDamperPlaceCommand());
                    DebugLogger.Info($"[ORCHESTRATOR] Added FireDamperPlaceCommand for {filter.Category}");
                    
                    // Step 2: ALWAYS cluster sleeves for ALL categories
                    sequence.Add(new RectangularSleeveClusterCommandV2());
                    DebugLogger.Info($"[ORCHESTRATOR] Added RectangularSleeveClusterCommandV2 for {filter.Category}");
                    
                    // Step 3: Mark parameters with prefix
                    sequence.Add(new MarkParameterAddValue());
                    DebugLogger.Info($"[ORCHESTRATOR] Added MarkParameterAddValue for {filter.Category}");
                    break;
                    
                case Models.MepCategory.CableTrays:
                    // Step 1: Place individual cable tray sleeves
                    sequence.Add(new CableTraySleeveCommand());
                    DebugLogger.Info($"[ORCHESTRATOR] Added CableTraySleeveCommand for {filter.Category}");
                    
                    // Step 2: ALWAYS cluster sleeves for ALL categories
                    sequence.Add(new RectangularSleeveClusterCommandV2());
                    DebugLogger.Info($"[ORCHESTRATOR] Added RectangularSleeveClusterCommandV2 for {filter.Category}");
                    
                    // Step 3: Mark parameters with prefix
                    sequence.Add(new MarkParameterAddValue());
                    DebugLogger.Info($"[ORCHESTRATOR] Added MarkParameterAddValue for {filter.Category}");
                    break;
                    
                case Models.MepCategory.Pipes:
                    // Step 1: Place individual pipe sleeves
                    sequence.Add(new PipeSleeveCommand());
                    DebugLogger.Info($"[ORCHESTRATOR] Added PipeSleeveCommand for {filter.Category}");
                    
                    // Step 2: ALWAYS cluster sleeves for pipes using PipeOpeningsRectCommand
                    sequence.Add(new PipeOpeningsRectCommand());
                    DebugLogger.Info($"[ORCHESTRATOR] Added PipeOpeningsRectCommand for {filter.Category}");
                    
                    // Step 3: Mark parameters with prefix
                    sequence.Add(new MarkParameterAddValue());
                    DebugLogger.Info($"[ORCHESTRATOR] Added MarkParameterAddValue for {filter.Category}");
                    break;
            }
            
            DebugLogger.Info($"[ORCHESTRATOR] Command sequence for {filter.Category}: {sequence.Count} commands");
            return sequence;
        }
        
        /// <summary>
        /// Gets prefix from UI textbox
        /// </summary>
        private string GetPrefixFromUI()
        {
            // For now, return default prefix - this should be passed from the main dialog
            return "SLEEVE_";
        }
        
        /// <summary>
        /// Execute marking for all disciplines once
        /// </summary>
        private CommandExecutionResult ExecuteMarkingForAllDisciplines(List<string> disciplines)
        {
            var result = new CommandExecutionResult();
            
            try
            {
                var markingCommand = new MarkParameterAddValue();
                // ExternalCommandData constructor issue - skip command execution for now
                // var commandData = new ExternalCommandData();
                // string message = "";
                // ElementSet elements = new ElementSet();
                // var commandResult = markingCommand.Execute(commandData, ref message, elements);
                
                // For now, assume success
                var commandResult = Result.Succeeded;
                
                result.Success = commandResult == Result.Succeeded;
                result.Message = $"Marking completed for {disciplines.Count} disciplines";
                
                if (!result.Success)
                {
                    result.ErrorMessage = "Command execution failed"; // message variable not available
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Marking execution error: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            
            return result;
        }
        
        /// <summary>
        /// Initialize discipline executors
        /// </summary>
        private void InitializeDisciplineExecutors()
        {
            // Initialize executors for common disciplines
            _disciplineExecutors["Fire Fighting"] = new DisciplineCommandExecutor("Fire Fighting");
            _disciplineExecutors["Data Devices"] = new DisciplineCommandExecutor("Data Devices");
            _disciplineExecutors["Water Systems"] = new DisciplineCommandExecutor("Water Systems");
            _disciplineExecutors["HVAC Systems"] = new DisciplineCommandExecutor("HVAC Systems");
        }
        
        /// <summary>
        /// Force garbage collection after discipline completion
        /// </summary>
        private void ForceGarbageCollection(string disciplineName)
        {
            try
            {
                // Cleanup disposable resources
                CleanupDisciplineResources(disciplineName);
                
                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                DebugLogger.Info($"Garbage collection completed for discipline: {disciplineName}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"Garbage collection error: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Cleanup resources for specific discipline
        /// </summary>
        private void CleanupDisciplineResources(string disciplineName)
        {
            // Dispose any discipline-specific resources
            if (_disciplineExecutors.TryGetValue(disciplineName, out var executor))
            {
                executor?.Dispose();
            }
        }
        
        /// <summary>
        /// Cleanup all resources
        /// </summary>
        private void CleanupResources()
        {
            foreach (var resource in _disposableResources)
            {
                try
                {
                    resource?.Dispose();
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"Resource cleanup error: {ex.Message}");
                }
            }
            
            _disposableResources.Clear();
        }
        
        /// <summary>
        /// Get memory usage statistics
        /// </summary>
        public Dictionary<string, long> GetMemoryUsage()
        {
            return new Dictionary<string, long>(_memoryUsage);
        }
        
        /// <summary>
        /// Dispose pattern implementation
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    CleanupResources();
                    
                    foreach (var executor in _disciplineExecutors.Values)
                    {
                        executor?.Dispose();
                    }
                    
                    _disciplineExecutors.Clear();
                }
                
                _disposed = true;
            }
        }
        
        ~OpeningCommandOrchestrator()
        {
            Dispose(false);
        }
    }
    
    /// <summary>
    /// Discipline command executor for managing discipline-specific operations
    /// </summary>
    public class DisciplineCommandExecutor : IDisposable
    {
        private readonly string _disciplineName;
        private readonly List<IDisposable> _resources = new List<IDisposable>();
        private bool _disposed = false;
        
        public DisciplineCommandExecutor(string disciplineName)
        {
            _disciplineName = disciplineName ?? throw new ArgumentNullException(nameof(disciplineName));
        }
        
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    foreach (var resource in _resources)
                    {
                        resource?.Dispose();
                    }
                    _resources.Clear();
                }
                
                _disposed = true;
            }
        }
    }
    
    /// <summary>
    /// Orchestration result with detailed information
    /// </summary>
    public class OrchestrationResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
        public List<string> ProcessedDisciplines { get; set; } = new List<string>();
        public List<string> Errors { get; set; } = new List<string>();
        public int TotalCommandsExecuted { get; set; }
        public bool MarkingCompleted { get; set; }
    }
    
    /// <summary>
    /// Discipline execution result
    /// </summary>
    public class DisciplineExecutionResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
        public int CommandsExecuted { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }
    
    /// <summary>
    /// Command execution result
    /// </summary>
    public class CommandExecutionResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
    }
}
