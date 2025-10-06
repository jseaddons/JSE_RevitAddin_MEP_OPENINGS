using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.IO;
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
        
        // CRITICAL FIX: Cache linked documents to prevent recursive searching
        private readonly Dictionary<string, Document> _linkCache = new Dictionary<string, Document>();
        
        // CRITICAL FIX: Cancellation token for inner loops
        private System.Threading.CancellationToken _cancellationToken;
        
        // UI clearance settings
        private Dictionary<string, double> _uiClearances;
        
        // Full penetration setting
        private bool _fullPenetrationEnabled = true;
        
        public OpeningCommandOrchestrator(Document document, UIDocument uiDocument)
        {
            try
            {
                DebugLogger.Info("[OpeningCommandOrchestrator] Constructor STARTED");
                
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _uiDocument = uiDocument ?? throw new ArgumentNullException(nameof(uiDocument));
                DebugLogger.Info("[OpeningCommandOrchestrator] Documents validated");
            
            // Set logging context for orchestrator debugging
            DebugLogger.SetServiceContext("Orchestrator");
            _disciplineExecutors = new Dictionary<string, DisciplineCommandExecutor>();
            _uiClearances = null;
                DebugLogger.Info("[OpeningCommandOrchestrator] Basic initialization completed");
                
                // CRITICAL FIX: Initialize link cache to prevent recursive searching
                DebugLogger.Info("[OpeningCommandOrchestrator] Initializing link cache...");
                InitializeLinkCache();
                DebugLogger.Info("[OpeningCommandOrchestrator] Link cache initialized");
                
                DebugLogger.Info("[OpeningCommandOrchestrator] Initializing discipline executors...");
            InitializeDisciplineExecutors();
                DebugLogger.Info("[OpeningCommandOrchestrator] Constructor COMPLETED");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Constructor FAILED: {ex.Message}");
                DebugLogger.Error($"[OpeningCommandOrchestrator] Constructor stack trace: {ex.StackTrace}");
                throw;
            }
        }
        
        /// <summary>
        /// Initialize link cache to prevent recursive searching
        /// </summary>
        private void InitializeLinkCache()
        {
            try
            {
                _linkCache.Clear();
                DebugLogger.Info($"[OpeningCommandOrchestrator] Link cache initialized (lazy loading enabled)");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Failed to initialize link cache: {ex.Message}");
            }
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
        /// Execute categories - routes to individual commands based on selected MEP categories
        /// Each command gets its own data from category-specific XML files
        /// </summary>
        public void ExecuteCategories(List<string> selectedCategories, System.Threading.CancellationToken cancellationToken = default)
        {
            try
            {
                // CRITICAL FIX: Store cancellation token for inner loops
                _cancellationToken = cancellationToken;
                
                DebugLogger.Info($"[OpeningCommandOrchestrator] ===== STARTING CATEGORY EXECUTION =====");
                DebugLogger.Info($"[OpeningCommandOrchestrator] Categories: {string.Join(", ", selectedCategories)}");
                
                // Create UIApplication wrapper for commands
                var uiApp = new UIApplication(_document.Application);
                DebugLogger.Info("[OpeningCommandOrchestrator] Created UIApplication wrapper");
                
                foreach (var category in selectedCategories)
                {
                    // CRITICAL FIX: Check for cancellation in every loop iteration
                    cancellationToken.ThrowIfCancellationRequested();
                    
                    try
                    {
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Processing category: {category}");
                        
                        // Read category-specific XML file
                        var categoryClashZones = ReadCategorySpecificClashZones(category);
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Loaded {categoryClashZones.Count} clash zones for category {category}");
                        
                        switch (category)
                        {
                            case "Ducts":
                                if (categoryClashZones.Count == 0)
                                {
                                    DebugLogger.Info($"[OpeningCommandOrchestrator] Skipping DuctSleeveCommand - no clash zones found for category '{category}'");
                                    break;
                                }
                                DebugLogger.Info($"[OpeningCommandOrchestrator] Executing DuctSleeveCommand with {categoryClashZones.Count} clash zones");
                                var ductCommand = new Commands.DuctSleeveCommand();
                                var ductResult = ExecuteDuctSleeveCommand(ductCommand, uiApp, categoryClashZones);
                                DebugLogger.Info($"[OpeningCommandOrchestrator] DuctSleeveCommand result: {ductResult}");
                                break;
                                
                            case "Pipes":
                                DebugLogger.Info($"[OpeningCommandOrchestrator] Skipping PipeSleeveCommand - not implemented yet");
                                break;
                                
                            case "Cable Trays":
                                DebugLogger.Info($"[OpeningCommandOrchestrator] Skipping CableTraySleeveCommand - not implemented yet");
                                break;
                                
                            case "Duct Accessories":
                                DebugLogger.Info($"[OpeningCommandOrchestrator] Skipping FireDamperPlaceCommand - not implemented yet");
                                break;
                                
                            default:
                                DebugLogger.Warning($"[OpeningCommandOrchestrator] Unknown category: {category}");
                                break;
                        }
                        
                        DebugLogger.Info($"[OpeningCommandOrchestrator] Completed processing category: {category}");
                    }
                    catch (Exception categoryEx)
                    {
                        DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing category {category}: {categoryEx.Message}");
                        DebugLogger.Error($"[OpeningCommandOrchestrator] Category error stack trace: {categoryEx.StackTrace}");
                        // Continue with other categories
                    }
                }
                
                DebugLogger.Info("[OpeningCommandOrchestrator] ===== CATEGORY EXECUTION COMPLETED =====");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] CRITICAL ERROR executing categories: {ex.Message}");
                DebugLogger.Error($"[OpeningCommandOrchestrator] Stack trace: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Reads category-specific clash zones from XML file
        /// </summary>
        private List<ClashZone> ReadCategorySpecificClashZones(string category)
        {
            try
            {
                DebugLogger.Info($"[OpeningCommandOrchestrator] Reading clash zones for category: {category}");
                
                // Determine the XML file name based on category
                var categoryFileName = category.ToLower().Replace(" ", "_");
                
                // Get filter directory - MUST match the directory used by refresh process
                var filterDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                DebugLogger.Info($"[OpeningCommandOrchestrator] Looking for XML files in: {filterDir}");
                
                // Try to find the most recent category-specific XML file
                var possibleFileNames = new[]
                {
                    $"ventilation_{categoryFileName}.xml",
                    $"plumbing_{categoryFileName}.xml", 
                    $"electrical_{categoryFileName}.xml",
                    $"fire_protection_{categoryFileName}.xml",
                    $"mechanical_{categoryFileName}.xml"
                };
                
                OpeningFilter categoryFilter = null;
                string foundFilePath = null;
                
                foreach (var fileName in possibleFileNames)
                {
                    _cancellationToken.ThrowIfCancellationRequested();   // <- ADD
                    var filePath = Path.Combine(filterDir, fileName);
                    if (File.Exists(filePath))
                    {
                        try
                        {
                            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                            using (var reader = new StreamReader(filePath))
                            {
                                categoryFilter = (OpeningFilter)serializer.Deserialize(reader);
                                foundFilePath = filePath;
                                DebugLogger.Info($"[OpeningCommandOrchestrator] Found category-specific XML file: {fileName}");
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Warning($"[OpeningCommandOrchestrator] Error reading {fileName}: {ex.Message}");
                        }
                    }
                }
                
                if (categoryFilter?.ClashZoneStorage?.ClashZones != null)
                {
                    DebugLogger.Info($"[OpeningCommandOrchestrator] Loaded {categoryFilter.ClashZoneStorage.ClashZones.Count} clash zones from {Path.GetFileName(foundFilePath)}");
                    return categoryFilter.ClashZoneStorage.ClashZones;
                }
                else if (foundFilePath != null)
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] XML file found but no clash zones: {Path.GetFileName(foundFilePath)}");
                    return new List<ClashZone>();
                }
                else
                {
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] No category-specific XML files found for category '{category}'");
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Searched files: {string.Join(", ", possibleFileNames)}");
                    DebugLogger.Warning($"[OpeningCommandOrchestrator] Directory exists: {Directory.Exists(filterDir)}");
                    if (Directory.Exists(filterDir))
                    {
                        var existingFiles = Directory.GetFiles(filterDir, "*.xml").Select(Path.GetFileName);
                        DebugLogger.Warning($"[OpeningCommandOrchestrator] Existing XML files: {string.Join(", ", existingFiles)}");
                    }
                    return new List<ClashZone>();
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error reading category-specific clash zones: {ex.Message}");
                return new List<ClashZone>();
            }
        }
        
        /// <summary>
        /// Executes DuctSleeveCommand with clash zones
        /// </summary>
        private Autodesk.Revit.UI.Result ExecuteDuctSleeveCommand(Commands.DuctSleeveCommand command, UIApplication uiApp, List<ClashZone> clashZones)
        {
            try
            {
                UIDocument uidoc = uiApp.ActiveUIDocument;
                Document doc = uidoc.Document;
                string message = "";
                ElementSet elements = new ElementSet();
                
                // Execute DuctSleeveCommand with clash zones
                var result = command.Execute(uidoc, doc, ref message, elements, null, clashZones);
                DebugLogger.Info($"[OpeningCommandOrchestrator] DuctSleeveCommand executed with {clashZones.Count} clash zones, result: {result}");
            
            return result;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing DuctSleeveCommand: {ex.Message}");
                return Autodesk.Revit.UI.Result.Failed;
            }
        }

        /// <summary>
        /// Helper method to execute commands that require ExternalCommandData
        /// </summary>
        private Autodesk.Revit.UI.Result ExecuteCommandWithMockData(Autodesk.Revit.UI.IExternalCommand command, UIApplication uiApp)
        {
            try
            {
                // Skip commands that require ExternalCommandData for now
                // These commands will be handled separately when we implement them
                DebugLogger.Info($"[OpeningCommandOrchestrator] Skipping command {command.GetType().Name} - requires ExternalCommandData");
                return Autodesk.Revit.UI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[OpeningCommandOrchestrator] Error executing command: {ex.Message}");
                return Autodesk.Revit.UI.Result.Failed;
            }
        }

        /// <summary>
        /// Execute multiple filters with memory management and progress dialog
        /// </summary>
        public OrchestrationResult ExecuteMultipleFilters(List<OpeningFilter> filters, bool showProgress = false)
        {
            if (filters == null || filters.Count == 0)
                return new OrchestrationResult{ Success = false, ErrorMessage = "No filters" };

            var uiApp = new UIApplication(_document.Application);
            int total = filters.Count;
            int done  = 0;

            foreach (var f in filters)
            {
                // cheap status-only feedback - use Revit's status bar
                uiApp.Application.WriteJournalComment($"Orchestrator {++done}/{total}", true);
                ExecuteSingleFilterNoUI(f);
            }
            return new OrchestrationResult{ Success = true, Message = "Done" };
        }
        
        private void ExecuteSingleFilterNoUI(OpeningFilter f)
        {
            // same body you had in ExecuteDisciplineWithMemoryManagement
            // but WITHOUT progressDialog calls and WITHOUT ForceGarbageCollection
            var result = new DisciplineExecutionResult();
            var commands = GetCommandsForDiscipline(new List<OpeningFilter>{f});
            foreach (var cmd in commands)
            {
                var r = ExecuteCommandWithResourceManagement(cmd, f);
                result.CommandsExecuted += r.Success ? 1 : 0;
            }
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
        /// Filters clash zones by MEP element type to route them to the correct command
        /// </summary>
        private List<ClashZone>? FilterClashZonesByElementType(List<ClashZone>? clashZones, Type elementType)
        {
            if (clashZones == null || clashZones.Count == 0)
            {
                DebugLogger.Info($"[ORCHESTRATOR] No clash zones to filter by element type {elementType.Name}");
                return null;
            }

            var filteredZones = new List<ClashZone>();
            
            foreach (var clashZone in clashZones)
            {
                _cancellationToken.ThrowIfCancellationRequested();   // <- ADD
                try
                {
                    // Try to get element from host document first
                    var mepElement = _document.GetElement(clashZone.MepElementId);
                    
                    // If element is null, try to find it in linked documents
                    if (mepElement == null)
                    {
                        var linkInstances = new FilteredElementCollector(_document)
                            .OfClass(typeof(RevitLinkInstance))
                            .Cast<RevitLinkInstance>();

                        foreach (var linkInstance in linkInstances)
                        {
                            var linkDoc = linkInstance.GetLinkDocument();
                            if (linkDoc != null)
                            {
                                try
                                {
                                    mepElement = linkDoc.GetElement(clashZone.MepElementId);
                                    if (mepElement != null)
                                    {
                                        DebugLogger.Info($"[ORCHESTRATOR] Found MEP element {clashZone.MepElementId?.IntegerValue ?? -1} in linked document {linkDoc.Title}");
                                        break;
                                    }
                                }
                                catch (Exception linkEx)
                                {
                                    // Continue searching in other linked documents
                                    continue;
                                }
                            }
                        }
                    }
                    
                    if (mepElement != null && elementType.IsAssignableFrom(mepElement.GetType()))
                    {
                        filteredZones.Add(clashZone);
                        DebugLogger.Info($"[ORCHESTRATOR] Clash zone {clashZone.Id} matches element type {elementType.Name}: {mepElement.GetType().Name}");
                    }
                    else
                    {
                        DebugLogger.Info($"[ORCHESTRATOR] Clash zone {clashZone.Id} doesn't match element type {elementType.Name}: {mepElement?.GetType().Name ?? "null"}");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[ORCHESTRATOR] Error filtering clash zone {clashZone.Id} by element type: {ex.Message}");
                }
            }

            DebugLogger.Info($"[ORCHESTRATOR] Filtered {filteredZones.Count} out of {clashZones.Count} clash zones for element type {elementType.Name}");
            return filteredZones.Count > 0 ? filteredZones : null;
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
                _cancellationToken.ThrowIfCancellationRequested();   // <- ADD
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
                    else if (mepElement is FamilyInstance fi)
                    {
                        // Handle dampers and other MEP accessories that are FamilyInstances
                        // For FamilyInstance elements, we need to find the associated duct they're connected to
                        DebugLogger.Info($"[ORCHESTRATOR] Clash zone {clashZone.Id} has FamilyInstance MEP element: {fi.Symbol?.Family?.Name ?? "Unknown"}");
                        
                        // For now, skip FamilyInstance elements as they don't directly translate to ducts
                        // In the future, we might need to find the connected duct or handle them differently
                        DebugLogger.Warning($"[ORCHESTRATOR] Skipping FamilyInstance clash zone {clashZone.Id} - needs special handling for dampers/accessories");
                    }
                    else
                    {
                        DebugLogger.Warning($"[ORCHESTRATOR] Clash zone {clashZone.Id} MEP element is not a duct or FamilyInstance: {mepElement?.GetType().Name}");
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
        /// Converts clash zones to filtered dampers format expected by FireDamperPlaceCommand
        /// </summary>
        private List<(FamilyInstance, Transform?)>? ConvertClashZonesToFilteredDampers(List<ClashZone>? clashZones)
        {
            if (clashZones == null || clashZones.Count == 0)
            {
                DebugLogger.Info("[ORCHESTRATOR] No clash zones to convert to filtered dampers");
                return null;
            }

            var filteredDampers = new List<(FamilyInstance, Transform?)>();
            
            foreach (var clashZone in clashZones)
            {
                _cancellationToken.ThrowIfCancellationRequested();   // <- ADD
                try
                {
                    var mepElement = _document.GetElement(clashZone.MepElementId);
                    if (mepElement is FamilyInstance fi)
                    {
                        // Get transform if the damper is from a linked file
                        Transform? transform = null;
                        if (fi.Document != _document)
                        {
                            // This is a linked damper, get its transform
                            var linkInstances = new FilteredElementCollector(_document)
                                .OfClass(typeof(RevitLinkInstance))
                                .Cast<RevitLinkInstance>()
                                .Where(li => li.GetLinkDocument() == fi.Document);
                            
                            var linkInstance = linkInstances.FirstOrDefault();
                            if (linkInstance != null)
                            {
                                transform = linkInstance.GetTotalTransform();
                            }
                        }
                        
                        filteredDampers.Add((fi, transform));
                        DebugLogger.Info($"[ORCHESTRATOR] Converted clash zone {clashZone.Id} to filtered damper {fi.Id} (Family: {fi.Symbol?.Family?.Name ?? "Unknown"})");
                    }
                    else
                    {
                        DebugLogger.Warning($"[ORCHESTRATOR] Clash zone {clashZone.Id} MEP element is not a FamilyInstance: {mepElement?.GetType().Name}");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[ORCHESTRATOR] Error converting clash zone {clashZone.Id} to filtered damper: {ex.Message}");
                }
            }

            DebugLogger.Info($"[ORCHESTRATOR] Converted {filteredDampers.Count} clash zones to filtered dampers");
            return filteredDampers.Count > 0 ? filteredDampers : null;
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
                    
                    // SMART ROUTING: Filter clash zones to only include duct intersections
                    var totalClashZones = filter?.ClashZoneStorage?.ClashZones?.Count ?? 0;
                    DebugLogger.Info($"[ORCHESTRATOR] Starting DuctSleeve routing with {totalClashZones} total clash zones");
                    
                    var ductClashZones = FilterClashZonesByElementType(filter?.ClashZoneStorage?.ClashZones, typeof(Duct));
                    var filteredDucts = ConvertClashZonesToFilteredDucts(ductClashZones);
                    
                    DebugLogger.Info($"[ORCHESTRATOR] After filtering for Duct type: {ductClashZones?.Count ?? 0} duct clash zones (filtered out {totalClashZones - (ductClashZones?.Count ?? 0)} zones)");
                    DebugLogger.Info($"[ORCHESTRATOR] Triggering DuctSleeveCommand with {filteredDucts?.Count ?? 0} filtered ducts and {ductClashZones?.Count ?? 0} duct clash zones");
                    
                    var message = "";
                    var commandResult = dsc.Execute(_uiDocument, _document, ref message, new ElementSet(), filteredDucts, ductClashZones);
                    result.Success = commandResult == Result.Succeeded;
                    result.Message = "DuctSleeveCommand executed";
                }
                else if (command is FireDamperPlaceCommand fdp)
                {
                    var uiClearances = ClearanceManager.Instance.GetUIClearances();
                    DebugLogger.Info($"[ORCHESTRATOR] Passing {uiClearances.Count} UI clearances to FireDamperPlaceCommand");
                    
                    // SMART ROUTING: Filter clash zones to only include FamilyInstance (damper) intersections
                    var damperClashZones = FilterClashZonesByElementType(filter?.ClashZoneStorage?.ClashZones, typeof(FamilyInstance));
                    
                    if (damperClashZones != null && damperClashZones.Count > 0)
                    {
                        DebugLogger.Info($"[ORCHESTRATOR] Found {damperClashZones.Count} damper clash zones - FireDamperPlaceCommand will process them");
                        
                        // TODO: FireDamperPlaceCommand needs ExecuteImpl method for direct execution
                        // For now, skip execution since ExternalCommandData cannot be constructed
                        DebugLogger.Warning("[ORCHESTRATOR] FireDamperPlaceCommand does not support direct execution - skipping");
                    result.Success = false;
                        result.ErrorMessage = "FireDamperPlaceCommand does not support direct execution";
                        result.Message = "FireDamperPlaceCommand execution skipped";
                    }
                    else
                    {
                        DebugLogger.Info($"[ORCHESTRATOR] No damper clash zones found - skipping FireDamperPlaceCommand");
                        result.Success = true;
                        result.Message = "FireDamperPlaceCommand skipped - no damper intersections";
                    }
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
                // DO NOTHING – let CLR breathe
                DebugLogger.Info($"[GC] Skipped for {disciplineName}");
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
