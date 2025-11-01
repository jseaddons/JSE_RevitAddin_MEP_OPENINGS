using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using static JSE_RevitAddin_MEP_OPENINGS.Models.MepCategoryConstants;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// External Event Handler - JUST a transaction context bridge
    /// 
    /// ⚠️ ARCHITECTURE COMPLIANCE REQUIRED ⚠️
    /// This class MUST use OpeningCommandOrchestrator for proper architecture compliance.
    /// 
    /// INTENDED ARCHITECTURE FLOW:
    /// 1. SleevePlacementExternalEvent (External Event Handler) - provides Revit API context
    /// 2. OpeningCommandOrchestrator (THE ACTUAL ORCHESTRATOR) - handles command creation, execution, memory management
    /// 3. UniversalSleevePlacementCommand (per category) - individual sleeve placement
    /// 4. UniversalSleevePlacerService - actual sleeve placement logic
    /// 
    /// ❌ DO NOT BYPASS THE ORCHESTRATOR ❌
    /// ❌ DO NOT CREATE COMMANDS DIRECTLY ❌
    /// ✅ ALWAYS USE OpeningCommandOrchestrator.ExecuteMultipleFilters() ✅
    /// 
    /// Only provides proper Revit API context and tells orchestrator which categories to process
    /// Orchestrator routes to individual commands, commands get their own data
    /// </summary>
    public class SleevePlacementExternalEvent : IExternalEventHandler
    {
        public event Action PlacementCompleted;
        private List<string> _selectedCategories;
        private Document _document;
        private UIDocument _uiDocument;
        private MarkPrefixSettings _markPrefixes;
        private string _selectedFilterName; // ✅ NEW: Store the selected filter name // ✅ NEW: Instance variable for mark prefixes

        /// <summary>
        /// ✅ NEW: Set context for sleeve placement operation
        /// Pass categories, mark prefixes, AND filter name from UI
        /// </summary>
        public void SetContext(List<string> categories, MarkPrefixSettings markPrefixes, string filterName)
        {
            DebugLogger.Info("[SleevePlacementExternalEvent] ===== SETCONTEXT METHOD CALLED =====");
            DebugLogger.Info($"[SleevePlacementExternalEvent] Categories: {string.Join(", ", categories)}");
            DebugLogger.Info($"[SleevePlacementExternalEvent] MarkPrefixes: {markPrefixes?.ProjectPrefix ?? "NULL"}");
            DebugLogger.Info($"[SleevePlacementExternalEvent] FilterName: {filterName ?? "NULL"}");
            
            _selectedCategories = categories ?? throw new ArgumentNullException(nameof(categories));
            _markPrefixes = markPrefixes ?? new MarkPrefixSettings(); // Use defaults if null
            _selectedFilterName = filterName ?? throw new ArgumentNullException(nameof(filterName));
            DebugLogger.Info($"[SleevePlacementExternalEvent] SetContext called - Categories: {string.Join(", ", categories)}, Filter: {filterName}");
        }

        public void Execute(UIApplication app)
        {
            // ✅ CRITICAL: Force direct file write to ensure we see this even if logger fails
            try
            {
                var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 EXECUTE METHOD CALLED - BUILD TIMESTAMP: {DateTime.Now:yyyy-MM-dd HH:mm:ss} 🔥🔥🔥\n");
            }
            catch { }
            
            try
            {
                // 🔥 CRITICAL DEBUG: Force direct file logging to bypass any logger issues
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 EXECUTE METHOD CALLED - BUILD TIMESTAMP: {DateTime.Now:yyyy-MM-dd HH:mm:ss} 🔥🔥🔥\n");
                
                // 🔥 CRITICAL DEBUG: Force direct file logging to trace execution
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 1: Execute method started\n");
                
                // ✅ CRITICAL: Log to file immediately to verify Execute() reaches this point
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 1: Execute method started\n");
                }
                catch { }
                
                // ✅ DEBUG: Add immediate logging to confirm Execute is called
                DebugLogger.Info("[SleevePlacementExternalEvent] ===== EXECUTE METHOD CALLED =====");
                DebugLogger.Info($"[SleevePlacementExternalEvent] _selectedCategories is null: {_selectedCategories == null}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] _markPrefixes is null: {_markPrefixes == null}");
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 2: _selectedCategories is null: {_selectedCategories == null}\n");
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 3: _markPrefixes is null: {_markPrefixes == null}\n");
                
                // ✅ CRITICAL: Log to file at each step
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 2: Categories null={_selectedCategories == null}, Count={_selectedCategories?.Count ?? 0}\n");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 3: MarkPrefixes null={_markPrefixes == null}\n");
                }
                catch { }
                
                // ✅ CORRECTED: Defensive null check with fallback
                if (_markPrefixes == null)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 4: Mark prefixes null, using defaults\n");
                    DebugLogger.Warning("[SleevePlacementExternalEvent] Mark prefixes not set, using defaults");
                    _markPrefixes = new MarkPrefixSettings();
                }
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 5: Starting sleeve placement process\n");
                DebugLogger.Info("[SleevePlacementExternalEvent] Starting sleeve placement process");
                
                _uiDocument = app.ActiveUIDocument;
                _document = _uiDocument.Document;
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 6: Got UI document and document\n");
                
                // Log document details for debugging
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - Path: {_document.PathName}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsModifiable: {_document.IsModifiable}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsLinked: {_document.IsLinked}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsWorkshared: {_document.IsWorkshared}");
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 7: Document: {_document.Title}, IsLinked: {_document.IsLinked}\n");
                
                // Ensure we're working with the host document, not a linked file
                // Sleeves must be placed in the host document where structural elements are located
                if (_document.IsLinked)
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 8: ❌ DOCUMENT IS LINKED - RETURNING EARLY\n");
                    try
                    {
                        var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 8: ❌ DOCUMENT IS LINKED - RETURNING EARLY\n");
                    }
                    catch { }
                    var msg = "Cannot place sleeves: Currently active document is a linked file.\n\n" +
                             "Please activate the host document (main project file) and try again.\n" +
                             "Sleeves must be placed in the host document, not in linked files.";
                    DebugLogger.Error($"[SleevePlacementExternalEvent] {msg}");
                    TaskDialog.Show("Wrong Document Active", msg);
                    return;
                }

                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 9: ✅ Document is not linked, continuing\n");
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 9: ✅ Document is not linked, continuing\n");
                }
                catch { }

                // ✅ CRITICAL FIX: Check for null _selectedCategories
                if (_selectedCategories == null)
                {
                    DebugLogger.Error("[SleevePlacementExternalEvent] _selectedCategories is null - cannot process");
                    try
                    {
                        var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌ _selectedCategories is NULL - RETURNING EARLY\n");
                    }
                    catch { }
                    TaskDialog.Show("Error", "No categories selected for processing");
                    return;
                }
                
                // ✅ CRITICAL: Log categories count
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 9.5: Categories count: {_selectedCategories.Count}, Categories: {string.Join(", ", _selectedCategories)}\n");
                }
                catch { }

                // Log immediate feedback (non-blocking)
                DebugLogger.Info($"[SleevePlacementExternalEvent] Processing {_selectedCategories.Count} categories: {string.Join(", ", _selectedCategories)}");

                // ✅ DEBUG: Add try-catch around LoadClusterConfigurationFromFilters
                try
                {
                    DebugLogger.Info("[SleevePlacementExternalEvent] About to call LoadClusterConfigurationFromFilters...");
                    LoadClusterConfigurationFromFilters();
                    DebugLogger.Info("[SleevePlacementExternalEvent] LoadClusterConfigurationFromFilters completed successfully");
                }
                catch (Exception loadEx)
                {
                    DebugLogger.Error($"[SleevePlacementExternalEvent] Error in LoadClusterConfigurationFromFilters: {loadEx.Message}");
                    DebugLogger.Error($"[SleevePlacementExternalEvent] LoadClusterConfigurationFromFilters stack trace: {loadEx.StackTrace}");
                    throw; // Re-throw to be caught by outer try-catch
                }

                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 10: About to create orchestrator\n");
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 10: About to create orchestrator\n");
                }
                catch { }
                
                // ✅ ARCHITECTURE COMPLIANCE: Use OpeningCommandOrchestrator for proper command execution
                DebugLogger.Info("[SleevePlacementExternalEvent] Creating OpeningCommandOrchestrator for proper architecture compliance");
                
                var clearanceSettings = GetClearanceSettingsFromUI();
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 10.5: Clearance settings count: {clearanceSettings?.Count ?? 0}\n");
                }
                catch { }
                
                var orchestrator = new OpeningCommandOrchestrator(_document, _uiDocument, clearanceSettings, _markPrefixes);
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 11: Orchestrator created successfully with clearances and mark prefixes\n");
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 11: Orchestrator created successfully\n");
                }
                catch { }
                
                DebugLogger.Info($"[SleevePlacementExternalEvent] Set {clearanceSettings.Count} UI clearance settings and mark prefixes in orchestrator");
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 12: Set UI clearances in orchestrator\n");
                
                // ✅ Convert categories to filters for orchestrator
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 12.5: About to convert categories to filters\n");
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 12.5: About to convert {_selectedCategories.Count} categories to filters\n");
                }
                catch { }
                
                var filters = ConvertCategoriesToFilters(_selectedCategories);
                DebugLogger.Info($"[SleevePlacementExternalEvent] Converted {_selectedCategories.Count} categories to {filters.Count} filters");
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 13: Converted {_selectedCategories.Count} categories to {filters.Count} filters\n");
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 13: Converted to {filters.Count} filters\n");
                    if (filters != null && filters.Count > 0)
                    {
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] STEP 13.5: Filter names: {string.Join(", ", filters.Select(f => f.Name))}\n");
                    }
                }
                catch { }
                
                // ✅ Execute through orchestrator (proper architecture)
                DebugLogger.Info("[SleevePlacementExternalEvent] Executing through OpeningCommandOrchestrator...");
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 14: About to execute orchestrator\n");
                
                // ✅ CRITICAL: Log before and after orchestrator call to see if it completes
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 BEFORE orchestrator.ExecuteMultipleFilters() - filters count: {filters?.Count ?? 0}\n");
                }
                catch { }
                
                orchestrator.ExecuteMultipleFilters(filters, showProgress: true);
                
                // ✅ CRITICAL: Log after orchestrator returns to confirm it completed
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔥 AFTER orchestrator.ExecuteMultipleFilters() - returned successfully\n");
                }
                catch { }
                
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 15: ✅ Orchestrator execution completed\n");
                
                DebugLogger.Info("[SleevePlacementExternalEvent] Orchestrator execution completed");
                
                // ✅ SIMPLIFIED: Coordinate saving now handled directly in UniversalSleevePlacerService
                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] STEP 16: ✅ Coordinate saving completed in UniversalSleevePlacerService\n");
                
                DebugLogger.Info("[SleevePlacementExternalEvent] Coordinate saving completed in UniversalSleevePlacerService");
                
                DebugLogger.Info("[SleevePlacementExternalEvent] ✓ COMPLETED THROUGH PROPER ORCHESTRATOR ARCHITECTURE");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Exception: {ex.Message}");
                DebugLogger.Error($"[SleevePlacementExternalEvent] Stack trace: {ex.StackTrace}");
                TaskDialog.Show("Error", $"Failed to complete sleeve placement: {ex.Message}");
            }
            finally
            {
                try { PlacementCompleted?.Invoke(); } catch { }
            }
        }

        /// <summary>
        /// ✅ NEW: Get clearance settings from UI for orchestrator
        /// </summary>
        private Dictionary<string, double> GetClearanceSettingsFromUI()
        {
            try
            {
                var clearanceSettings = new Dictionary<string, double>();
                
                // ✅ CRITICAL FIX: Check for null _selectedCategories
                if (_selectedCategories == null)
                {
                    DebugLogger.Error("[SleevePlacementExternalEvent] _selectedCategories is null in GetClearanceSettingsFromUI");
                    return clearanceSettings;
                }
                
                // Get clearance settings from FilterUiStateProvider
                foreach (var category in _selectedCategories)
                {
                    var categorySettings = FilterUiStateProvider.GetClearanceSettings?.Invoke(category) ?? new Dictionary<string, double>();
                    foreach (var kvp in categorySettings)
                    {
                        clearanceSettings[kvp.Key] = kvp.Value;
                    }
                }
                
                DebugLogger.Info($"[SleevePlacementExternalEvent] Collected {clearanceSettings.Count} clearance settings from UI");
                return clearanceSettings;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Error getting clearance settings: {ex.Message}");
                return new Dictionary<string, double>();
            }
        }
        
        /// <summary>
        /// ✅ NEW: Convert categories to filters for orchestrator
        /// </summary>
        private List<OpeningFilter> ConvertCategoriesToFilters(List<string> categories)
        {
            try
            {
                var filters = new List<OpeningFilter>();
                
                // ✅ CRITICAL FIX: Check for null categories parameter
                if (categories == null)
                {
                    DebugLogger.Error("[SleevePlacementExternalEvent] categories parameter is null in ConvertCategoriesToFilters");
                    return filters;
                }
                
                foreach (var categoryName in categories)
                {
                    var (clashZones, xmlFilePath) = GetClashZonesForCategory(categoryName);
                    if (clashZones.Count > 0)
                    {
                        // Convert string category to MepCategory enum
                        var mepCategory = MepCategoryConstants.Parse(categoryName);
                        
                        var filter = new OpeningFilter
                        {
                            Name = _selectedFilterName, // ✅ FIXED: Use actual filter name from UI
                            Category = mepCategory,
                            SelectedMepCategoryName = categoryName,
                            IsEnabled = true,
                            CreatedDate = DateTime.Now,
                            LastModified = DateTime.Now
                        };
                        filters.Add(filter);
                        DebugLogger.Info($"[SleevePlacementExternalEvent] Created filter '{filter.Name}' for category '{categoryName}' with {clashZones.Count} clash zones");
                    }
                }
                
                return filters;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Error converting categories to filters: {ex.Message}");
                return new List<OpeningFilter>();
            }
        }

        /// <summary>
        /// ⚠️ CRITICAL: Load cluster configuration from filter settings
        /// This must be called BEFORE executing any sleeve placement commands
        /// to ensure cluster command respects user's JoinOpeningsDistance setting
        /// </summary>
        private void LoadClusterConfigurationFromFilters()
        {
            try
            {
                DebugLogger.Info("[SleevePlacementExternalEvent] Loading cluster configuration from filters...");
                
                // ✅ CRITICAL FIX: Use project-specific directory (matches Refresh saves)
                // This was using hardcoded "Default" path
                var filtersDirectory = _document != null 
                    ? ProjectPathService.GetFiltersDirectory(_document)
                    : ProjectPathService.GetFiltersDirectory(_document);
                
                DebugLogger.Info($"[SleevePlacementExternalEvent] Loading cluster config from: {filtersDirectory}");
                
                // ✅ Ensure directory exists before searching
                if (!Directory.Exists(filtersDirectory))
                {
                    DebugLogger.Warning($"[SleevePlacementExternalEvent] Filters directory does not exist: {filtersDirectory} - using default cluster configuration");
                    ClusterConfigurationManager.Instance.SetJoinOpeningsDistance(200.0, "Default (directory not found)");
                    return;
                }
                
                // Search for any filter XML file to extract advanced settings
                var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                
                if (xmlFiles.Length > 0)
                {
                    // Use the most recently modified file
                    var xmlFile = xmlFiles
                        .OrderByDescending(f => File.GetLastWriteTime(f))
                        .First();
                    
                    DebugLogger.Info($"[SleevePlacementExternalEvent] Reading configuration from: {Path.GetFileName(xmlFile)}");
                    
                    // Try to load as UserConfiguration first (which contains AdvancedSettings)
                    try
                    {
                        var userConfigSerializer = new System.Xml.Serialization.XmlSerializer(typeof(Models.UserConfiguration));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var userConfig = (Models.UserConfiguration)userConfigSerializer.Deserialize(reader);
                            
                            if (userConfig?.AdvancedSettings != null)
                            {
                                var joinDistance = userConfig.AdvancedSettings.JoinOpeningsDistance;
                                
                                // Set cluster configuration manager
                                ClusterConfigurationManager.Instance.SetJoinOpeningsDistance(
                                    joinDistance, 
                                    $"UserConfiguration: {xmlFile}");
                                
                                DebugLogger.Info($"[SleevePlacementExternalEvent] ✓ Cluster configuration loaded from UserConfiguration: JoinOpeningsDistance = {joinDistance}mm");
                                return; // Success - exit method
                            }
                        }
                    }
                    catch
                    {
                        // Not a UserConfiguration file, try OpeningFilter format
                    }
                    
                    // Fallback: Try OpeningFilter format
                    try
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Loaded filter: {filter?.Name}, using default cluster configuration (200mm)");
                        }
                    }
                    catch
                    {
                        // Ignore deserialization errors
                    }
                    
                    // No advanced settings found - use default
                    DebugLogger.Warning("[SleevePlacementExternalEvent] No AdvancedSettings found in XML files - using default cluster configuration (200mm)");
                    ClusterConfigurationManager.Instance.SetJoinOpeningsDistance(200.0, "Default (no settings in filter)");
                }
                else
                {
                    DebugLogger.Warning($"[SleevePlacementExternalEvent] No filter XML files found in {filtersDirectory} - using default cluster configuration (200mm)");
                    ClusterConfigurationManager.Instance.SetJoinOpeningsDistance(200.0, "Default (no filters found)");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Error loading cluster configuration: {ex.Message}");
                DebugLogger.Error($"[SleevePlacementExternalEvent] Using default cluster configuration (200mm)");
                ClusterConfigurationManager.Instance.SetJoinOpeningsDistance(200.0, "Default (error loading from filter)");
            }
        }

        private (List<ClashZone> clashZones, string xmlFilePath) GetClashZonesForCategory(string category)
        {
            // Implementation to read category-specific XML files
            // and return relevant clash zones WITH the source XML file path
            var clashZones = new List<ClashZone>();
            string xmlFilePath = string.Empty;
            
            try
            {
                // ✅ CRITICAL FIX: Check for null category parameter
                if (string.IsNullOrEmpty(category))
                {
                    DebugLogger.Error("[SleevePlacementExternalEvent] category parameter is null or empty in GetClashZonesForCategory");
                    return (clashZones, xmlFilePath);
                }
                
                // ✅ CRITICAL FIX: Use project-specific directory (matches Refresh saves)
                var filtersDirectory = _document != null 
                    ? ProjectPathService.GetFiltersDirectory(_document)
                    : ProjectPathService.GetFiltersDirectory(_document);
                
                // ✅ CRITICAL: Log which directory we're looking in
                DebugLogger.Info($"[SleevePlacementExternalEvent] GetClashZonesForCategory: Looking in directory: {filtersDirectory}");
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] GetClashZonesForCategory: category='{category}', directory='{filtersDirectory}', filterName='{_selectedFilterName}'\n");
                }
                catch { }
                
                // ✅ CRITICAL FIX: Use MepCategoryConstants.GetXmlSuffix() for consistent naming
                // This matches RefreshService normalization EXACTLY
                var categoryPattern = GetXmlSuffix(category);
                
                // ✅ EXACT FILE NAME MATCH: Extract base filter name (remove any category suffix) to match RefreshService
                // RefreshService saves as: {baseFilterName}_{normalizedCategory}.xml
                // We must load as: {baseFilterName}_{categoryPattern}.xml
                string baseFilterName = ExtractBaseFilterName(_selectedFilterName, categoryPattern);
                
                // ✅ EXACT FILE NAME MATCH: No variations, no case-insensitive, no fallbacks
                var exactFileName = $"{baseFilterName}_{categoryPattern}.xml";
                xmlFilePath = Path.Combine(filtersDirectory, exactFileName);
                
                DebugLogger.Info($"[SleevePlacementExternalEvent] Looking for EXACT file: '{exactFileName}'");
                try
                {
                    var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] GetClashZonesForCategory: Looking for EXACT file: '{exactFileName}'\n");
                }
                catch { }
                
                // ✅ EXACT FILE EXISTENCE CHECK: No variations, no fallbacks
                if (!File.Exists(xmlFilePath))
                {
                    DebugLogger.Warning($"[SleevePlacementExternalEvent] ❌ EXACT file not found: '{exactFileName}' in directory '{filtersDirectory}'");
                    try
                    {
                        var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                        if (!Directory.Exists(filtersDirectory))
                        {
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌ GetClashZonesForCategory: Directory does not exist: {filtersDirectory}\n");
                }
                else
                {
                            var allFiles = Directory.GetFiles(filtersDirectory, "*.xml");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌ GetClashZonesForCategory: EXACT file '{exactFileName}' not found. Directory has {allFiles.Length} XML files.\n");
                            if (allFiles.Length > 0)
                    {
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] GetClashZonesForCategory: Available files: {string.Join(", ", allFiles.Select(Path.GetFileName))}\n");
                            }
                        }
                    }
                    catch { }
                    return (clashZones, string.Empty);
                }
                
                if (!string.IsNullOrEmpty(xmlFilePath))
                {
                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                    using (var reader = new StreamReader(xmlFilePath))
                    {
                        var filter = (OpeningFilter)serializer.Deserialize(reader);
                        if (filter?.ClashZoneStorage?.ClashZones != null)
                        {
                            clashZones.AddRange(filter.ClashZoneStorage.ClashZones);
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Loaded {filter.ClashZoneStorage.ClashZones.Count} clash zones from {Path.GetFileName(xmlFilePath)}");
                            
                            // ✅ CRITICAL: Log to file
                            try
                            {
                                var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] GetClashZonesForCategory: Loaded {clashZones.Count} clash zones from {Path.GetFileName(xmlFilePath)}\n");
                            }
                            catch { }
                            
                            // Debug: Check if document titles are populated
                            var clashZonesWithDocTitle = clashZones.Count(cz => !string.IsNullOrEmpty(cz.StructuralElementDocumentTitle));
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Clash zones with document titles: {clashZonesWithDocTitle}/{clashZones.Count}");
                        }
                        else
                        {
                            // ✅ CRITICAL: Log if filter or clash zones are null
                            try
                            {
                                var logPath = SafeFileLogger.GetLogFilePath("placement_event_trace.log");
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌ GetClashZonesForCategory: Filter or ClashZoneStorage is null in file {Path.GetFileName(xmlFilePath)}\n");
                            }
                            catch { }
                            DebugLogger.Warning($"[SleevePlacementExternalEvent] Filter or ClashZoneStorage is null in file {xmlFilePath}");
                        }
                    }
                }
                else
                {
                    DebugLogger.Warning($"[SleevePlacementExternalEvent] No XML files found for category '{category}' with filter '{_selectedFilterName}' in directory: {filtersDirectory}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Error loading clash zones for {category}: {ex.Message}");
            }
            
            return (clashZones, xmlFilePath);
        }

        public string GetName()
        {
            DebugLogger.Info("[SleevePlacementExternalEvent] GetName() called - returning 'Sleeve Placement Handler'");
            return "Sleeve Placement Handler";
        }

        public void SetSelectedCategories(List<string> selectedCategories)
        {
            _selectedCategories = selectedCategories;
            DebugLogger.Info($"[SleevePlacementExternalEvent] Set categories for processing: {string.Join(", ", selectedCategories)}");
        }

        /// <summary>
        /// ✅ NEW: Get category priority for sequential processing
        /// Lower number = higher priority (processed first)
        /// </summary>
        private int GetCategoryPriority(string mepCategory)
        {
            return mepCategory?.ToLowerInvariant() switch
            {
                "duct accessories" => 1, // Dampers - HIGHEST priority
                "ducts" => 2, // Ducts - processed after dampers
                "pipes" => 3,
                "cable trays" => 4,
                "cable tray fittings" => 5,
                _ => 999 // Unknown categories last
            };
        }


        private ICommand CreateCommandForCategory(string category, List<ClashZone> clashZones)
        {
            try
            {
                // ✅ ALL categories now use UniversalSleevePlacementCommand
                // Normalize category names to Revit API standard (prevent Duct/Ducts confusion)
                string normalizedCategory = MepCategoryConstants.Normalize(category);
                
                DebugLogger.Info($"[SleevePlacementExternalEvent] Category '{category}' normalized to '{normalizedCategory}'");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Creating UniversalSleevePlacementCommand for {normalizedCategory} with {clashZones.Count} clash zones");
                
                // ✅ NEW: Get clearance settings from UI for this category
                var clearanceSettings = GetClearanceSettingsForCategory(normalizedCategory);
                
                return new UniversalSleevePlacementCommand(_document, clashZones, normalizedCategory, _selectedFilterName, clearanceSettings);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Error creating command for {category}: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Get clearance settings from UI for a specific category
        /// </summary>
        private Dictionary<string, double> GetClearanceSettingsForCategory(string category)
        {
            try
            {
                // Use FilterUiStateProvider to get clearance settings from UI
                var clearanceSettings = FilterUiStateProvider.GetClearanceSettings?.Invoke(category) ?? new Dictionary<string, double>();
                
                DebugLogger.Info($"[SleevePlacementExternalEvent] Retrieved {clearanceSettings.Count} clearance settings for category '{category}'");
                foreach (var kvp in clearanceSettings)
                {
                    DebugLogger.Info($"[SleevePlacementExternalEvent] Clearance: {kvp.Key} = {kvp.Value}mm");
                }
                
                return clearanceSettings;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Error getting clearance settings for {category}: {ex.Message}");
                return new Dictionary<string, double>();
            }
        }
        
        /// <summary>
        /// Extracts the base filter name by removing any existing category suffix.
        /// E.g., "Ventilation_ducts" → "Ventilation" (if normalizedCategory is "ducts")
        /// This matches RefreshService.ExtractBaseFilterName logic exactly.
        /// </summary>
        private string ExtractBaseFilterName(string filterName, string normalizedCategory)
        {
            if (string.IsNullOrWhiteSpace(filterName)) return filterName;
            if (string.IsNullOrWhiteSpace(normalizedCategory)) return filterName;
            
            // Check if filter name ends with the category suffix
            string suffixPattern = $"_{normalizedCategory}";
            if (filterName.EndsWith(suffixPattern, StringComparison.OrdinalIgnoreCase))
            {
                return filterName.Substring(0, filterName.Length - suffixPattern.Length);
            }
            
            return filterName;
        }

    }
}

