using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Commands;
using static JSE_RevitAddin_MEP_OPENINGS.Models.MepCategoryConstants;

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
            _markPrefixes = markPrefixes ?? throw new ArgumentNullException(nameof(markPrefixes));
            _selectedFilterName = filterName ?? throw new ArgumentNullException(nameof(filterName));
            DebugLogger.Info($"[SleevePlacementExternalEvent] SetContext called - Categories: {string.Join(", ", categories)}, Filter: {filterName}");
        }

        public void Execute(UIApplication app)
        {
            try
            {
                // 🔥 CRITICAL DEBUG: Force direct file logging to bypass any logger issues
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 EXECUTE METHOD CALLED - BUILD TIMESTAMP: {DateTime.Now:yyyy-MM-dd HH:mm:ss} 🔥🔥🔥\n");
                
                // 🔥 CRITICAL DEBUG: Force direct file logging to trace execution
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 1: Execute method started\n");
                
                // ✅ DEBUG: Add immediate logging to confirm Execute is called
                DebugLogger.Info("[SleevePlacementExternalEvent] ===== EXECUTE METHOD CALLED =====");
                DebugLogger.Info($"[SleevePlacementExternalEvent] _selectedCategories is null: {_selectedCategories == null}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] _markPrefixes is null: {_markPrefixes == null}");
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 2: _selectedCategories is null: {_selectedCategories == null}\n");
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 3: _markPrefixes is null: {_markPrefixes == null}\n");
                
                // ✅ CORRECTED: Defensive null check with fallback
                if (_markPrefixes == null)
                {
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                        $"[{DateTime.Now:HH:mm:ss}] STEP 4: Mark prefixes null, using defaults\n");
                    DebugLogger.Warning("[SleevePlacementExternalEvent] Mark prefixes not set, using defaults");
                    _markPrefixes = new MarkPrefixSettings();
                }
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 5: Starting sleeve placement process\n");
                DebugLogger.Info("[SleevePlacementExternalEvent] Starting sleeve placement process");
                
                _uiDocument = app.ActiveUIDocument;
                _document = _uiDocument.Document;
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 6: Got UI document and document\n");
                
                // Log document details for debugging
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - Path: {_document.PathName}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsModifiable: {_document.IsModifiable}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsLinked: {_document.IsLinked}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsWorkshared: {_document.IsWorkshared}");
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 7: Document: {_document.Title}, IsLinked: {_document.IsLinked}\n");
                
                // Ensure we're working with the host document, not a linked file
                // Sleeves must be placed in the host document where structural elements are located
                if (_document.IsLinked)
                {
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                        $"[{DateTime.Now:HH:mm:ss}] STEP 8: ❌ DOCUMENT IS LINKED - RETURNING EARLY\n");
                    var msg = "Cannot place sleeves: Currently active document is a linked file.\n\n" +
                             "Please activate the host document (main project file) and try again.\n" +
                             "Sleeves must be placed in the host document, not in linked files.";
                    DebugLogger.Error($"[SleevePlacementExternalEvent] {msg}");
                    TaskDialog.Show("Wrong Document Active", msg);
                    return;
                }

                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 9: ✅ Document is not linked, continuing\n");

                // ✅ CRITICAL FIX: Check for null _selectedCategories
                if (_selectedCategories == null)
                {
                    DebugLogger.Error("[SleevePlacementExternalEvent] _selectedCategories is null - cannot process");
                    TaskDialog.Show("Error", "No categories selected for processing");
                    return;
                }

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

                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 10: About to create orchestrator\n");
                
                // ✅ ARCHITECTURE COMPLIANCE: Use OpeningCommandOrchestrator for proper command execution
                DebugLogger.Info("[SleevePlacementExternalEvent] Creating OpeningCommandOrchestrator for proper architecture compliance");
                
                var clearanceSettings = GetClearanceSettingsFromUI();
                var orchestrator = new OpeningCommandOrchestrator(_document, _uiDocument, clearanceSettings, _markPrefixes);
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 11: Orchestrator created successfully with clearances and mark prefixes\n");
                
                DebugLogger.Info($"[SleevePlacementExternalEvent] Set {clearanceSettings.Count} UI clearance settings and mark prefixes in orchestrator");
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 12: Set UI clearances in orchestrator\n");
                
                // ✅ Convert categories to filters for orchestrator
                var filters = ConvertCategoriesToFilters(_selectedCategories);
                DebugLogger.Info($"[SleevePlacementExternalEvent] Converted {_selectedCategories.Count} categories to {filters.Count} filters");
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 13: Converted {_selectedCategories.Count} categories to {filters.Count} filters\n");
                
                // ✅ Execute through orchestrator (proper architecture)
                DebugLogger.Info("[SleevePlacementExternalEvent] Executing through OpeningCommandOrchestrator...");
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 14: About to execute orchestrator\n");
                
                orchestrator.ExecuteMultipleFilters(filters, showProgress: true);
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\external_event_execute.log", 
                    $"[{DateTime.Now:HH:mm:ss}] STEP 15: ✅ Orchestrator execution completed\n");
                
                DebugLogger.Info("[SleevePlacementExternalEvent] Orchestrator execution completed");
                
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
                
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
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
                
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                // Search for files matching pattern: *_{category}.xml
                var pattern = $"*_{category.ToLower().Replace(" ", "_")}.xml";
                var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);
                
                if (matchingFiles.Length > 0)
                {
                    // Use the most recently modified file to get the latest clash zones
                    xmlFilePath = matchingFiles
                        .OrderByDescending(f => File.GetLastWriteTime(f))
                        .First();
                    DebugLogger.Info($"[SleevePlacementExternalEvent] Found matching XML file: {xmlFilePath}");
                    
                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                    using (var reader = new StreamReader(xmlFilePath))
                    {
                        var filter = (OpeningFilter)serializer.Deserialize(reader);
                        if (filter?.ClashZoneStorage?.ClashZones != null)
                        {
                            clashZones.AddRange(filter.ClashZoneStorage.ClashZones);
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Loaded {filter.ClashZoneStorage.ClashZones.Count} clash zones from {Path.GetFileName(xmlFilePath)}");
                            
                            // Debug: Check if document titles are populated
                            var clashZonesWithDocTitle = clashZones.Count(cz => !string.IsNullOrEmpty(cz.StructuralElementDocumentTitle));
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Clash zones with document titles: {clashZonesWithDocTitle}/{clashZones.Count}");
                        }
                    }
                }
                else
                {
                    DebugLogger.Warning($"[SleevePlacementExternalEvent] No XML files found matching pattern: {pattern} in directory: {filtersDirectory}");
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
                
                return new UniversalSleevePlacementCommand(_document, clashZones, normalizedCategory, clearanceSettings);
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

    }
}

