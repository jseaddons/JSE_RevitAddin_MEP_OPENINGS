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
    /// Only provides proper Revit API context and tells orchestrator which categories to process
    /// Orchestrator routes to individual commands, commands get their own data
    /// </summary>
    public class SleevePlacementExternalEvent : IExternalEventHandler
    {
        private List<string> _selectedCategories;
        private Document _document;
        private UIDocument _uiDocument;
        private MarkPrefixSettings _markPrefixes; // ✅ NEW: Instance variable for mark prefixes

        /// <summary>
        /// ✅ NEW: Set context for sleeve placement operation
        /// Pass both categories AND mark prefixes from UI
        /// </summary>
        public void SetContext(List<string> categories, MarkPrefixSettings markPrefixes)
        {
            _selectedCategories = categories ?? throw new ArgumentNullException(nameof(categories));
            _markPrefixes = markPrefixes ?? throw new ArgumentNullException(nameof(markPrefixes));
            DebugLogger.Info($"[SleevePlacementExternalEvent] SetContext called - Categories: {string.Join(", ", categories)}");
        }

        public void Execute(UIApplication app)
        {
            try
            {
                // ✅ CORRECTED: Defensive null check with fallback
                if (_markPrefixes == null)
                {
                    DebugLogger.Warning("[SleevePlacementExternalEvent] Mark prefixes not set, using defaults");
                    _markPrefixes = new MarkPrefixSettings();
                }
                
                DebugLogger.Info("[SleevePlacementExternalEvent] Starting sleeve placement process");
                
                _uiDocument = app.ActiveUIDocument;
                _document = _uiDocument.Document;
                
                // Log document details for debugging
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - Path: {_document.PathName}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsModifiable: {_document.IsModifiable}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsLinked: {_document.IsLinked}");
                DebugLogger.Info($"[SleevePlacementExternalEvent] Active document - IsWorkshared: {_document.IsWorkshared}");
                
                // Ensure we're working with the host document, not a linked file
                // Sleeves must be placed in the host document where structural elements are located
                if (_document.IsLinked)
                {
                    var msg = "Cannot place sleeves: Currently active document is a linked file.\n\n" +
                             "Please activate the host document (main project file) and try again.\n" +
                             "Sleeves must be placed in the host document, not in linked files.";
                    DebugLogger.Error($"[SleevePlacementExternalEvent] {msg}");
                    TaskDialog.Show("Wrong Document Active", msg);
                    return;
                }

                // ⚠️ CRITICAL: Load cluster configuration from filter settings BEFORE executing commands
                LoadClusterConfigurationFromFilters();

                // Log immediate feedback (non-blocking)
                DebugLogger.Info($"[SleevePlacementExternalEvent] Processing {_selectedCategories.Count} categories: {string.Join(", ", _selectedCategories)}");

                // Process each category: Place individual sleeves → Cluster → Update XML
                foreach (var category in _selectedCategories)
                {
                    var (clashZones, xmlFilePath) = GetClashZonesForCategory(category);
                    if (clashZones.Count > 0)
                    {
                        // Step 1: Place individual sleeves
                        ICommand placementCommand = CreateCommandForCategory(category, clashZones);
                        
                        if (placementCommand != null)
                        {
                            // Step 1: Place individual sleeves
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Placing individual sleeves for {category} ({clashZones.Count} clash zones)");
                            placementCommand.Execute(app);
                            
                        // Step 2: Immediately cluster this category's sleeves
                        DebugLogger.Info($"[SleevePlacementExternalEvent] Clustering {category} sleeves...");
                        var clusterCommand = new Commands.UniversalClusterCommand(category, xmlFilePath);
                        clusterCommand.Execute(app);
                            
                            // ✅ NEW: Step 3: Apply MEPMARK to clusters using stored prefixes
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Applying MEPMARK to {category} clusters...");
                            
                            // ✅ CORRECTED: Use instance variable (safe after null check)
                            string projectPrefix = _markPrefixes.ProjectPrefix;
                            string disciplinePrefix = _markPrefixes.GetDisciplinePrefix(category);
                            bool remarkAll = _markPrefixes.RemarkAll;
                            
                            var markCommand = new Commands.MarkParameterCommand(category, projectPrefix, disciplinePrefix, remarkAll);
                            markCommand.Execute(app);
                            
                            DebugLogger.Info($"[SleevePlacementExternalEvent] ✓ Completed placement, clustering, and MEPMARK for {category}");
                        }
                        else
                        {
                            DebugLogger.Warning($"[SleevePlacementExternalEvent] No command available for category: {category}");
                        }
                    }
                    else
                    {
                        DebugLogger.Warning($"[SleevePlacementExternalEvent] No clash zones found for category: {category}");
                    }
                }
                
                DebugLogger.Info("[SleevePlacementExternalEvent] All categories processed (placement + clustering + MEPMARK)");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Exception: {ex.Message}");
                DebugLogger.Error($"[SleevePlacementExternalEvent] Stack trace: {ex.StackTrace}");
                TaskDialog.Show("Error", $"Failed to complete sleeve placement: {ex.Message}");
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
            return "Sleeve Placement Handler";
        }

        public void SetSelectedCategories(List<string> selectedCategories)
        {
            _selectedCategories = selectedCategories;
            DebugLogger.Info($"[SleevePlacementExternalEvent] Set categories for processing: {string.Join(", ", selectedCategories)}");
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
                
                return new UniversalSleevePlacementCommand(_document, clashZones, normalizedCategory);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Error creating command for {category}: {ex.Message}");
                return null;
            }
        }

    }
}

