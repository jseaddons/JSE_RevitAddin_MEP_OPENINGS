using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Commands;

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

        public void Execute(UIApplication app)
        {
            try
            {
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

                // Log immediate feedback (non-blocking)
                DebugLogger.Info($"[SleevePlacementExternalEvent] Processing {_selectedCategories.Count} categories: {string.Join(", ", _selectedCategories)}");

                // Create and queue the appropriate command for each category
                foreach (var category in _selectedCategories)
                {
                    var clashZones = GetClashZonesForCategory(category);
                    if (clashZones.Count > 0)
                    {
                        ICommand command = CreateCommandForCategory(category, clashZones);
                        
                        if (command != null)
                        {
                            // Queue the command for execution (non-blocking)
                            RevitTask.Run(app => command.Execute(app));
                            
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Command queued for category: {category} with {clashZones.Count} clash zones");
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
                
                DebugLogger.Info("[SleevePlacementExternalEvent] All commands queued successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Exception: {ex.Message}");
                DebugLogger.Error($"[SleevePlacementExternalEvent] Stack trace: {ex.StackTrace}");
                TaskDialog.Show("Error", $"Failed to start sleeve placement: {ex.Message}");
            }
        }

        private List<ClashZone> GetClashZonesForCategory(string category)
        {
            // Implementation to read category-specific XML files
            // and return relevant clash zones
            var clashZones = new List<ClashZone>();
            
            try
            {
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                // Search for files matching pattern: *_{category}.xml
                var pattern = $"*_{category.ToLower().Replace(" ", "_")}.xml";
                var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);
                
                if (matchingFiles.Length > 0)
                {
                    // Use the most recently modified file to get the latest clash zones
                    var xmlFile = matchingFiles
                        .OrderByDescending(f => File.GetLastWriteTime(f))
                        .First();
                    DebugLogger.Info($"[SleevePlacementExternalEvent] Found matching XML file: {xmlFile}");
                    
                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                    using (var reader = new StreamReader(xmlFile))
                    {
                        var filter = (OpeningFilter)serializer.Deserialize(reader);
                        if (filter?.ClashZoneStorage?.ClashZones != null)
                        {
                            clashZones.AddRange(filter.ClashZoneStorage.ClashZones);
                            DebugLogger.Info($"[SleevePlacementExternalEvent] Loaded {filter.ClashZoneStorage.ClashZones.Count} clash zones from {xmlFile}");
                            
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
            
            return clashZones;
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
                switch (category.ToLower())
                {
                    case "ducts":
                    case "duct":
                        return new DuctSleevePlacementCommand(_document, clashZones);
                    
                    case "pipes":
                    case "pipe":
                        // TODO: Create PipeSleevePlacementCommand when available
                        DebugLogger.Warning($"[SleevePlacementExternalEvent] PipeSleevePlacementCommand not yet implemented for {clashZones.Count} clash zones");
                        return null;
                    
                    case "cable trays":
                    case "cable tray":
                        // TODO: Create CableTraySleevePlacementCommand when available
                        DebugLogger.Warning($"[SleevePlacementExternalEvent] CableTraySleevePlacementCommand not yet implemented for {clashZones.Count} clash zones");
                        return null;
                    
                    case "duct accessories":
                    case "duct accessory":
                        // TODO: Create DuctAccessorySleevePlacementCommand when available
                        DebugLogger.Warning($"[SleevePlacementExternalEvent] DuctAccessorySleevePlacementCommand not yet implemented for {clashZones.Count} clash zones");
                        return null;
                    
                    default:
                        DebugLogger.Warning($"[SleevePlacementExternalEvent] Unknown category: {category}");
                        return null;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[SleevePlacementExternalEvent] Error creating command for {category}: {ex.Message}");
                return null;
            }
        }

    }
}

