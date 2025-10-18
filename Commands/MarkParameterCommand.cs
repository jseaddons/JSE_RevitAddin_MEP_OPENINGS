using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Mark Parameter Command - Applies MEPMARK to cluster sleeves for a specific category
    /// Can be called from any context (ICommand, ExternalEvent, etc.)
    /// No UI dialogs - fully automated workflow using prefixes configured in main UI
    /// </summary>
    public class MarkParameterCommand : ICommand
    {
        private readonly string _targetCategory;
        private readonly string _projectPrefix;
        private readonly string _disciplinePrefix;
        private readonly bool _remarkAll;
        private readonly MarkPrefixSettings _markPrefixes;
        
        public MarkParameterCommand(string targetCategory, string projectPrefix, string disciplinePrefix, bool remarkAll = false, MarkPrefixSettings markPrefixes = null)
        {
            _targetCategory = targetCategory ?? throw new ArgumentNullException(nameof(targetCategory));
            _projectPrefix = projectPrefix ?? throw new ArgumentNullException(nameof(projectPrefix));
            _disciplinePrefix = disciplinePrefix ?? throw new ArgumentNullException(nameof(disciplinePrefix));
            _remarkAll = remarkAll;
            _markPrefixes = markPrefixes;
        }
        
        public void Execute(UIApplication app)
        {
            // 🔥 CRITICAL DEBUG: Direct file logging to trace MarkParameterCommand execution
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔥 MarkParameterCommand.Execute CALLED 🔥\n");
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] Target Category: {_targetCategory}\n");
            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] Project Prefix: '{_projectPrefix}', Discipline Prefix: '{_disciplinePrefix}', RemarkAll: {_remarkAll}\n");
            
            try
            {
                DebugLogger.Info($"[MarkParameterCommand] Starting MEPMARK application for {_targetCategory}");
                DebugLogger.Info($"[MarkParameterCommand] Project Prefix: '{_projectPrefix}', Discipline Prefix: '{_disciplinePrefix}'");
                
                var doc = app.ActiveUIDocument.Document;
                var uiDoc = app.ActiveUIDocument;
                
                // ✅ FIX: Handle "ALL" category by processing each category individually
                if (_targetCategory.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] Processing ALL categories individually\n");
                    
                    // Get all available categories from XML files
                    var availableCategories = GetAllAvailableCategories();
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] Found {availableCategories.Count} categories: {string.Join(", ", availableCategories)}\n");
                    
                    // Process each category with its specific discipline prefix from UI
                    foreach (var category in availableCategories)
                    {
                        var disciplinePrefix = _markPrefixes?.GetDisciplinePrefix(category) ?? GetDisciplinePrefixForCategory(category);
                        System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] Processing category: {category}, discipline: {disciplinePrefix}\n");
                        
                        using (var tx = new Transaction(doc, $"Mark {category} Clusters"))
                        {
                            tx.Start();
                            
                            var markService = new MarkParameterService();
                            var (processedCount, errorCount) = markService.ApplyMepMarkToClusters(
                                doc, category, _projectPrefix, disciplinePrefix, _remarkAll);
                            
                            tx.Commit();
                            
                            DebugLogger.Info($"[MarkParameterCommand] ✓ MEPMARK complete for {category}: {processedCount} clusters processed, {errorCount} errors");
                            System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\orchestrator_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ✅ Category {category}: {processedCount} processed, {errorCount} errors\n");
                        }
                    }
                }
                else
                {
                    // Process single category
                    using (var tx = new Transaction(doc, $"Mark {_targetCategory} Clusters"))
                    {
                        tx.Start();
                        
                        var markService = new MarkParameterService();
                        var (processedCount, errorCount) = markService.ApplyMepMarkToClusters(
                            doc, _targetCategory, _projectPrefix, _disciplinePrefix, _remarkAll);
                        
                        tx.Commit();
                        
                        DebugLogger.Info($"[MarkParameterCommand] ✓ MEPMARK complete for {_targetCategory}: {processedCount} clusters processed, {errorCount} errors");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterCommand] Error applying MEPMARK to {_targetCategory}: {ex.Message}");
                DebugLogger.Error($"[MarkParameterCommand] Stack trace: {ex.StackTrace}");
                // Don't show TaskDialog here - let orchestrator handle errors
                throw; // Re-throw to let caller handle
            }
        }
        
        /// <summary>
        /// Get all available categories from XML files
        /// </summary>
        private List<string> GetAllAvailableCategories()
        {
            var categories = new List<string>();
            
            try
            {
                var filtersDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                if (!Directory.Exists(filtersDirectory))
                {
                    return categories;
                }

                var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");

                foreach (var xmlFile in xmlFiles)
                {
                    try
                    {
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            if (filter?.ClashZoneStorage?.ClashZones != null)
                            {
                                foreach (var clashZone in filter.ClashZoneStorage.ClashZones)
                                {
                                    if (!categories.Contains(clashZone.MepElementCategory))
                                    {
                                        categories.Add(clashZone.MepElementCategory);
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterCommand] Error getting available categories: {ex.Message}");
            }
            
            return categories;
        }
        
        /// <summary>
        /// Get discipline prefix for a specific category
        /// </summary>
        private string GetDisciplinePrefixForCategory(string category)
        {
            return category switch
            {
                "Ducts" => "DCT",
                "Pipes" => "PLU", 
                "Cable Trays" => "ELE",
                "Duct Accessories" => "DMP",
                _ => "OPN" // Generic fallback
            };
        }
    }
}
