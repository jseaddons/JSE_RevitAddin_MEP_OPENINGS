using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

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
        
        public MarkParameterCommand(string targetCategory, string projectPrefix, string disciplinePrefix, bool remarkAll = false, MarkPrefixSettings? markPrefixes = null)
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
            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔥 MarkParameterCommand.Execute CALLED 🔥\n");
            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Target Category: {_targetCategory}\n");
            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Project Prefix: '{_projectPrefix}', Discipline Prefix: '{_disciplinePrefix}', RemarkAll: {_remarkAll}\n");

            try
            {
                DebugLogger.Info($"[MarkParameterCommand] Starting MEPMARK application for {_targetCategory}");
                DebugLogger.Info($"[MarkParameterCommand] Project Prefix: '{_projectPrefix}', Discipline Prefix: '{_disciplinePrefix}'");

                var doc = app.ActiveUIDocument.Document;
                var uiDoc = app.ActiveUIDocument;

                // ✅ FIX: Handle "ALL" category by processing each category individually
                if (_targetCategory.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                {
                    DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Processing ALL categories individually\n");

                    // Get all available categories from XML files
                    var availableCategories = GetAllAvailableCategories(doc);
                    
                    // ✅ DEBUG: Log available categories
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string orchestratorLogPath = SafeFileLogger.GetLogFilePath("orchestrator_debug.log");
                        System.IO.File.AppendAllText(orchestratorLogPath,
                        $"[{DateTime.Now:HH:mm:ss}] Found {availableCategories.Count} categories: {string.Join(", ", availableCategories)}\n");
                    }

                    // ✅ BIM 360 OPTIMIZATION: Use SINGLE transaction for ALL categories
                    // This reduces cloud sync overhead from N syncs to 1 sync (25-50x faster on BIM 360)
                    using (var tx = new Transaction(doc, "Mark All Categories"))
                    {
                        tx.Start();
                        
                        int totalProcessed = 0;
                        int totalErrors = 0;

                        // ✅ CRITICAL FIX: Process each category OUTSIDE the deployment mode check
                        // Process each category with its specific discipline prefix from UI
                        foreach (var category in availableCategories)
                        {
                            var disciplinePrefix = _markPrefixes?.GetDisciplinePrefix(category) ?? GetDisciplinePrefixForCategory(category);
                            // ✅ FIX: Get remark flag per category from MarkPrefixSettings
                            var remarkFlag = _markPrefixes?.GetRemarkFlag(category) ?? _remarkAll;
                            
                            // ✅ DEBUG: Log remark flag details
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Processing category: {category}, discipline: {disciplinePrefix}, remark: {remarkFlag}");
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}]   RemarkAll={_markPrefixes?.RemarkAll ?? false}, RemarkProjectPrefix={_markPrefixes?.RemarkProjectPrefix ?? false}");
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}]   RemarkDuctPrefix={_markPrefixes?.RemarkDuctPrefix ?? false}, RemarkPipePrefix={_markPrefixes?.RemarkPipePrefix ?? false}");
                                DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}]   RemarkCableTrayPrefix={_markPrefixes?.RemarkCableTrayPrefix ?? false}, RemarkDamperPrefix={_markPrefixes?.RemarkDamperPrefix ?? false}\n");
                            }
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Processing category: {category}, discipline: {disciplinePrefix}, remark: {remarkFlag}\n");

                            var markService = new MarkParameterService();
                            var numberFormat = _markPrefixes?.NumberFormat ?? "000";
                            var (processedCount, errorCount) = markService.ApplyMepMarkToClusters(
                                doc, category, _projectPrefix, disciplinePrefix, remarkFlag, numberFormat, _markPrefixes);

                            totalProcessed += processedCount;
                            totalErrors += errorCount;

                            DebugLogger.Info($"[MarkParameterCommand] ✓ MEPMARK complete for {category}: {processedCount} clusters processed, {errorCount} errors");
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Category {category}: {processedCount} processed, {errorCount} errors\n");
                        }
                        
                        // ✅ BIM 360 OPTIMIZATION: Single commit for all categories
                        tx.Commit(); // Cloud sync happens ONCE for all categories
                        
                        DebugLogger.Info($"[MarkParameterCommand] ✅ ALL CATEGORIES COMPLETE: {totalProcessed} total processed, {totalErrors} total errors");
                    }
                }
                else
                    {
                        // Process single category
                        // ✅ CRITICAL FIX: Get remark flag per category from MarkPrefixSettings (not just _remarkAll)
                        // This ensures that when Remark Selected button is clicked, each category's checkbox state is respected
                        var remarkFlag = _markPrefixes?.GetRemarkFlag(_targetCategory) ?? _remarkAll;
                        DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Processing single category: {_targetCategory}, discipline: {_disciplinePrefix}, remark: {remarkFlag}\n");
                        
                        using (var tx = new Transaction(doc, $"Mark {_targetCategory} Clusters"))
                        {
                            tx.Start();

                            var markService = new MarkParameterService();
                            var numberFormat = _markPrefixes?.NumberFormat ?? "000";
                            var (processedCount, errorCount) = markService.ApplyMepMarkToClusters(
                                doc, _targetCategory, _projectPrefix, _disciplinePrefix, remarkFlag, numberFormat, _markPrefixes);

                            tx.Commit();

                            DebugLogger.Info($"[MarkParameterCommand] ✓ MEPMARK complete for {_targetCategory}: {processedCount} clusters processed, {errorCount} errors");
                            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] ✅ Category {_targetCategory}: {processedCount} processed, {errorCount} errors\n");
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
        /// ✅ DATABASE-BASED: Get all available categories from the database
        /// Queries ClashZones table to find unique MEP categories
        /// Falls back to XML if database check fails
        /// </summary>
        private List<string> GetAllAvailableCategories(Document doc)
        {
            var categories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            try
            {
                // ✅ DATABASE MIGRATION: Query database first (primary source)
                var dbContext = new Data.SleeveDbContext(doc);
                var clashZoneRepo = new Data.Repositories.ClashZoneRepository(dbContext, null);
                
                // Query database for unique MEP categories
                using (var cmd = dbContext.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT DISTINCT MepCategory 
                        FROM ClashZones 
                        WHERE MepCategory IS NOT NULL AND MepCategory != ''
                        ORDER BY MepCategory";
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            var category = reader.GetString(0);
                            if (!string.IsNullOrEmpty(category))
                            {
                                categories.Add(category);
                            }
                        }
                    }
                }
                
                // ✅ LOG: Log categories found in database
                if (categories.Count > 0)
                {
                    DebugLogger.Info($"[MarkParameterCommand] ✅ Found {categories.Count} categories in database: {string.Join(", ", categories)}");
                }
            }
            catch (Exception dbEx)
            {
                DebugLogger.Warning($"[MarkParameterCommand] Database query failed, falling back to XML: {dbEx.Message}");
            }
            
            // ✅ FALLBACK: If database returned no categories, try XML files (for backward compatibility)
            if (categories.Count == 0)
            {
                try
                {
                    var filtersDirectory = ProjectPathService.GetFiltersDirectory(doc);
                    
                    if (Directory.Exists(filtersDirectory))
                    {
                        var xmlFiles = Directory.GetFiles(filtersDirectory, "*.xml");

                        foreach (var xmlFile in xmlFiles)
                        {
                            try
                            {
                                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                                using (var reader = new StreamReader(xmlFile))
                                {
                                    var filter = (OpeningFilter)serializer.Deserialize(reader);
                                    if (filter?.ClashZoneStorage?.AllZones != null)
                                    {
                                        foreach (var clashZone in filter.ClashZoneStorage.AllZones)
                                        {
                                            if (!string.IsNullOrEmpty(clashZone.MepElementCategory))
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
                        
                        if (categories.Count > 0)
                        {
                            DebugLogger.Info($"[MarkParameterCommand] Fallback: Found {categories.Count} categories from XML files: {string.Join(", ", categories)}");
                        }
                    }
                }
                catch (Exception xmlEx)
                {
                    DebugLogger.Error($"[MarkParameterCommand] Error getting available categories from XML fallback: {xmlEx.Message}");
                }
            }
            
            return categories.ToList();
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
