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
    /// Supports splitting Prefix and Numbering logic.
    /// </summary>
    public class MarkParameterCommand : ICommand
    {
        public enum MarkingMode
        {
            Full,
            PrefixOnly,
            NumberOnly
        }

        private readonly string _targetCategory;
        private readonly string _projectPrefix;
        private readonly string _disciplinePrefix;
        private readonly bool _remarkAll;
        private readonly MarkPrefixSettings _markPrefixes;
        private readonly MarkingMode _mode;
        
        public MarkParameterCommand(string targetCategory, string projectPrefix, string disciplinePrefix, bool remarkAll = false, MarkPrefixSettings? markPrefixes = null, MarkingMode mode = MarkingMode.Full)
        {
            _targetCategory = targetCategory ?? throw new ArgumentNullException(nameof(targetCategory));
            _projectPrefix = projectPrefix ?? throw new ArgumentNullException(nameof(projectPrefix));
            _disciplinePrefix = disciplinePrefix ?? throw new ArgumentNullException(nameof(disciplinePrefix));
            _remarkAll = remarkAll;
            _markPrefixes = markPrefixes;
            _mode = mode;
        }
        
        public void Execute(UIApplication app)
        {
            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] 🔥 MarkParameterCommand.Execute CALLED 🔥 Mode: {_mode}\n");
            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] Target Category: {_targetCategory}\n");
            
            try
            {
                var doc = app.ActiveUIDocument.Document;
                
                // ✅ HANDLE "ALL" or "SELECTED" CATEGORIES
                if (_targetCategory.Equals("ALL", StringComparison.OrdinalIgnoreCase) || 
                    _targetCategory.Equals("SELECTED", StringComparison.OrdinalIgnoreCase))
                {
                    var availableCategories = GetAllAvailableCategories(doc);
                    RemarkDebugLogger.LogInfo($"Available categories from DB/XML: {string.Join(", ", availableCategories)}");
                    var categoriesToProcess = new List<string>();
                    
                    // Filter categories based on selection/remark flags
                    foreach (var category in availableCategories)
                    {
                        var remarkFlag = _markPrefixes?.GetRemarkFlag(category) ?? _remarkAll;
                        RemarkDebugLogger.LogInfo($"Category '{category}' RemarkFlag: {remarkFlag}");
                        if (_targetCategory.Equals("SELECTED", StringComparison.OrdinalIgnoreCase) && !remarkFlag) 
                            continue;
                        categoriesToProcess.Add(category);
                    }
                    RemarkDebugLogger.LogStep($"Categories to process in phase 1 & 2: {string.Join(", ", categoriesToProcess)}");

                    using (var tx = new Transaction(doc, "Mark Categories"))
                    {
                        tx.Start();
                        
                        int totalProcessed = 0;
                        int totalErrors = 0;
                        // ✅ LOGGING FIX: Inject logger action
                        var markService = new MarkParameterService(doc, NumberingDebugLogger.LogInfo);
                        var numberFormat = _markPrefixes?.NumberFormat ?? "000";
                        var allowedPrefixes = new HashSet<string>();

                        // ---------------------------------------------------------
                        // PHASE 1: APPLY PREFIXES (Per Category, DB Driven)
                        // ---------------------------------------------------------
                        if (_mode != MarkingMode.NumberOnly)
                        {
                            foreach (var category in categoriesToProcess)
                            {
                                var disciplinePrefix = _markPrefixes?.GetDisciplinePrefix(category) ?? GetDisciplinePrefixForCategory(category);
                                var remarkFlag = _markPrefixes?.GetRemarkFlag(category) ?? _remarkAll;
                                
                                RemarkDebugLogger.LogStep($"Calling ApplyPrefixesOnly for '{category}' (Remark: {remarkFlag})");
                                var (p1, e1) = markService.ApplyPrefixesOnly(doc, category, _projectPrefix, disciplinePrefix, remarkFlag, _markPrefixes);
                                RemarkDebugLogger.LogInfo($"ApplyPrefixesOnly for '{category}' Result: {p1} updated, {e1} errors");
                                totalErrors += e1;
                                // Note: ApplyPrefixesOnly returns count of updates. 
                            }

                            // Handle Combined Sleeves Prefixes
                            var comboService = new CombinedSleeveMarkService();
                            var combinedSleeves = comboService.GetAllCombinedSleeves(doc);
                            if (combinedSleeves.Count > 0)
                            {
                                bool remarkCombined = _markPrefixes?.RemarkAll ?? _remarkAll;
                                var prefixAssignments = comboService.CalculateCombinedSleevePrefixes(doc, combinedSleeves, _projectPrefix, remarkCombined);
                                var (s, f) = comboService.ApplyCombinedSleeveMarksBatch(doc, prefixAssignments);
                                totalErrors += f;
                            }
                        }

                        // ---------------------------------------------------------
                        // PHASE 2: APPLY NUMBERS (Batch, Revit Scan)
                        // ---------------------------------------------------------
                        if (_mode != MarkingMode.PrefixOnly)
                        {
                            // ✅ USER REQUEST: REMOVE ALL PREFIX CONSTRAINTS
                            // "WE WANT ONLY PREFIX BASED"
                            // Passing null means "Process ALL prefixes found in the model"
                            var (p2, e2) = markService.ApplyNumbersBatch(doc, numberFormat, null, _markPrefixes);
                            totalProcessed += p2;
                            totalErrors += e2;
                        }
                        
                        tx.Commit();
                        DebugLogger.Info($"[MarkParameterCommand] ✅ ALL COMPLETE: {totalProcessed} numbered, {totalErrors} errors");
                    }
                }
                else  // ✅ SINGLE CATEGORY
                {
                    var remarkFlag = _markPrefixes?.GetRemarkFlag(_targetCategory) ?? _remarkAll;
                    
                    using (var tx = new Transaction(doc, $"Mark {_targetCategory}"))
                    {
                        tx.Start();

                        // ✅ LOGGING FIX: Inject logger action
                        var markService = new MarkParameterService(doc, NumberingDebugLogger.LogInfo);
                        var numberFormat = _markPrefixes?.NumberFormat ?? "000";
                        
                        // 1. PREFIX
                        if (_mode != MarkingMode.NumberOnly)
                        {
                            var disciplinePrefix = _markPrefixes?.GetDisciplinePrefix(_targetCategory) ?? GetDisciplinePrefixForCategory(_targetCategory);
                            markService.ApplyPrefixesOnly(doc, _targetCategory, _projectPrefix, disciplinePrefix, remarkFlag, _markPrefixes);
                            
                            // Combined? Usually not for single category click, but existing logic did it.
                            // We'll skip complex combined logic for single category click to stay focused? 
                            // Or just include it blindly. Existing code included it.
                            var comboService = new CombinedSleeveMarkService();
                            var combinedSleeves = comboService.GetAllCombinedSleeves(doc);
                            if (combinedSleeves.Count > 0)
                            {
                                 bool remarkCombined = _markPrefixes?.RemarkAll ?? _remarkAll;
                                 var prefixAssignments = comboService.CalculateCombinedSleevePrefixes(doc, combinedSleeves, _projectPrefix, remarkCombined);
                                 comboService.ApplyCombinedSleeveMarksBatch(doc, prefixAssignments);
                            }
                        }
                        
                        // 2. NUMBER
                         if (_mode != MarkingMode.PrefixOnly)
                        {
                            // ✅ USER REQUEST: REMOVE ALL PREFIX CONSTRAINTS
                            // "WE WANT ONLY PREFIX BASED" - Do not filter by category allowed list.
                            markService.ApplyNumbersBatch(doc, numberFormat, null, _markPrefixes);
                        }

                        tx.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterCommand] Error: {ex.Message}");
                throw;
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
