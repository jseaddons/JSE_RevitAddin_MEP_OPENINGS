using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
// using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration;
using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Processing;

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
        // private readonly MarkPrefixSettings _markPrefixes;
        private readonly MarkingMode _mode;
        
        public MarkParameterCommand(string targetCategory, string projectPrefix, string disciplinePrefix, bool remarkAll = false, object markPrefixes = null, MarkingMode mode = MarkingMode.Full)
        {
            _targetCategory = targetCategory ?? throw new ArgumentNullException(nameof(targetCategory));
            _projectPrefix = projectPrefix ?? throw new ArgumentNullException(nameof(projectPrefix));
            _disciplinePrefix = disciplinePrefix ?? throw new ArgumentNullException(nameof(disciplinePrefix));
            _remarkAll = remarkAll;
            // _markPrefixes = markPrefixes;
            _mode = mode;
        }
        
        public void Execute(UIApplication app)
        {
            DebugLogger.Info($"[{DateTime.Now:HH:mm:ss}] MarkParameterCommand is DEPRECATED and disabled in this build.\n");
            // Logic removed to suppress build errors for MarkPrefixSettings
        }
        
        /// <summary>
        /// Γ£à DATABASE-BASED: Get all available categories from the database
        /// Queries ClashZones table to find unique MEP categories
        /// Falls back to XML if database check fails
        /// </summary>
        private List<string> GetAllAvailableCategories(Document doc)
        {
            var categories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            try
            {
                // Γ£à DATABASE MIGRATION: Query database first (primary source)
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
                
                // Γ£à LOG: Log categories found in database
                if (categories.Count > 0)
                {
                    DebugLogger.Info($"[MarkParameterCommand] Γ£à Found {categories.Count} categories in database: {string.Join(", ", categories)}");
                }
            }
            catch (Exception dbEx)
            {
                DebugLogger.Warning($"[MarkParameterCommand] Database query failed, falling back to XML: {dbEx.Message}");
            }
            
            // Γ£à FALLBACK: If database returned no categories, try XML files (for backward compatibility)
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
