using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Represents a catalog of unique system and service types discovered from persisted clash data.
    /// </summary>
    public class SystemTypeCatalog
    {
        public HashSet<string> SystemTypes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> ServiceTypes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public bool IsEmpty => SystemTypes.Count == 0 && ServiceTypes.Count == 0;
    }

    /// <summary>
    /// Loads unique System Type / Service Type values from stored filter XML (and future persistence sources).
    /// Used to drive UI dropdowns for System Type overrides.
    /// </summary>
    public class SystemTypeCatalogService
    {
        private readonly Action<string> _logger;

        public SystemTypeCatalogService(Action<string> logger = null)
        {
            _logger = logger ?? (_ => { });
        }

        public SystemTypeCatalog Load(Document document, string category = null)
        {
            var catalog = new SystemTypeCatalog();

            if (document == null)
            {
                _logger("[SystemTypeCatalog] No document context supplied – returning empty catalog.");
                return catalog;
            }

            try
            {
                // ✅ STEP 1: Try loading directly from Revit first (primary method)
                // This is simpler, always up-to-date, and doesn't depend on database/JSON parsing
                LoadSystemTypesFromRevit(document, catalog, category);
                
                // ✅ STEP 2: If Revit method returned no results, fallback to database
                if (catalog.SystemTypes.Count == 0 && catalog.ServiceTypes.Count == 0)
                {
                    _logger($"[SystemTypeCatalog] ⚠️ Revit query returned no results - falling back to database");
                    LoadSystemTypesFromDatabase(document, catalog, category);
                }
            }
            catch (Exception ex)
            {
                _logger($"[SystemTypeCatalog] Unexpected error in primary method: {ex.Message}");
                // ✅ FALLBACK: Try database if Revit method fails
                try
                {
                    _logger($"[SystemTypeCatalog] Attempting database fallback...");
                    LoadSystemTypesFromDatabase(document, catalog, category);
                }
                catch (Exception dbEx)
                {
                    _logger($"[SystemTypeCatalog] Database fallback also failed: {dbEx.Message}");
                    _logger($"[SystemTypeCatalog] Stack trace: {ex.StackTrace}");
                }
            }

            return catalog;
        }

        /// <summary>
        /// ✅ SIMPLIFIED: Loads system/service types directly from Revit by querying placed sleeves
        /// Gets unique System Type (for Ducts/Pipes) or Service Type (for Cable Trays) from MEP elements
        /// ✅ NEW: Filters by category if specified - only returns system/service types for that category
        /// </summary>
        private void LoadSystemTypesFromRevit(Document document, SystemTypeCatalog catalog, string category = null)
        {
            try
            {
                // ✅ CRITICAL FIX: Ensure document is valid
                if (document == null || document.IsValidObject == false)
                {
                    _logger($"[SystemTypeCatalog] ⚠️ Invalid document - cannot load system types from Revit");
                    return;
                }
                
                // ✅ STEP 1: Get individual sleeve IDs from database (SleeveInstanceId > 0, NOT ClusterInstanceId)
                var individualSleeveIds = new HashSet<int>();
                try
                {
                    using (var context = new Data.SleeveDbContext(document))
                    {
                        using (var cmd = context.Connection.CreateCommand())
                        {
                            // ✅ CRITICAL: Only get individual sleeves (SleeveInstanceId > 0), exclude cluster sleeves
                            cmd.CommandText = @"
                                SELECT DISTINCT SleeveInstanceId 
                                FROM ClashZones 
                                WHERE SleeveInstanceId > 0
                                UNION
                                SELECT DISTINCT SleeveInstanceId 
                                FROM SleeveSnapshots 
                                WHERE SleeveInstanceId > 0";
                            
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var sleeveId = Convert.ToInt32(reader["SleeveInstanceId"] ?? 0);
                                    if (sleeveId > 0)
                                    {
                                        individualSleeveIds.Add(sleeveId);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception dbEx)
                {
                    _logger($"[SystemTypeCatalog] ⚠️ Error querying database for individual sleeves: {dbEx.Message}");
                    // Fallback: process all sleeves if database query fails
                }
                
                // ✅ STEP 2: Collect only individual sleeves (those with IDs in our set)
                var allSleeves = new FilteredElementCollector(document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi =>
                    {
                        var famName = fi.Symbol?.Family?.Name ?? string.Empty;
                        // Match only the 4 specific opening families
                        bool isOpeningFamily = famName.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0
                            || famName.IndexOf("OpeningOnSlab", StringComparison.OrdinalIgnoreCase) >= 0;
                        
                        if (!isOpeningFamily) return false;
                        
                        // ✅ CRITICAL: Only include if this sleeve ID is in our individual sleeves set
                        // If database query failed, include all (fallback)
                        if (individualSleeveIds.Count == 0) return true; // Fallback: include all
                        
                        return individualSleeveIds.Contains(fi.Id.GetIntegerValue());
                    })
                    .ToList();
                
                _logger($"[SystemTypeCatalog] Found {allSleeves.Count} individual sleeves (SleeveInstanceId > 0) in document");
                
                // ✅ STEP 3: Use HashSet for proper deduplication
                var uniqueSystemTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var uniqueServiceTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                
                int processedSleeves = 0;
                int skippedNoMepId = 0;
                int skippedMepNotFound = 0;
                
                // ✅ STEP 4: For each individual sleeve, get its MEP element and extract system/service type
                foreach (var sleeve in allSleeves)
                {
                    try
                    {
                        // Get MEP_ElementId parameter from sleeve
                        var mepIdParam = sleeve.LookupParameter("MEP_ElementId");
                        if (mepIdParam == null || !mepIdParam.HasValue)
                        {
                            skippedNoMepId++;
                            continue;
                        }
                        
                        int mepElementId = mepIdParam.AsInteger();
                        if (mepElementId <= 0)
                        {
                            skippedNoMepId++;
                            continue;
                        }
                        
                        // Get MEP element from document (could be in linked file)
                        Element mepElement = null;
                        
                        // Try active document first
#if REVIT2023
                        mepElement = document.GetElement(new ElementId(mepElementId));
#else
                        mepElement = document.GetElement(new ElementId((long)mepElementId));
#endif
                        
                        // If not found, try linked files
                        if (mepElement == null)
                        {
                            var linkInstances = new FilteredElementCollector(document)
                                .OfClass(typeof(RevitLinkInstance))
                                .Cast<RevitLinkInstance>();
                            
                            foreach (var linkInstance in linkInstances)
                            {
                                var linkDoc = linkInstance.GetLinkDocument();
                                if (linkDoc != null)
                                {
#if REVIT2023
                                    mepElement = linkDoc.GetElement(new ElementId(mepElementId));
#else
                                    mepElement = linkDoc.GetElement(new ElementId((long)mepElementId));
#endif
                                    if (mepElement != null) break;
                                }
                            }
                        }
                        
                        if (mepElement == null)
                        {
                            skippedMepNotFound++;
                            continue;
                        }
                        
                        // ✅ STEP 3: Determine category from MEP element
                        var mepCategory = mepElement.Category?.Name ?? string.Empty;
                        
                        // ✅ STEP 4: Filter by category if specified
                        if (!string.IsNullOrWhiteSpace(category))
                        {
                            // Map category names
                            var categoryMap = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
                            {
                                { "Ducts", new[] { "Ducts", "Duct Fittings", "Duct Accessories" } },
                                { "Pipes", new[] { "Pipes", "Pipe Fittings", "Plumbing Fixtures" } },
                                { "Cable Trays", new[] { "Cable Trays", "Conduits" } },
                                { "Duct Accessories", new[] { "Duct Accessories" } }
                            };
                            
                            if (categoryMap.ContainsKey(category))
                            {
                                var allowedCategories = categoryMap[category];
                                if (!allowedCategories.Any(c => mepCategory.Equals(c, StringComparison.OrdinalIgnoreCase)))
                                {
                                    continue; // Skip this sleeve - category doesn't match
                                }
                            }
                        }
                        
                        // ✅ STEP 5: Extract System Type or Service Type based on category
                        // ✅ CRITICAL: Use HashSet for deduplication - only add unique values
                        if (mepCategory.IndexOf("Cable Tray", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            mepCategory.IndexOf("Conduit", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // For Cable Trays/Conduits: Get Service Type
                            var serviceType = GetSystemTypeFromElement(mepElement, isServiceType: true);
                            if (!string.IsNullOrWhiteSpace(serviceType))
                            {
                                var trimmed = serviceType.Trim();
                                if (uniqueServiceTypes.Add(trimmed)) // Add returns true if new item added
                                {
                                    catalog.ServiceTypes.Add(trimmed);
                                }
                            }
                        }
                        else if (mepCategory.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 mepCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // For Ducts/Pipes: Get System Type
                            var systemType = GetSystemTypeFromElement(mepElement, isServiceType: false);
                            if (!string.IsNullOrWhiteSpace(systemType))
                            {
                                var trimmed = systemType.Trim();
                                if (uniqueSystemTypes.Add(trimmed)) // Add returns true if new item added
                                {
                                    catalog.SystemTypes.Add(trimmed);
                                }
                            }
                        }
                        else
                        {
                            // Unknown category - try both
                            var systemType = GetSystemTypeFromElement(mepElement, isServiceType: false);
                            if (!string.IsNullOrWhiteSpace(systemType))
                            {
                                var trimmed = systemType.Trim();
                                if (uniqueSystemTypes.Add(trimmed))
                                {
                                    catalog.SystemTypes.Add(trimmed);
                                }
                            }
                            
                            var serviceType = GetSystemTypeFromElement(mepElement, isServiceType: true);
                            if (!string.IsNullOrWhiteSpace(serviceType))
                            {
                                var trimmed = serviceType.Trim();
                                if (uniqueServiceTypes.Add(trimmed))
                                {
                                    catalog.ServiceTypes.Add(trimmed);
                                }
                            }
                        }
                        
                        processedSleeves++;
                    }
                    catch (Exception ex)
                    {
                        _logger($"[SystemTypeCatalog] Error processing sleeve {sleeve.Id}: {ex.Message}");
                    }
                }
                
                // ✅ DIAGNOSTIC: Log statistics (show unique counts, not total extracted)
                _logger($"[SystemTypeCatalog] Revit query stats: {allSleeves.Count} individual sleeves found, {processedSleeves} processed, {skippedNoMepId} skipped (no MEP ID), {skippedMepNotFound} skipped (MEP not found), {uniqueSystemTypes.Count} UNIQUE System Types, {uniqueServiceTypes.Count} UNIQUE Service Types");
            }
            catch (Exception ex)
            {
                _logger($"[SystemTypeCatalog] Error loading from Revit: {ex.Message}");
                if (ex.InnerException != null)
                {
                    _logger($"[SystemTypeCatalog] Inner exception: {ex.InnerException.Message}");
                }
            }
        }
        
        /// <summary>
        /// ✅ FALLBACK: Loads system types from SQLite database (fallback method)
        /// Uses "dumb once, use many times" principle - data was collected during refresh and stored in DB
        /// ✅ NEW: Filters by category if specified - only returns system/service types for that category
        /// </summary>
        private void LoadSystemTypesFromDatabase(Document document, SystemTypeCatalog catalog, string category = null)
        {
            try
            {
                // ✅ CRITICAL FIX: Ensure document is valid before accessing database
                if (document == null || document.IsValidObject == false)
                {
                    _logger($"[SystemTypeCatalog] ⚠️ Invalid document - cannot load system types from database");
                    return;
                }
                
                using (var context = new Data.SleeveDbContext(document, msg =>
                {
                    _logger($"[SystemTypeCatalog][SQLite] {msg}");
                }))
                {
                    // ✅ CRITICAL FIX: Ensure connection is valid
                    if (context.Connection == null)
                    {
                        _logger($"[SystemTypeCatalog] ⚠️ Database connection is null - cannot load system types");
                        return;
                    }
                    
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        // ✅ CRITICAL FIX: Only query individual sleeves (SleeveInstanceId > 0), NOT cluster sleeves
                        // Query SleeveSnapshots for individual sleeves only (SleeveInstanceId > 0, ClusterInstanceId = 0 or NULL)
                        if (!string.IsNullOrWhiteSpace(category))
                        {
                            // ✅ STEP 1: Get only SleeveInstanceIds (NOT ClusterInstanceIds) for this category
                            var sleeveIds = new List<int>();
                            
                            using (var categoryCmd = context.Connection.CreateCommand())
                            {
                                categoryCmd.CommandText = @"
                                    SELECT DISTINCT SleeveInstanceId
                                    FROM ClashZones 
                                    WHERE MepCategory = @Category
                                      AND SleeveInstanceId > 0";
                                categoryCmd.Parameters.AddWithValue("@Category", category);
                                
                                using (var categoryReader = categoryCmd.ExecuteReader())
                                {
                                    while (categoryReader.Read())
                                    {
                                        var sleeveId = Convert.ToInt32(categoryReader["SleeveInstanceId"] ?? 0);
                                        if (sleeveId > 0) sleeveIds.Add(sleeveId);
                                    }
                                }
                            }
                            
                            // ✅ STEP 2: Query SleeveSnapshots using only individual sleeve IDs
                            if (sleeveIds.Count > 0)
                            {
                                var placeholders = string.Join(",", sleeveIds.Select((_, i) => $"@SleeveId{i}"));
                                cmd.CommandText = $@"
                                    SELECT DISTINCT MepParametersJson 
                                    FROM SleeveSnapshots 
                                    WHERE SleeveInstanceId IN ({placeholders})
                                      AND MepParametersJson IS NOT NULL 
                                      AND MepParametersJson != '{{}}'
                                      AND MepParametersJson != ''";
                                
                                for (int i = 0; i < sleeveIds.Count; i++)
                                {
                                    cmd.Parameters.AddWithValue($"@SleeveId{i}", sleeveIds[i]);
                                }
                            }
                            else
                            {
                                // No individual sleeves found for this category - return empty result
                                cmd.CommandText = "SELECT NULL WHERE 1=0";
                            }
                        }
                        else
                        {
                            // No category filter - get all individual sleeves only (SleeveInstanceId > 0, NOT ClusterInstanceId)
                            cmd.CommandText = @"
                                SELECT DISTINCT MepParametersJson 
                                FROM SleeveSnapshots 
                                WHERE SleeveInstanceId > 0
                                  AND MepParametersJson IS NOT NULL 
                                  AND MepParametersJson != '{}'
                                  AND MepParametersJson != ''";
                        }

                        // ✅ CRITICAL: Use HashSet for proper deduplication
                        var uniqueSystemTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        var uniqueServiceTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        
                        int rowCount = 0;
                        int jsonParseSuccess = 0;
                        int jsonParseFailed = 0;
                        
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                rowCount++;
                                try
                                {
                                    var jsonText = reader["MepParametersJson"]?.ToString();
                                    if (string.IsNullOrWhiteSpace(jsonText))
                                    {
                                        jsonParseFailed++;
                                        continue;
                                    }

                                    // Parse JSON dictionary
                                    var mepParams = DeserializeDictionary(jsonText);
                                    if (mepParams == null || mepParams.Count == 0)
                                    {
                                        jsonParseFailed++;
                                        continue;
                                    }
                                    
                                    jsonParseSuccess++;
                                    
                                    // ✅ DIAGNOSTIC: Log sample parameter keys found (first row only)
                                    if (rowCount == 1 && !DeploymentConfiguration.DeploymentMode)
                                    {
                                        var sampleKeys = mepParams.Keys.Take(10).ToList();
                                        _logger($"[SystemTypeCatalog] Sample parameter keys in first row: {string.Join(", ", sampleKeys)}");
                                    }

                                    // ✅ CATEGORY-SPECIFIC: Extract System Type for Ducts/Pipes, Service Type for Cable Trays
                                    // ✅ CRITICAL: Use HashSet.Add() to ensure only unique values are added
                                    if (category == "Cable Trays")
                                    {
                                        // For Cable Trays, extract Service Type (try multiple parameter name variations)
                                        var serviceType = GetParameterValue(mepParams, 
                                            "Service Type", 
                                            "MEP Service Type", 
                                            "SERVICE TYPE",
                                            "service type",
                                            "ServiceType",
                                            "MEPServiceType");
                                        
                                        if (!string.IsNullOrWhiteSpace(serviceType))
                                        {
                                            var trimmed = serviceType.Trim();
                                            if (uniqueServiceTypes.Add(trimmed))
                                            {
                                                catalog.ServiceTypes.Add(trimmed);
                                            }
                                        }
                                    }
                                    else if (category == "Ducts" || category == "Pipes" || category == "Duct Accessories")
                                    {
                                        // For Ducts/Pipes, extract System Type
                                        var systemType = GetParameterValue(mepParams, "System Type", "MEP System Type", "System Classification");
                                        if (!string.IsNullOrWhiteSpace(systemType))
                                        {
                                            var trimmed = systemType.Trim();
                                            if (uniqueSystemTypes.Add(trimmed))
                                            {
                                                catalog.SystemTypes.Add(trimmed);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // No category filter - extract both System Type and Service Type
                                        var systemType = GetParameterValue(mepParams, "System Type", "MEP System Type", "System Classification");
                                        if (!string.IsNullOrWhiteSpace(systemType))
                                        {
                                            var trimmed = systemType.Trim();
                                            if (uniqueSystemTypes.Add(trimmed))
                                            {
                                                catalog.SystemTypes.Add(trimmed);
                                            }
                                        }

                                        var serviceType = GetParameterValue(mepParams, "Service Type", "MEP Service Type");
                                        if (!string.IsNullOrWhiteSpace(serviceType))
                                        {
                                            var trimmed = serviceType.Trim();
                                            if (uniqueServiceTypes.Add(trimmed))
                                            {
                                                catalog.ServiceTypes.Add(trimmed);
                                            }
                                        }
                                    }
                                }
                                catch (Exception jsonEx)
                                {
                                    jsonParseFailed++;
                                    _logger($"[SystemTypeCatalog] Error parsing MepParametersJson (row {rowCount}): {jsonEx.Message}");
                                }
                            }
                        }
                        
                        // ✅ DIAGNOSTIC: Log query statistics (show unique counts)
                        _logger($"[SystemTypeCatalog] Database query stats: {rowCount} rows read, {jsonParseSuccess} JSON parsed successfully, {jsonParseFailed} failed, {uniqueSystemTypes.Count} UNIQUE System Types, {uniqueServiceTypes.Count} UNIQUE Service Types");
                    }
                }

                
                var categoryInfo = string.IsNullOrWhiteSpace(category) ? "all categories" : $"category '{category}'";
                _logger($"[SystemTypeCatalog] ✅ Loaded {catalog.SystemTypes.Count} unique system types and {catalog.ServiceTypes.Count} unique service types from database ({categoryInfo})");
                
                // ✅ ENHANCED LOGGING: Log ALL system types and service types found
                if (catalog.SystemTypes.Count > 0)
                {
                    var allSystemTypes = catalog.SystemTypes.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToList();
                    _logger($"[SystemTypeCatalog] ✅ ALL System Types found ({allSystemTypes.Count}): {string.Join(", ", allSystemTypes)}");
                }
                else
                {
                    _logger($"[SystemTypeCatalog] ⚠️ No System Types found for {categoryInfo}");
                }
                
                if (catalog.ServiceTypes.Count > 0)
                {
                    var allServiceTypes = catalog.ServiceTypes.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).ToList();
                    _logger($"[SystemTypeCatalog] ✅ ALL Service Types found ({allServiceTypes.Count}): {string.Join(", ", allServiceTypes)}");
                }
                else
                {
                    _logger($"[SystemTypeCatalog] ⚠️ No Service Types found for {categoryInfo}");
                }
            }
            catch (Exception ex)
            {
                _logger($"[SystemTypeCatalog] Error loading from Revit: {ex.Message}");
                if (ex.InnerException != null)
                {
                    _logger($"[SystemTypeCatalog] Inner exception: {ex.InnerException.Message}");
                }
            }
        }
        
        /// <summary>
        /// Get System Type or Service Type parameter value from a MEP element
        /// </summary>
        private static string GetSystemTypeFromElement(Element element, bool isServiceType = false)
        {
            if (element == null) return null;
            
            try
            {
                if (isServiceType)
                {
                    // For Cable Trays: Get Service Type
                    var param = element.LookupParameter("Service Type") 
                             ?? element.LookupParameter("MEP Service Type");
                    
                    if (param != null && param.HasValue)
                    {
                        return param.AsString() ?? param.AsValueString();
                    }
                }
                else
                {
                    // For Ducts/Pipes: Get System Type
                    // Try built-in parameters first (faster and more reliable)
                    var builtInParam = element.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
                                     ?? element.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)
                                     ?? element.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
                    
                    if (builtInParam != null && builtInParam.HasValue)
                    {
                        // System Type is stored as ElementId - resolve to element name
                        if (builtInParam.StorageType == StorageType.ElementId)
                        {
                            var systemTypeId = builtInParam.AsElementId();
                            if (systemTypeId != null && systemTypeId != ElementId.InvalidElementId)
                            {
                                var systemTypeElement = element.Document.GetElement(systemTypeId);
                                if (systemTypeElement != null)
                                {
                                    return systemTypeElement.Name ?? systemTypeElement.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)?.AsString();
                                }
                            }
                        }
                        else
                        {
                            return builtInParam.AsString() ?? builtInParam.AsValueString();
                        }
                    }
                    
                    // Fallback to instance parameter
                    var instanceParam = element.LookupParameter("System Type")
                                     ?? element.LookupParameter("MEP System Type")
                                     ?? element.LookupParameter("System Classification");
                    
                    if (instanceParam != null && instanceParam.HasValue)
                    {
                        return instanceParam.AsString() ?? instanceParam.AsValueString();
                    }
                }
            }
            catch
            {
                // Ignore errors - return null
            }
            
            return null;
        }

        /// <summary>
        /// Helper to get parameter value from dictionary with multiple key options
        /// ✅ CRITICAL: Uses case-insensitive matching to handle parameter name variations
        /// </summary>
        private static string GetParameterValue(Dictionary<string, string> dict, params string[] keys)
        {
            if (dict == null || keys == null || dict.Count == 0)
                return null;

            // First try exact match (fast path)
            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                if (dict.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            // If exact match fails, try case-insensitive match
            // This handles variations like "Service Type" vs "SERVICE TYPE" vs "service type"
            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                var match = dict.FirstOrDefault(kv => 
                    kv.Key != null && 
                    kv.Key.Equals(key, StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(kv.Value));
                
                if (match.Key != null)
                {
                    return match.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// Deserialize JSON dictionary using System.Text.Json (same as ClashZoneRepository)
        /// </summary>
        private static Dictionary<string, string> DeserializeDictionary(string json)
        {
            if (string.IsNullOrWhiteSpace(json) || json == "{}")
                return new Dictionary<string, string>();

            try
            {
                return JsonSerializer.Deserialize<Dictionary<string, string>>(json) 
                    ?? new Dictionary<string, string>();
            }
            catch
            {
                return new Dictionary<string, string>();
            }
        }

        private static string ExtractParameterValue(ClashZone zone, params string[] keys)
        {
            if (zone?.MepParameterValues == null || zone.MepParameterValues.Count == 0 || keys == null)
                return null;

            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                var match = zone.MepParameterValues
                    .FirstOrDefault(kv => kv != null &&
                                           kv.Key != null &&
                                           kv.Key.Equals(key, StringComparison.OrdinalIgnoreCase));

                if (match != null && !string.IsNullOrWhiteSpace(match.Value))
                {
                    return match.Value;
                }
            }

            return null;
        }
    }
}

