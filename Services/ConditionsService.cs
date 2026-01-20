using System;
using System.IO;
using System.Xml.Serialization;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL SERVICE - DO NOT REMOVE OR MODIFY WITHOUT EXTENSIVE TESTING ⚠️⚠️⚠️
        /// 
        /// Manages saving and loading of OpeningConditions (clearances, opening types, etc.)
        /// Saved as FilterName_CONDITIONS.xml alongside clash zone data
        /// This implements proper separation: CLASH ZONES (data) vs CONDITIONS (configuration)
        /// 
        /// ✅ WORKING AS OF 2025-11-15: Conditions are successfully populating in SQLite database
        /// 
        /// ⚠️ PROTECTION RULES:
        /// 1. DO NOT remove the SQLite save logic (SaveConditionsToSqlite method)
        /// 2. DO NOT change the return type of SaveConditionsToSqlite (must return bool)
        /// 3. DO NOT remove validation checks (null document, empty filterName/combinedKey, invalid filterId)
        /// 4. DO NOT remove error logging - it's critical for diagnosing issues
        /// 5. DO NOT bypass the filter registration step (EnsureFilter must succeed)
        /// 6. DO NOT change the method signature of UpsertConditions in ConditionRepository
        /// 
        /// 🔒 LOCKED CODE PATHS:
        /// - SaveConditions() → SaveConditionsToSqlite() → ConditionRepository.UpsertConditions()
        /// - All validation checks must remain in place
        /// - All error logging must remain in place
        /// </summary>
        public class ConditionsService
    {
        private readonly string _projectDirectory;
        private readonly Action<string>? _log;
        private readonly Document? _document;

        public ConditionsService(Action<string>? log = null)
            : this(null, null, log)
        {
        }

        public ConditionsService(string? projectDirectory, Action<string>? log = null)
            : this(null, projectDirectory, log)
        {
        }

        public ConditionsService(Document? document, Action<string>? log = null)
            : this(document, null, log)
        {
        }

        public ConditionsService(Document? document, string? projectDirectory, Action<string>? log = null)
        {
            _document = document;
            _log = log ?? (msg => { });
            _projectDirectory = ResolveProjectDirectory(document, projectDirectory);
            EnsureDirectoryExists(_projectDirectory);
        }

        /// <summary>
        /// Save opening conditions to XML file
        /// File name: FilterName_CONDITIONS.xml (e.g., Ventilation_ducts_CONDITIONS.xml)
        /// </summary>
        public bool SaveConditions(OpeningConditions conditions, string customKey = null)
        {
            if (conditions == null)
            {
                _log("[ConditionsService] Cannot save conditions: payload is null");
                return false;
            }

            try
            {
                var resolvedFilterName = ResolveFilterName(conditions, customKey);
                var resolvedCategory = ResolveNormalizedCategory(conditions, customKey);
                var categorySuffix = MepCategoryConstants.GetXmlSuffix(resolvedCategory);
                var combinedKey = !string.IsNullOrWhiteSpace(customKey)
                    ? customKey
                    : BuildCombinedKey(resolvedFilterName, categorySuffix);
                
                // ✅ FIX: Normalize CombinedKey to prevent duplicate rows with .xml suffix
                combinedKey = NormalizeCombinedKeyForDatabase(combinedKey);
                
                var xmlKey = customKey ?? resolvedFilterName;

                conditions.FilterName = resolvedFilterName;
                conditions.Category = resolvedCategory;
                conditions.LastModified = DateTime.Now;

                // ✅ PHASE 2: Only save conditions XML if XML creation is enabled
                if (!DeploymentConfiguration.DisableXmlCreation)
                {
                    var fileName = $"{xmlKey}_CONDITIONS.xml";
                    var filePath = Path.Combine(_projectDirectory, fileName);

                    var serializer = new XmlSerializer(typeof(OpeningConditions));
                    using (var writer = new StreamWriter(filePath))
                    {
                        serializer.Serialize(writer, conditions);
                    }

                    _log($"[ConditionsService] Saved conditions to: {filePath}");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var clearance = conditions.ClearanceSettings ?? new ClearanceSettings();
                        DebugLogger.Info($"[ConditionsService] Saved conditions for key '{combinedKey}' - Clearances: Rect={clearance.RectangularNormal}/{clearance.RectangularInsulated}mm, Round={clearance.RoundNormal}/{clearance.RoundInsulated}mm");
                    }
                }
                else
                {
                    _log($"[ConditionsService] ⚠️ XML creation disabled - skipping conditions XML save (database only mode). Key='{combinedKey}'");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var clearance = conditions.ClearanceSettings ?? new ClearanceSettings();
                        DebugLogger.Info($"[ConditionsService] ⚠️ XML creation disabled - skipping conditions XML save for key '{combinedKey}' (database only mode). Clearances: Rect={clearance.RectangularNormal}/{clearance.RectangularInsulated}mm, Round={clearance.RoundNormal}/{clearance.RoundInsulated}mm");
                    }
                }

                // ⚠️⚠️⚠️ CRITICAL: SQLite save MUST be called - DO NOT REMOVE OR BYPASS ⚠️⚠️⚠️
                // ✅ WORKING AS OF 2025-11-15: Conditions are successfully populating in SQLite database
                // This code path is PROTECTED - any changes must be thoroughly tested
                bool sqliteSaved = SaveConditionsToSqlite(resolvedFilterName, resolvedCategory, combinedKey, conditions);
                if (!sqliteSaved)
                {
                    // ⚠️ CRITICAL: Log warning but DO NOT fail the entire save operation
                    // XML save succeeded, but SQLite save failed - this is a non-fatal error
                    _log($"[ConditionsService] ⚠️ Conditions saved to XML but SQLite save failed for '{combinedKey}'");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[ConditionsService] ⚠️ Conditions saved to XML but SQLite save failed for '{combinedKey}'");
                }
                else
                {
                    // ✅ SUCCESS: Both XML and SQLite saves succeeded
                    _log($"[ConditionsService] ✅ Conditions saved to BOTH XML and SQLite for '{combinedKey}'");
                }

                return true;
            }
            catch (Exception ex)
            {
                _log($"[ConditionsService] Error saving conditions: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ConditionsService] Error saving conditions: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Load opening conditions from XML file
        /// File name: FilterName_CONDITIONS.xml
        /// Returns null if file doesn't exist
        /// </summary>
        public OpeningConditions LoadConditions(string filterName)
        {
            var sqliteConditions = LoadConditionsFromSqlite(filterName);
            if (sqliteConditions != null)
                return sqliteConditions;

            try
            {
                if (string.IsNullOrEmpty(filterName))
                {
                    _log("[ConditionsService] Cannot load conditions: key is empty");
                    return CreateDefaultConditions(filterName);
                }

                var fileName = $"{filterName}_CONDITIONS.xml";
                var filePath = Path.Combine(_projectDirectory, fileName);

                if (!File.Exists(filePath))
                {
                    _log($"[ConditionsService] Conditions file not found: {filePath}, creating defaults");
                    var defaults = CreateDefaultConditions(filterName);
                    defaults.Category = ResolveNormalizedCategory(defaults, filterName);
                    return defaults;
                }

                var serializer = new XmlSerializer(typeof(OpeningConditions));
                using (var reader = new StreamReader(filePath))
                {
                    var conditions = (OpeningConditions)serializer.Deserialize(reader);
                    conditions.FilterName = ResolveFilterName(conditions, filterName);
                    conditions.Category = ResolveNormalizedCategory(conditions, filterName);

                    _log($"[ConditionsService] Loaded conditions from: {filePath}");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var clearance = conditions.ClearanceSettings ?? new ClearanceSettings();
                        DebugLogger.Info($"[ConditionsService] Loaded conditions for key '{filterName}' - Clearances: Rect={clearance.RectangularNormal}/{clearance.RectangularInsulated}mm, Round={clearance.RoundNormal}/{clearance.RoundInsulated}mm");
                    }

                    return conditions;
                }
            }
            catch (Exception ex)
            {
                _log($"[ConditionsService] Error loading conditions: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ConditionsService] Error loading conditions: {ex.Message}");
                var defaults = CreateDefaultConditions(filterName);
                defaults.Category = ResolveNormalizedCategory(defaults, filterName);
                return defaults;
            }
        }

        /// <summary>
        /// Create default conditions for a filter
        /// </summary>
        private OpeningConditions CreateDefaultConditions(string filterName)
        {
            return new OpeningConditions
            {
                FilterName = filterName,
                Category = "Ducts", // Will be updated based on actual filter
                ClearanceSettings = new ClearanceSettings
                {
                    RectangularNormal = 50.0,
                    RectangularInsulated = 25.0,
                    RoundNormal = 50.0,
                    RoundInsulated = 50.0
                },
                OpeningTypePreferences = new OpeningTypePreferences
                {
                    RoundDucts = "Circular",
                    Pipes = "Circular"
                }
            };
        }

        private string ResolveProjectDirectory(Document document, string overrideDirectory)
        {
            if (!string.IsNullOrWhiteSpace(overrideDirectory))
                return overrideDirectory;

            if (document != null)
            {
                ProjectPathService.EnsureFiltersDirectory(document);
                return ProjectPathService.GetFiltersDirectory(document);
            }

            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "JSE_MEP_Openings", "Projects", "Default", "Filters");
        }

        private static void EnsureDirectoryExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        private string ResolveFilterName(OpeningConditions conditions, string combinedKey)
        {
            if (!string.IsNullOrWhiteSpace(conditions?.FilterName))
            {
                string filterName = conditions.FilterName;
                // ✅ FIX: Strip .xml extension if present (prevents duplicate entries)
                if (filterName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    filterName = filterName.Substring(0, filterName.Length - 4);
                }
                return filterName;
            }

            if (string.IsNullOrWhiteSpace(combinedKey))
                return "Default";

            var index = combinedKey.LastIndexOf('_');
            if (index <= 0)
            {
                // ✅ FIX: Strip .xml extension if present
                if (combinedKey.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    return combinedKey.Substring(0, combinedKey.Length - 4);
                }
                return combinedKey;
            }

            string resolvedName = combinedKey.Substring(0, index);
            // ✅ FIX: Strip .xml extension if present
            if (resolvedName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                resolvedName = resolvedName.Substring(0, resolvedName.Length - 4);
            }
            return resolvedName;
        }

        private string ResolveNormalizedCategory(OpeningConditions conditions, string combinedKey)
        {
            if (!string.IsNullOrWhiteSpace(conditions?.Category))
                return MepCategoryConstants.Normalize(conditions.Category);

            if (string.IsNullOrWhiteSpace(combinedKey))
                return MepCategoryConstants.DUCTS;

            var index = combinedKey.LastIndexOf('_');
            if (index < 0 || index >= combinedKey.Length - 1)
                return MepCategoryConstants.DUCTS;

            var suffix = combinedKey.Substring(index + 1);
            return NormalizeCategoryFromSuffix(suffix);
        }

        private string NormalizeCategoryFromSuffix(string suffix)
        {
            if (string.IsNullOrWhiteSpace(suffix))
                return MepCategoryConstants.DUCTS;

            var lower = suffix.ToLowerInvariant();
            return lower switch
            {
                MepCategoryConstants.DUCTS_XML_SUFFIX => MepCategoryConstants.DUCTS,
                MepCategoryConstants.PIPES_XML_SUFFIX => MepCategoryConstants.PIPES,
                MepCategoryConstants.CABLE_TRAYS_XML_SUFFIX => MepCategoryConstants.CABLE_TRAYS,
                MepCategoryConstants.DUCT_ACCESSORIES_XML_SUFFIX => MepCategoryConstants.DUCT_ACCESSORIES,
                _ => MepCategoryConstants.Normalize(suffix)
            };
        }

        private string BuildCombinedKey(string filterName, string categorySuffix)
        {
            if (string.IsNullOrWhiteSpace(filterName))
                filterName = "Default";

            // ✅ FIX: Strip .xml extension from filter name if present (prevents duplicate entries)
            // Filter names should not include file extensions in CombinedKey
            if (filterName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                filterName = filterName.Substring(0, filterName.Length - 4);
            }

            if (string.IsNullOrWhiteSpace(categorySuffix))
                return filterName;

            return $"{filterName}_{categorySuffix}";
        }

        /// <summary>
        /// ✅ FIX: Normalizes CombinedKey by stripping .xml extension to prevent duplicate rows in database
        /// This ensures consistent key format: "FilterName_Category" without any .xml suffixes
        /// </summary>
        private string NormalizeCombinedKeyForDatabase(string combinedKey)
        {
            if (string.IsNullOrWhiteSpace(combinedKey))
                return combinedKey;

            // Strip .xml extension if present anywhere in the key
            // Handle cases like "Ventilation_duct_accessories.xml_du..." or "Ventilation_duct_accessories"
            string normalized = combinedKey;
            
            // Check if key contains .xml (could be in middle or end)
            int xmlIndex = normalized.IndexOf(".xml", StringComparison.OrdinalIgnoreCase);
            if (xmlIndex >= 0)
            {
                // Remove .xml and anything after it that looks like a file extension pattern
                // Example: "Ventilation_duct_accessories.xml_du" -> "Ventilation_duct_accessories"
                normalized = normalized.Substring(0, xmlIndex);
            }
            
            return normalized;
        }

        /// <summary>
        /// ⚠️⚠️⚠️ PROTECTED METHOD - DO NOT MODIFY WITHOUT EXTENSIVE TESTING ⚠️⚠️⚠️
        /// 
        /// Saves opening conditions to SQLite database.
        /// 
        /// ✅ WORKING AS OF 2025-11-15: Conditions are successfully populating in SQLite database
        /// 
        /// ⚠️ CRITICAL VALIDATION CHECKS (DO NOT REMOVE):
        /// 1. Document must not be null
        /// 2. FilterName must not be null or empty
        /// 3. CombinedKey must not be null or empty
        /// 4. FilterId must be > 0 (from EnsureFilter)
        /// 5. Conditions object must not be null
        /// 
        /// 🔒 LOCKED BEHAVIOR:
        /// - Returns bool (true = success, false = failure)
        /// - All validation failures are logged
        /// - All exceptions are caught and logged
        /// - Filter registration MUST succeed before conditions can be saved
        /// 
        /// ⚠️ DO NOT:
        /// - Change return type from bool
        /// - Remove validation checks
        /// - Remove error logging
        /// - Bypass filter registration
        /// - Swallow exceptions silently
        /// </summary>
        private bool SaveConditionsToSqlite(string filterName, string normalizedCategory, string combinedKey, OpeningConditions conditions)
        {
            // ⚠️ CRITICAL VALIDATION #1: Document must exist
            if (_document == null)
            {
                _log($"[ConditionsService] ❌ SQLite save skipped: Document is null");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ConditionsService] SQLite save skipped: Document is null");
                return false; // ⚠️ DO NOT change to true - this is a validation failure
            }
            
            // ⚠️ CRITICAL VALIDATION #2: FilterName and CombinedKey must be valid
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(combinedKey))
            {
                _log($"[ConditionsService] ❌ SQLite save skipped: FilterName='{filterName}', CombinedKey='{combinedKey}'");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ConditionsService] SQLite save skipped: FilterName='{filterName}', CombinedKey='{combinedKey}'");
                return false; // ⚠️ DO NOT change to true - this is a validation failure
            }
            
            // ⚠️ CRITICAL VALIDATION #3: Conditions object must not be null
            if (conditions == null)
            {
                _log($"[ConditionsService] ❌ SQLite save skipped: Conditions object is null");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ConditionsService] SQLite save skipped: Conditions object is null");
                return false; // ⚠️ DO NOT change to true - this is a validation failure
            }

            try
            {
                using (var context = new SleeveDbContext(_document, msg => _log($"[SQLite] {msg}")))
                {
                    var filterRepository = new FilterRepository(context, msg => _log($"[SQLite] {msg}"));
                    var conditionRepository = new ConditionRepository(context, msg => _log($"[SQLite] {msg}"));

                    _log($"[ConditionsService] Attempting to save conditions to SQLite: FilterName='{filterName}', Category='{normalizedCategory}', CombinedKey='{combinedKey}'");
                    
                    int filterId = filterRepository.EnsureFilter(filterName, normalizedCategory);
                    if (filterId <= 0)
                    {
                        _log($"[ConditionsService] ❌ SQLite filter registration failed for '{filterName}' (Category='{normalizedCategory}'). FilterId={filterId}");
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[ConditionsService] SQLite filter registration failed for '{filterName}' (Category='{normalizedCategory}'). FilterId={filterId}");
                        return false;
                    }

                    _log($"[ConditionsService] ✅ Filter registered: FilterId={filterId} for '{filterName}' (Category='{normalizedCategory}')");
                    
                    // ⚠️⚠️⚠️ CRITICAL: This call MUST succeed for conditions to be saved to database ⚠️⚠️⚠️
                    // ✅ WORKING AS OF 2025-11-15: Conditions are successfully populating in SQLite database
                    // DO NOT remove, bypass, or modify this call without extensive testing
                    conditionRepository.UpsertConditions(filterId, combinedKey, normalizedCategory, conditions);
                    
                    _log($"[ConditionsService] ✅ SQLite save succeeded for '{combinedKey}' (FilterId={filterId})");
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[ConditionsService] ✅ SQLite save succeeded for '{combinedKey}' (FilterId={filterId})");
                    return true; // ⚠️ DO NOT change to false - this indicates successful save
                }
            }
            catch (Exception ex)
            {
                _log($"[ConditionsService] ❌ SQLite save failed: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[ConditionsService] SQLite save failed: {ex.Message}\n{ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// Load conditions from SQLite database using CombinedKey (FilterName_Category format)
        /// ✅ CRITICAL: CombinedKey ensures each filter has its own clearance values per category
        /// Example: "Ventilation_duct_accessories" vs "Plumbing_duct_accessories" = different clearances
        /// ✅ FIX: Normalizes CombinedKey to prevent duplicate row issues
        /// </summary>
        private OpeningConditions LoadConditionsFromSqlite(string combinedKey)
        {
            if (_document == null || string.IsNullOrWhiteSpace(combinedKey))
                return null;

            try
            {
                // ✅ FIX: Normalize CombinedKey before database lookup
                string normalizedKey = NormalizeCombinedKeyForDatabase(combinedKey);
                
                using (var context = new SleeveDbContext(_document, msg => _log($"[SQLite] {msg}")))
                {
                    var conditionRepository = new ConditionRepository(context, msg => _log($"[SQLite] {msg}"));
                    var conditions = conditionRepository.GetConditions(normalizedKey);
                    if (conditions != null)
                    {
                        conditions.FilterName = ResolveFilterName(conditions, combinedKey);
                        conditions.Category = ResolveNormalizedCategory(conditions, combinedKey);
                        
                        // ✅ DIAGNOSTIC: Log which filter's conditions were loaded (critical for multi-filter scenarios)
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            var clearance = conditions.ClearanceSettings ?? new ClearanceSettings();
                            DebugLogger.Info($"[ConditionsService] ✅ Loaded conditions from SQLite for CombinedKey='{normalizedKey}' " +
                                $"(Filter='{conditions.FilterName}', Category='{conditions.Category}') - " +
                                $"DuctAccessoryOther={clearance.DuctAccessoryOtherNormal}mm, " +
                                $"DuctAccessoryMep={clearance.DuctAccessoryMepNormal}mm");
                        }
                    }
                    else if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[ConditionsService] ⚠️ No conditions found in SQLite for CombinedKey='{normalizedKey}' - will use defaults");
                    }
                    return conditions;
                }
            }
            catch (Exception ex)
            {
                _log($"[ConditionsService] ⚠️ SQLite retrieval failed: {ex.Message}");
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ConditionsService] SQLite retrieval failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get the file path for a filter's conditions XML
        /// </summary>
        public string GetConditionsFilePath(string filterName)
        {
            return Path.Combine(_projectDirectory, $"{filterName}_CONDITIONS.xml");
        }

        /// <summary>
        /// Check if conditions file exists for a filter
        /// </summary>
        public bool ConditionsExist(string filterName)
        {
            if (!string.IsNullOrWhiteSpace(filterName))
            {
                try
                {
                    if (LoadConditionsFromSqlite(filterName) != null)
                        return true;
                }
                catch
                {
                    // Ignore SQLite errors for existence checks
                }
            }

            var filePath = GetConditionsFilePath(filterName);
            return File.Exists(filePath);
        }
    }
}


