using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Caches all XML data in memory for a refresh operation.
    /// Eliminates redundant XML loading (was 4+ loads, now 1 load per file).
    /// </summary>
    public class XmlCache
    {
        // Filter XML cache: filterName -> ClashZoneStorage
        public Dictionary<string, ClashZoneStorage> FilterXml { get; private set; }
        
        // Global XML cache: category -> CategoryGlobalIndex
        public Dictionary<string, CategoryGlobalIndex> GlobalXml { get; private set; }
        
        // Processed file combos: category -> HashSet of normalized combo keys
        public Dictionary<string, HashSet<string>> ProcessedCombos { get; private set; }
        
        // Resolved GUIDs: category -> HashSet of resolved GUIDs
        public Dictionary<string, HashSet<Guid>> ResolvedGuids { get; private set; }
        
        public XmlCache()
        {
            FilterXml = new Dictionary<string, ClashZoneStorage>(StringComparer.OrdinalIgnoreCase);
            GlobalXml = new Dictionary<string, CategoryGlobalIndex>(StringComparer.OrdinalIgnoreCase);
            ProcessedCombos = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            ResolvedGuids = new Dictionary<string, HashSet<Guid>>(StringComparer.OrdinalIgnoreCase);
        }
        
        public void Clear()
        {
            FilterXml?.Clear();
            GlobalXml?.Clear();
            ProcessedCombos?.Clear();
            ResolvedGuids?.Clear();
        }
    }
    
    public class XmlCacheManager
    {
        private readonly Document _document;
        private readonly string _refreshLogName;
        
        public XmlCacheManager(Document document, string refreshLogName)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _refreshLogName = refreshLogName;
        }
        
        /// <summary>
        /// Loads ALL XML data once at start of refresh.
        /// Eliminates 4+ redundant XML loads (was: load -> sync -> check -> load again).
        /// ✅ OPTIMIZATION: Parallelized file I/O operations (non-Revit operations).
        /// </summary>
        public XmlCache LoadAll(List<string> filterNames, List<string> categories)
        {
            var cache = new XmlCache();
            
            Log($"[XML-CACHE] Loading XML data once for reuse (parallelized file I/O)...");
            
            var sw = System.Diagnostics.Stopwatch.StartNew();
            
            // ✅ PARALLELIZATION: Load Filter XML files in parallel (pure file I/O, no Revit API)
            var filterResults = new ConcurrentDictionary<string, ClashZoneStorage>();
            var validFilterNames = (filterNames ?? new List<string>())
                .Where(fn => !string.IsNullOrWhiteSpace(fn))
                .ToList();
            
            if (validFilterNames.Count > 0)
            {
                Parallel.ForEach(validFilterNames, filterName =>
                {
                    var storage = LoadFilterXml(filterName, categories);
                    if (storage != null)
                    {
                        filterResults[filterName] = storage;
                        Log($"[XML-CACHE] ✅ Loaded Filter XML: {filterName} ({storage.ClashZones?.Count ?? 0} zones)");
                    }
                });
            }
            
            // Copy results to cache (thread-safe - ConcurrentDictionary)
            foreach (var kvp in filterResults)
            {
                cache.FilterXml[kvp.Key] = kvp.Value;
            }
            
            // ✅ PARALLELIZATION: Load Global XML files in parallel (pure file I/O, no Revit API)
            // Note: GlobalIndexService.LoadOrCreate() may use Revit API, so we need to check
            // For now, keeping sequential for Global XML to be safe, but Filter XML is parallelized
            var validCategories = (categories ?? new List<string>())
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .ToList();
            
            foreach (var category in validCategories)
            {
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                cache.GlobalXml[category] = globalIndex;
                
                // Extract processed combos
                var processedKeys = GlobalIndexService.GetProcessedFileComboKeys(_document, category);
                cache.ProcessedCombos[category] = processedKeys;
                
                // Extract resolved GUIDs
                var resolvedGuids = GlobalIndexService.GetResolvedGuidsForCategories(_document, new HashSet<string> { category });
                cache.ResolvedGuids[category] = resolvedGuids;
                
                Log($"[XML-CACHE] ✅ Loaded Global XML: {category} ({processedKeys.Count} combos, {resolvedGuids.Count} resolved)");
            }
            
            sw.Stop();
            
            Log($"[XML-CACHE] ✅ XML cache loaded: {cache.FilterXml.Count} filters, {cache.GlobalXml.Count} categories in {sw.ElapsedMilliseconds}ms");
            
            return cache;
        }
        
        private ClashZoneStorage LoadFilterXml(string filterName, List<string> categories)
        {
            var mergedStorage = new ClashZoneStorage
            {
                ClashZones = new List<ClashZone>(),
                LastUpdated = DateTime.Now
            };
            
            // ✅ PHASE SQLITE-2: Load from SQLite FIRST (primary source), XML as fallback
            if (DeploymentConfiguration.UseSqliteAsPrimary)
            {
                try
                {
                    var sqliteZones = LoadFromSqlite(filterName, categories ?? new List<string>());
                    if (sqliteZones != null && sqliteZones.Count > 0)
                    {
                        mergedStorage.ClashZones.AddRange(sqliteZones);
                        Log($"[XML-CACHE] ✅ PHASE 2: Loaded {sqliteZones.Count} zones from SQLite (PRIMARY) for filter '{filterName}'");
                        return mergedStorage;
                    }
                    else
                    {
                        Log($"[XML-CACHE] PHASE 2: SQLite has no zones for '{filterName}', falling back to XML");
                    }
                }
                catch (Exception sqliteEx)
                {
                    Log($"[XML-CACHE] PHASE 2: SQLite load failed for '{filterName}', falling back to XML: {sqliteEx.Message}");
                }
            }
            
            // Fallback to XML (legacy mode or when SQLite has no data)
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_document);
                if (!Directory.Exists(filtersDirectory))
                    return mergedStorage.ClashZones.Count > 0 ? mergedStorage : null;
                
                // Load category-specific XML files
                foreach (var category in categories ?? new List<string>())
                {
                    var pattern = $"{filterName}_{category.ToLower().Replace(" ", "_")}.xml";
                    var matchingFiles = Directory.GetFiles(filtersDirectory, pattern);
                    
                    if (matchingFiles.Length > 0)
                    {
                        var xmlFile = matchingFiles.First();
                        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                        
                        using (var reader = new StreamReader(xmlFile))
                        {
                            var filter = (OpeningFilter)serializer.Deserialize(reader);
                            if (filter?.ClashZoneStorage?.AllZones != null)
                            {
                                mergedStorage.ClashZones.AddRange(filter.ClashZoneStorage.AllZones);
                            }
                        }
                    }
                }
                
                return mergedStorage.ClashZones.Count > 0 ? mergedStorage : null;
            }
            catch (Exception ex)
            {
                Log($"[XML-CACHE] ⚠️ Error loading Filter XML '{filterName}': {ex.Message}");
                return mergedStorage.ClashZones.Count > 0 ? mergedStorage : null;
            }
        }
        
        /// <summary>
        /// ✅ PHASE SQLITE-2: Load clash zones from SQLite database (primary source)
        /// </summary>
        private List<ClashZone> LoadFromSqlite(string filterName, List<string> categories)
        {
            var allZones = new List<ClashZone>();
            
            if (string.IsNullOrWhiteSpace(filterName) || categories == null || categories.Count == 0)
                return allZones;
            
            try
            {
                SleeveDbContext context;
                bool disposeContext = false;
                if (OptimizationFlags.ReuseDbContextDuringRefresh)
                {
                    context = SharedDbContextProvider.GetOrCreate(_document, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[XML-CACHE][SQLite] {msg}");
                    });
                }
                else
                {
                    context = new SleeveDbContext(_document, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[XML-CACHE][SQLite] {msg}");
                    });
                    disposeContext = true;
                }

                try
                {
                    var repository = new ClashZoneRepository(context, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[XML-CACHE][SQLite] {msg}");
                    });
                    
                    // Load zones for each category
                    // ✅ NOTE: ReadyForPlacementFlag will be set AFTER flag manager resets flags for deleted sleeves
                    // This ensures we check unresolved status AFTER flags are properly reset
                    foreach (var category in categories)
                    {
                        if (string.IsNullOrWhiteSpace(category))
                            continue;
                        
                        try
                        {
                            // Load all zones (not just unresolved) for refresh operations
                            var categoryZones = repository.GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
                            
                            foreach (var zone in categoryZones)
                            {
                                if (zone != null)
                                {
                                    zone.EnsureSleevePlacementPointReconstructed();
                                    zone.EnsureSleevePlacementPointActiveDocumentReconstructed();
                                    
                                    // ✅ USE DATABASE VALUE: ReadyForPlacement is already loaded from ReadyForPlacementFlag column
                                    // At START of refresh, all flags were reset to 0
                                    // ReadyForPlacementFlag will be set to 1 AFTER flag manager resets flags for deleted sleeves
                                    // (in refresh_service_refactored.cs, after Phase 5B)
                                    
                                    allZones.Add(zone);
                                }
                            }
                            
                            if (categoryZones.Count > 0 && !DeploymentConfiguration.DeploymentMode)
                            {
                                Log($"[XML-CACHE] ✅ SQLite loaded {categoryZones.Count} zones for filter '{filterName}', category '{category}' (ReadyForPlacementFlag will be set after flag reset)");
                            }
                        }
                        catch (Exception categoryEx)
                        {
                            Log($"[XML-CACHE] SQLite load failed for category '{category}': {categoryEx.Message}");
                        }
                    }
                }
                finally
                {
                    if (disposeContext)
                        context?.Dispose();
                }
            }
            catch (Exception ex)
            {
                Log($"[XML-CACHE] SQLite load failed: {ex.Message}");
            }
            
            return allZones;
        }
        
        /// <summary>
        /// Check if file combo is already processed (O(1) lookup from cache)
        /// </summary>
        public bool IsComboProcessed(XmlCache cache, string category, string linkedFile, string hostFile)
        {
            if (!cache.ProcessedCombos.TryGetValue(category, out var combos))
                return false;
            
            var combo = new ProcessedFileCombo
            {
                LinkedFile = linkedFile,
                HostFile = hostFile
            };

            return combos.Contains(combo.GetNormalizedKey());
        }
        
        /// <summary>
        /// Check if GUID is already resolved (O(1) lookup from cache)
        /// </summary>
        public bool IsGuidResolved(XmlCache cache, string category, Guid guid)
        {
            if (!cache.ResolvedGuids.TryGetValue(category, out var guids))
                return false;
            
            return guids.Contains(guid);
        }
        
        /// <summary>
        /// Check if all file combos are already processed (database-only check)
        /// Used by IntersectionProcessor to determine Replace/Replay/FullDetection mode
        /// ✅ DATABASE-ONLY: Checks IsFilterComboNew flag in FileCombos table (no XML dependency)
        /// </summary>
        public bool AreAllFileCombosProcessed(
            List<string> selectedFilterNames,
            List<string> selectedMepCategories,
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles)
        {
            if (selectedFilterNames == null || selectedFilterNames.Count == 0)
                return false;
            
            if (selectedMepCategories == null || selectedMepCategories.Count == 0)
                return false;
            
            if (selectedReferenceFiles == null || selectedReferenceFiles.Count == 0)
                return false;
            
            if (selectedHostFiles == null || selectedHostFiles.Count == 0)
                return false;
            
            try
            {
                SleeveDbContext dbContext;
                bool disposeDb = false;
                if (OptimizationFlags.ReuseDbContextDuringRefresh)
                {
                    dbContext = SharedDbContextProvider.GetOrCreate(_document);
                }
                else
                {
                    dbContext = new SleeveDbContext(_document);
                    disposeDb = true;
                }

                try
                {
                    var filterRepository = new FilterRepository(dbContext, _ => { });
                    
                    // Build all file combos from selections
                    var allCombos = new List<(string LinkedFileKey, string HostFileKey)>();
                    foreach (var refFile in selectedReferenceFiles)
                    {
                        foreach (var hostFile in selectedHostFiles)
                        {
                            // Normalize file keys (same logic as used in database)
                            var linkedKey = NormalizeDocumentKey(refFile);
                            var hostKey = NormalizeDocumentKey(hostFile);
                            allCombos.Add((linkedKey, hostKey));
                        }
                    }
                    
                    // Check if all combos are processed for all filter+category combinations
                    foreach (var filterName in selectedFilterNames)
                    {
                        foreach (var category in selectedMepCategories)
                        {
                            int filterId = filterRepository.GetFilterId(filterName, category);
                            if (filterId <= 0)
                            {
                                // Filter doesn't exist = not processed
                                Log($"[DB-COMBO-CHECK] Filter not found: Filter='{filterName}', Category='{category}' → not processed");
                                return false;
                            }
                            
                            // Check each file combo for this filter+category
                            foreach (var combo in allCombos)
                            {
                                using (var cmd = dbContext.Connection.CreateCommand())
                                {
                                    cmd.CommandText = @"
                                        SELECT IsFilterComboNew 
                                        FROM FileCombos 
                                        WHERE FilterId = @FilterId 
                                          AND Category = @Category 
                                          AND LinkedFileKey = @LinkedFileKey 
                                          AND HostFileKey = @HostFileKey";
                                    cmd.Parameters.AddWithValue("@FilterId", filterId);
                                    cmd.Parameters.AddWithValue("@Category", category);
                                    cmd.Parameters.AddWithValue("@LinkedFileKey", combo.LinkedFileKey);
                                    cmd.Parameters.AddWithValue("@HostFileKey", combo.HostFileKey);
                                    
                                    var result = cmd.ExecuteScalar();
                                    if (result == null || result == DBNull.Value)
                                    {
                                        // File combo doesn't exist = not processed
                                        Log($"[DB-COMBO-CHECK] Combo not found: Filter='{filterName}', Category='{category}', Linked='{combo.LinkedFileKey}', Host='{combo.HostFileKey}' → not processed");
                                        return false;
                                    }
                                    
                                    int isNew = Convert.ToInt32(result);
                                    if (isNew != 0)
                                    {
                                        // IsFilterComboNew = 1 means not processed yet
                                        Log($"[DB-COMBO-CHECK] Combo not processed: Filter='{filterName}', Category='{category}', Linked='{combo.LinkedFileKey}', Host='{combo.HostFileKey}' → IsFilterComboNew={isNew}");
                                        return false;
                                    }
                                }
                            }
                        }
                    }
                    
                    Log($"[DB-COMBO-CHECK] ✅ All file combos are processed (IsFilterComboNew=0) for all filter+category combinations");
                    return true;
                }
                finally
                {
                    if (disposeDb)
                        dbContext?.Dispose();
                }
            }
            catch (Exception ex)
            {
                Log($"[DB-COMBO-CHECK] ❌ Error checking file combos in database: {ex.Message}");
                return false; // On error, assume not processed to be safe
            }
        }
        
        /// <summary>
        /// Normalizes document key (same logic as ClashZoneRepository.NormalizeDocumentKey)
        /// </summary>
        private static string NormalizeDocumentKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "Unknown";
            }

            var trimmed = value.Trim();

            // Strip ": <number> : location Shared" style suffixes
            var locationMatch = System.Text.RegularExpressions.Regex.Match(
                trimmed,
                @":\s*\d+\s*:\s*location\s+Shared",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (locationMatch.Success)
            {
                trimmed = trimmed.Substring(0, locationMatch.Index).Trim();
            }

            // Remove trailing "(xx elements)" or similar
            var parenIndex = trimmed.IndexOf('(');
            if (parenIndex >= 0)
            {
                trimmed = trimmed.Substring(0, parenIndex).Trim();
            }

            // If the string is a full path, reduce to file name
            trimmed = System.IO.Path.GetFileName(trimmed);

            // Drop extension
            var withoutExtension = System.IO.Path.GetFileNameWithoutExtension(trimmed);
            if (string.IsNullOrWhiteSpace(withoutExtension))
            {
                withoutExtension = trimmed;
            }

            return withoutExtension.Trim();
        }
        
        /// <summary>
        /// Load existing clash zones from XML cache
        /// Used by IntersectionProcessor.PrepareExistingZones()
        /// </summary>
        public List<ClashZone> LoadExistingClashZones(
            List<string> selectedFilterNames,
            List<string> selectedMepCategories)
        {
            var allZones = new List<ClashZone>();
            
            // Load from cache if available
            var cache = LoadAll(selectedFilterNames ?? new List<string>(), selectedMepCategories ?? new List<string>());
            
            foreach (var filterName in selectedFilterNames ?? new List<string>())
            {
                if (cache.FilterXml.TryGetValue(filterName, out var storage))
                {
                    if (storage?.ClashZones != null)
                    {
                        allZones.AddRange(storage.ClashZones);
                    }
                }
            }
            
            Log($"[XML-CACHE] Loaded {allZones.Count} existing clash zones from cache");
            return allZones;
        }
        
        /// <summary>
        /// Update cache with new clash zones
        /// Used by IntersectionProcessor.PostProcess()
        /// </summary>
        public void UpdateCache(XmlCache cache, List<ClashZone> clashZones)
        {
            if (cache == null || clashZones == null || clashZones.Count == 0)
                return;
            
            // Group by filter name (extract from clash zones)
            // For now, update all filter XML entries that match the categories
            foreach (var kvp in cache.FilterXml.ToList())
            {
                var filterName = kvp.Key;
                var storage = kvp.Value;
                
                if (storage?.ClashZones == null)
                    storage.ClashZones = new List<ClashZone>();
                
                // Merge new zones for this filter (avoid duplicates)
                var existingIds = new HashSet<Guid>(storage.ClashZones.Select(z => z.Id));
                var matchingZones = clashZones.Where(cz => 
                {
                    // Match zones to filters by category (simplified - could be enhanced)
                    return true; // For now, add all zones to all filters
                }).ToList();
                
                foreach (var zone in matchingZones)
                {
                    if (!existingIds.Contains(zone.Id))
                    {
                        storage.ClashZones.Add(zone);
                    }
                }
            }
            
            Log($"[XML-CACHE] Updated cache with {clashZones.Count} clash zones");
        }
        
        private void Log(string message)
        {
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info(message);
            SafeFileLogger.SafeAppendText(_refreshLogName, $"[{DateTime.Now}] {message}\n");
        }
    }
}
