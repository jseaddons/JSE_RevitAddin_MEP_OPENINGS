using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

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
            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_document);
                if (!Directory.Exists(filtersDirectory))
                    return null;
                
                var mergedStorage = new ClashZoneStorage
                {
                    ClashZones = new List<ClashZone>(),
                    LastUpdated = DateTime.Now
                };
                
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
                return null;
            }
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
        /// Check if all file combos are already processed
        /// Used by IntersectionProcessor to determine Replace/Replay/FullDetection mode
        /// </summary>
        public bool AreAllFileCombosProcessed(
            List<string> selectedMepCategories,
            List<string> selectedReferenceFiles,
            List<string> selectedHostFiles)
        {
            if (selectedMepCategories == null || selectedMepCategories.Count == 0)
                return false;
            
            if (selectedReferenceFiles == null || selectedReferenceFiles.Count == 0)
                return false;
            
            if (selectedHostFiles == null || selectedHostFiles.Count == 0)
                return false;
            
            // Build all file combos from selections
            var allCombos = new List<(string LinkedFile, string HostFile)>();
            foreach (var refFile in selectedReferenceFiles)
            {
                foreach (var hostFile in selectedHostFiles)
                {
                    allCombos.Add((refFile, hostFile));
                }
            }
            
            // Check if all combos are processed for all categories
            foreach (var category in selectedMepCategories)
            {
                var processedKeys = GlobalIndexService.GetProcessedFileComboKeys(_document, category);
                
                foreach (var combo in allCombos)
                {
                    var normalizedCombo = new ProcessedFileCombo 
                    { 
                        LinkedFile = combo.LinkedFile, 
                        HostFile = combo.HostFile 
                    };
                    var comboKey = normalizedCombo.GetNormalizedKey();
                    
                    if (!processedKeys.Contains(comboKey))
                    {
                        Log($"[XML-CACHE] Combo not processed: Category={category}, Linked={combo.LinkedFile}, Host={combo.HostFile}");
                        return false;
                    }
                }
            }
            
            return true;
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
