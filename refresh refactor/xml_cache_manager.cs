using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        /// </summary>
        public XmlCache LoadAll(List<string> filterNames, List<string> categories)
        {
            var cache = new XmlCache();
            
            Log($"[XML-CACHE] Loading XML data once for reuse...");
            
            // Load Filter XML for each filter
            foreach (var filterName in filterNames ?? new List<string>())
            {
                if (string.IsNullOrWhiteSpace(filterName)) continue;
                
                var storage = LoadFilterXml(filterName, categories);
                if (storage != null)
                {
                    cache.FilterXml[filterName] = storage;
                    Log($"[XML-CACHE] ✅ Loaded Filter XML: {filterName} ({storage.ClashZones?.Count ?? 0} zones)");
                }
            }
            
            // Load Global XML for each category
            foreach (var category in categories ?? new List<string>())
            {
                if (string.IsNullOrWhiteSpace(category)) continue;
                
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
            
            Log($"[XML-CACHE] ✅ XML cache loaded: {cache.FilterXml.Count} filters, {cache.GlobalXml.Count} categories");
            
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
        
        private void Log(string message)
        {
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info(message);
            SafeFileLogger.SafeAppendText(_refreshLogName, $"[{DateTime.Now}] {message}\n");
        }
    }
}
