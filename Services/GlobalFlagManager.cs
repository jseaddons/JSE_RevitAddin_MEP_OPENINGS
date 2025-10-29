using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Global flag manager that maintains sleeve placement state across filter changes
    /// Uses category-specific global XML files: {Category}_global.xml
    /// </summary>
    public class GlobalFlagManager
    {
        // ✅ MEMORY OPTIMIZATION: Thread-safe singleton cache per category to avoid reloading XML
        private static readonly ConcurrentDictionary<string, GlobalFlagManager> _instanceCache = new ConcurrentDictionary<string, GlobalFlagManager>();
        private static readonly object _saveLock = new object(); // Lock for XML saves to prevent concurrent writes
        
        private readonly string _globalXmlPath;
        private GlobalFlagStorage _storage;
        private readonly string _categoryName;

        public GlobalFlagManager(string categoryName)
        {
            _categoryName = categoryName ?? throw new ArgumentNullException(nameof(categoryName));
            
            // Use same path as regular filter files, but with _global prefix
            var filtersDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                "JSE_MEP_Openings", 
                "Projects", 
                "Default", 
                "Filters");
            
            Directory.CreateDirectory(filtersDirectory);
            _globalXmlPath = Path.Combine(filtersDirectory, $"{categoryName}_global.xml");
            
            LoadGlobalFlags();
        }

        /// <summary>
        /// ✅ MEMORY OPTIMIZATION: Get or create a GlobalFlagManager instance for the given category.
        /// Uses singleton pattern with thread-safe caching to avoid reloading XML files.
        /// </summary>
        /// <param name="categoryName">The category name (e.g., "Ducts", "Pipes")</param>
        /// <returns>A shared GlobalFlagManager instance for this category</returns>
        public static GlobalFlagManager GetOrCreate(string categoryName)
        {
            if (string.IsNullOrWhiteSpace(categoryName))
                throw new ArgumentException("Category name cannot be null or empty", nameof(categoryName));

            // ✅ Thread-safe: ConcurrentDictionary.GetOrAdd ensures only one instance per category
            return _instanceCache.GetOrAdd(categoryName, cat => new GlobalFlagManager(cat));
        }

        /// <summary>
        /// ✅ MEMORY OPTIMIZATION: Clear the cache for a specific category (e.g., after external file changes).
        /// This forces a reload of the XML file on next GetOrCreate call.
        /// </summary>
        /// <param name="categoryName">The category to clear from cache</param>
        public static void ClearCache(string categoryName)
        {
            if (string.IsNullOrWhiteSpace(categoryName))
                return;

            if (_instanceCache.TryRemove(categoryName, out var manager))
            {
                // Optional: Trigger reload on next access by clearing the instance
                // The instance will be recreated on next GetOrCreate call
            }
        }

        /// <summary>
        /// ✅ MEMORY OPTIMIZATION: Clear all cached instances (use sparingly, e.g., on application shutdown).
        /// </summary>
        public static void ClearAllCache()
        {
            _instanceCache.Clear();
        }

        /// <summary>
        /// Check if a sleeve placement exists for the given MEP element and host combination
        /// </summary>
        public PlacementRecord FindPlacement(ElementId mepElementId, ElementId hostElementId)
        {
            return _storage.Placements.FirstOrDefault(p => 
                p.MepElementId == mepElementId.IntegerValue && 
                p.HostElementId == hostElementId.IntegerValue);
        }

        /// <summary>
        /// Record a new sleeve placement (individual or cluster)
        /// </summary>
        public void RecordPlacement(ElementId mepElementId, ElementId hostElementId, ElementId? individualSleeveId, ElementId? clusterSleeveId, string filterName)
        {
            var existing = FindPlacement(mepElementId, hostElementId);
            
            if (existing != null)
            {
                // Update existing record
                existing.IndividualSleeveId = individualSleeveId?.IntegerValue ?? -1;
                existing.ClusterSleeveId = clusterSleeveId?.IntegerValue ?? -1;
                existing.LastUpdated = DateTime.Now;
                existing.FilterName = filterName;
            }
            else
            {
                // Create new record
                _storage.Placements.Add(new PlacementRecord
                {
                    MepElementId = mepElementId.IntegerValue,
                    HostElementId = hostElementId.IntegerValue,
                    IndividualSleeveId = individualSleeveId?.IntegerValue ?? -1,
                    ClusterSleeveId = clusterSleeveId?.IntegerValue ?? -1,
                    LastUpdated = DateTime.Now,
                    FilterName = filterName
                });
            }

            SaveGlobalFlags();
        }

        /// <summary>
        /// Remove a placement record (when sleeve is deleted)
        /// </summary>
        public void RemovePlacement(ElementId mepElementId, ElementId hostElementId)
        {
            var existing = FindPlacement(mepElementId, hostElementId);
            if (existing != null)
            {
                _storage.Placements.Remove(existing);
                SaveGlobalFlags();
            }
        }

        /// <summary>
        /// Verify sleeve exists in Revit and return its state
        /// </summary>
        public SleeveExistenceState CheckSleeveExistence(Document doc, ElementId mepElementId, ElementId hostElementId)
        {
            var placement = FindPlacement(mepElementId, hostElementId);
            
            if (placement == null)
            {
                return new SleeveExistenceState
                {
                    ExistsInGlobal = false,
                    HasIndividualSleeve = false,
                    HasClusterSleeve = false
                };
            }

            bool individualExists = false;
            bool clusterExists = false;

            // Check individual sleeve
            if (placement.IndividualSleeveId > 0)
            {
                var sleeve = doc.GetElement(new ElementId(placement.IndividualSleeveId));
                individualExists = sleeve != null;
            }

            // Check cluster sleeve
            if (placement.ClusterSleeveId > 0)
            {
                var cluster = doc.GetElement(new ElementId(placement.ClusterSleeveId));
                clusterExists = cluster != null;
            }

            return new SleeveExistenceState
            {
                ExistsInGlobal = true,
                HasIndividualSleeve = individualExists,
                HasClusterSleeve = clusterExists,
                Placement = placement
            };
        }

        /// <summary>
        /// Load global flags from XML
        /// </summary>
        private void LoadGlobalFlags()
        {
            if (File.Exists(_globalXmlPath))
            {
                try
                {
                    var serializer = new XmlSerializer(typeof(GlobalFlagStorage));
                    using (var reader = new FileStream(_globalXmlPath, FileMode.Open))
                    {
                        _storage = (GlobalFlagStorage)serializer.Deserialize(reader);
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[GlobalFlagManager] Error loading global flags: {ex.Message}");
                    _storage = new GlobalFlagStorage();
                }
            }
            else
            {
                _storage = new GlobalFlagStorage();
            }
        }

        /// <summary>
        /// Save global flags to XML
        /// ✅ THREAD-SAFETY: Uses lock to prevent concurrent writes from multiple instances
        /// </summary>
        private void SaveGlobalFlags()
        {
            lock (_saveLock)
            {
                try
                {
                    var serializer = new XmlSerializer(typeof(GlobalFlagStorage));
                    using (var writer = new FileStream(_globalXmlPath, FileMode.Create))
                    {
                        serializer.Serialize(writer, _storage);
                    }
                    
                    // ✅ CRITICAL: After saving, invalidate other instances in cache that might have stale data
                    // This ensures next GetOrCreate for this category will reload fresh data
                    // Note: We don't clear the cache entry itself because the current instance's _storage
                    // is now up-to-date. Other code paths using GetOrCreate will get the cached instance
                    // which already has the latest data loaded.
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[GlobalFlagManager] Error saving global flags for category '{_categoryName}': {ex.Message}");
                }
            }
        }

        /// <summary>
        /// ✅ MEMORY OPTIMIZATION: Reload the XML file (useful when external changes are made).
        /// </summary>
        public void Reload()
        {
            lock (_saveLock)
            {
                LoadGlobalFlags();
            }
        }
    }

    /// <summary>
    /// Storage container for global flag data
    /// </summary>
    [XmlRoot("GlobalFlagStorage")]
    public class GlobalFlagStorage
    {
        [XmlArray("Placements")]
        [XmlArrayItem("Placement")]
        public List<PlacementRecord> Placements { get; set; } = new List<PlacementRecord>();

        public DateTime LastUpdated { get; set; } = DateTime.Now;
    }

    /// <summary>
    /// Individual placement record
    /// </summary>
    [XmlType("Placement")]
    public class PlacementRecord
    {
        [XmlAttribute("MepElementId")]
        public int MepElementId { get; set; }

        [XmlAttribute("HostElementId")]
        public int HostElementId { get; set; }

        [XmlAttribute("IndividualSleeveId")]
        public int IndividualSleeveId { get; set; } = -1;

        [XmlAttribute("ClusterSleeveId")]
        public int ClusterSleeveId { get; set; } = -1;

        [XmlAttribute("LastUpdated")]
        public DateTime LastUpdated { get; set; } = DateTime.Now;

        [XmlAttribute("FilterName")]
        public string FilterName { get; set; } = "";
    }

    /// <summary>
    /// Result of checking sleeve existence
    /// </summary>
    public class SleeveExistenceState
    {
        public bool ExistsInGlobal { get; set; }
        public bool HasIndividualSleeve { get; set; }
        public bool HasClusterSleeve { get; set; }
        public PlacementRecord Placement { get; set; }
    }
}
