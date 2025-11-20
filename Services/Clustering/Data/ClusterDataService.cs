using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data
{
    /// <summary>
    /// Service for loading and caching clash zone data for clustering operations.
    /// DATABASE-FIRST ARCHITECTURE: Primary data source is SQLite database.
    /// XML fallback is LEGACY-ONLY for backward compatibility with pre-migration projects.
    /// </summary>
    public class ClusterDataService : IClusterDataService
    {
        private readonly Document _doc;
        private readonly Dictionary<long, ClashZone> _clashZoneCache = new Dictionary<long, ClashZone>();

        public int CacheCount => _clashZoneCache.Count;

        public ClusterDataService(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// Load clash zones from database (primary) or XML files (LEGACY fallback for pre-migration projects).
        /// ⚠️ DATABASE-FIRST: New projects use SQLite exclusively. XML is backward compatibility only.
        /// Returns zones with valid SleeveInstanceId for clustering.
        /// </summary>
        public List<ClashZone> LoadClashZonesFromRegularXml(string xmlFilePath, string targetCategory, Document doc)
        {
            var clashZones = new List<ClashZone>();

            try
            {
                // ✅ DATABASE-FIRST ARCHITECTURE - Primary data source is SQLite
                // XML fallback is LEGACY-ONLY for backward compatibility with pre-migration projects
                if (!string.IsNullOrEmpty(targetCategory) && doc != null)
                {
                    try
                    {
                        using (var dbContext = new SleeveDbContext(doc))
                        {
                            var repository = new ClashZoneRepository(dbContext);
                            var dbZones = repository.GetClashZonesByCategory(targetCategory);
                            
                            if (dbZones != null && dbZones.Count > 0)
                            {
                                // Filter to only zones with SleeveInstanceId > 0 (placed sleeves)
                                var placedZones = dbZones.Where(z => z != null && z.SleeveInstanceId > 0 && !z.IsClusterResolved).ToList();
                                
                                foreach (var cz in placedZones)
                                {
                                    // ✅ CRITICAL: Reconstruct SleevePlacementPoint from database properties
                                    cz.EnsureSleevePlacementPointReconstructed();
                                    
                                    // ✅ ROTATION DATA FROM DB: MepElementRotationAngle is already loaded from database
                                    // via GetClashZonesByCategory -> MepRotationAngleRad column (see ClashZoneRepository line 1235).
                                    // This rotation angle is used by DetermineDominantRotationAngle to calculate cluster rotation.
                                    // No additional loading needed - rotation data is already in the ClashZone object from DB.
                                    
                                    clashZones.Add(cz);
                                }
                                
                                // ✅ DATABASE SUCCESS: Return database data (XML not needed for modern projects)
                                return clashZones;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Fall through to XML loading (legacy fallback only)
                    }
                }
                
                // ⚠️ LEGACY FALLBACK: Load from XML only if database has no data
                // This path is for pre-migration projects or database initialization failure
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_doc);

                if (!Directory.Exists(filtersDirectory))
                {
                    return clashZones;
                }

                // Load from regular XML files
                var xmlFiles = string.IsNullOrEmpty(xmlFilePath)
                    ? Directory.GetFiles(filtersDirectory, "*.xml")
                    : new[] { xmlFilePath };

                // ⚠️ LEGACY: Multi-threading for XML loading (backward compatibility only)
                // Modern projects use database exclusively - this code path is rarely executed
                if (xmlFiles.Length > 1)
                {
                    // Parallel XML loading for multiple files (file I/O only, no Revit API)
                    var loadedZones = new System.Collections.Concurrent.ConcurrentBag<ClashZone>();

                    System.Threading.Tasks.Parallel.ForEach(xmlFiles, xmlFile =>
                    {
                        try
                        {
                            // Skip CONDITIONS files
                            if (Path.GetFileName(xmlFile).Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase))
                                return;

                            // Skip _CLUSTER.xml files - we want regular XML files
                            if (Path.GetFileName(xmlFile).Contains("_CLUSTER", StringComparison.OrdinalIgnoreCase))
                                return;

                            if (!File.Exists(xmlFile))
                                return;

                            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                            using (var reader = new StreamReader(xmlFile))
                            {
                                var filter = (OpeningFilter)serializer.Deserialize(reader);

                                // Load from hierarchical structure (PRIMARY)
                                if (filter?.ClashZoneStorage?.Filters != null)
                                {
                                    foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                                    {
                                        if (filterGroup?.FileCombos != null)
                                        {
                                            foreach (var fileCombo in filterGroup.FileCombos)
                                            {
                                                if (fileCombo?.ClashZones != null)
                                                {
                                                    foreach (var cz in fileCombo.ClashZones)
                                                    {
                                                        if (!string.IsNullOrEmpty(targetCategory))
                                                        {
                                                            if (!string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                                                continue;
                                                        }

                                                        if (cz.SleeveInstanceId > 0)
                                                        {
                                                            cz.EnsureSleevePlacementPointReconstructed();
                                                            loadedZones.Add(cz);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                // Load from flat structure (BACKWARD COMPATIBILITY)
                                foreach (var cz in GetStorageZones(filter))
                                {
                                    if (cz == null) continue;
                                    if (!string.IsNullOrEmpty(targetCategory) &&
                                        !string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                    {
                                        continue;
                                    }

                                    if (cz.SleeveInstanceId > 0 && !loadedZones.Any(c => c.Id == cz.Id))
                                    {
                                        cz.EnsureSleevePlacementPointReconstructed();
                                        loadedZones.Add(cz);
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            // Skip file with errors
                        }
                    });

                    clashZones.AddRange(loadedZones);
                }
                else
                {
                    // Sequential loading (fallback or single file)
                    foreach (var xmlFile in xmlFiles)
                    {
                        // Skip CONDITIONS files
                        if (Path.GetFileName(xmlFile).Contains("CONDITIONS", StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Skip _CLUSTER.xml files - we want regular XML files
                        if (Path.GetFileName(xmlFile).Contains("_CLUSTER", StringComparison.OrdinalIgnoreCase))
                            continue;

                        if (File.Exists(xmlFile))
                        {
                            var serializer = new System.Xml.Serialization.XmlSerializer(typeof(OpeningFilter));
                            using (var reader = new StreamReader(xmlFile))
                            {
                                var filter = (OpeningFilter)serializer.Deserialize(reader);

                                // 1. Load from hierarchical structure (PRIMARY)
                                if (filter?.ClashZoneStorage?.Filters != null)
                                {
                                    foreach (var filterGroup in filter.ClashZoneStorage.Filters)
                                    {
                                        if (filterGroup?.FileCombos != null)
                                        {
                                            foreach (var fileCombo in filterGroup.FileCombos)
                                            {
                                                if (fileCombo?.ClashZones != null)
                                                {
                                                    foreach (var cz in fileCombo.ClashZones)
                                                    {
                                                        // Filter by target category during loading
                                                        if (!string.IsNullOrEmpty(targetCategory))
                                                        {
                                                            if (!string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                                            {
                                                                continue;
                                                            }
                                                        }

                                                        // Only process clash zones with valid SleeveInstanceId (placed sleeves)
                                                        if (cz.SleeveInstanceId > 0)
                                                        {
                                                            // ✅ CRITICAL: Reconstruct SleevePlacementPoint from XML-serializable properties
                                                            cz.EnsureSleevePlacementPointReconstructed();
                                                            clashZones.Add(cz);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                // 2. Load from flat structure (BACKWARD COMPATIBILITY)
                                foreach (var cz in GetStorageZones(filter))
                                {
                                    if (cz == null) continue;
                                    if (!string.IsNullOrEmpty(targetCategory) &&
                                        !string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                                    {
                                        continue;
                                    }

                                    if (cz.SleeveInstanceId > 0 && !clashZones.Any(c => c.Id == cz.Id))
                                    {
                                        cz.EnsureSleevePlacementPointReconstructed();
                                        clashZones.Add(cz);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Return whatever we loaded successfully
            }
            
            return clashZones;
        }

        /// <summary>
        /// Load clash zone cache for cleanup operations.
        /// Populates internal cache with clash zones for fast lookups during clustering.
        /// </summary>
        public void LoadClashZoneCacheForCleanup(string xmlFilePath, string targetCategory, Document doc, string filterName)
        {
            // ✅ CRITICAL: Clear cache before reloading to get fresh data
            _clashZoneCache.Clear();
            
            // Load from database or XML
            var clashZones = LoadClashZonesFromRegularXml(xmlFilePath, targetCategory, doc);
            
            // Populate cache
            foreach (var cz in clashZones)
            {
                if (cz != null && cz.MepElementIdValue > 0)
                {
                    cz.IsCurrentClash = true;
                    _clashZoneCache[cz.MepElementIdValue] = cz;
                }
            }
        }

        /// <summary>
        /// Populate cache from already-loaded clash zones.
        /// Optimization to avoid duplicate XML/database loading.
        /// </summary>
        public void LoadClashZoneCacheFromLoadedClashZones(List<ClashZone> clashZones, string targetCategory = null)
        {
            _clashZoneCache.Clear();
            
            try
            {
                foreach (var cz in clashZones)
                {
                    if (cz == null)
                        continue;

                    if (!string.IsNullOrEmpty(targetCategory) &&
                        !string.Equals(cz.MepElementCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (cz.SleeveInstanceId <= 0 && cz.ClusterSleeveInstanceId <= 0 && cz.AfterClusterSleevePlacedSleeveInstanceId <= 0)
                    {
                        continue;
                    }

                    if (cz.MepElementIdValue <= 0)
                        continue;

                    cz.EnsureSleevePlacementPointReconstructed();
                    cz.IsCurrentClash = true;
                    _clashZoneCache[cz.MepElementIdValue] = cz;
                }
            }
            catch (Exception)
            {
                // Cache partially populated - continue
            }
        }

        /// <summary>
        /// Get clash zone from cache by MEP element ID.
        /// Returns null if not found.
        /// </summary>
        public ClashZone GetClashZoneFromCache(long mepElementId)
        {
            _clashZoneCache.TryGetValue(mepElementId, out var clashZone);
            return clashZone;
        }

        /// <summary>
        /// Clear the clash zone cache.
        /// </summary>
        public void ClearCache()
        {
            _clashZoneCache.Clear();
        }

        /// <summary>
        /// Get storage zones from filter (flat structure).
        /// </summary>
        private static List<ClashZone> GetStorageZones(OpeningFilter filter)
        {
            return filter?.ClashZoneStorage?.AllZones ?? new List<ClashZone>();
        }
    }
}
