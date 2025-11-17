using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Shared data-access helper that loads clash zones from SQLite (primary) with optional XML fallback.
    /// Centralises the logic so placement, clustering, and refresh code paths behave consistently.
    /// </summary>
    public class ClashZoneDataService
    {
        private readonly Document _document;
        private readonly Action<string>? _log;

        public ClashZoneDataService(Document document, Action<string>? log = null)
        {
            _document = document;
            _log = log;
        }

        private void Log(string message)
        {
            if (!string.IsNullOrWhiteSpace(message))
                _log?.Invoke(message);
        }

        /// <summary>
        /// Load clash zones for the specified filter/category. SQLite is queried first; if it returns no rows,
        /// the method falls back to XML to maintain backward compatibility.
        /// </summary>
        public List<ClashZone> LoadClashZonesForCategory(string filterName, string category, bool unresolvedOnly = false)
        {
            var result = new List<ClashZone>();
            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return result;

            try
            {
                Log($"[ClashZoneDataService] Loading zones for filter '{filterName}', category '{category}', unresolvedOnly={unresolvedOnly}");
                var sqliteZones = LoadFromSqlite(filterName, category, unresolvedOnly);
                Log($"[ClashZoneDataService] SQLite returned {sqliteZones.Count} zones for filter '{filterName}', category '{category}', unresolvedOnly={unresolvedOnly}");
                
                if (sqliteZones.Count > 0)
                {
                    Log($"[ClashZoneDataService] ✅ Using SQLite zones (PRIMARY) for filter '{filterName}', category '{category}' - DATA SOURCE: DATABASE");
                    // ✅ CRITICAL: Log data source for placement debugging
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var placementLogPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                        try
                        {
                            File.AppendAllText(placementLogPath, $"[{DateTime.Now:HH:mm:ss}] [DATA-SOURCE] ✅ Using DATABASE for filter '{filterName}', category '{category}' ({sqliteZones.Count} zones)\n");
                        }
                        catch { }
                    }
                    return sqliteZones;
                }

                Log($"[ClashZoneDataService] ⚠️ SQLite returned 0 zones for filter '{filterName}', category '{category}' (unresolvedOnly={unresolvedOnly}). Falling back to XML.");
                var xmlZones = LoadFromXml(filterName, category);
                Log($"[ClashZoneDataService] XML returned {xmlZones.Count} zones for filter '{filterName}', category '{category}' - DATA SOURCE: XML");
                // ✅ CRITICAL: Log data source for placement debugging
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var placementLogPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                    try
                    {
                        File.AppendAllText(placementLogPath, $"[{DateTime.Now:HH:mm:ss}] [DATA-SOURCE] ⚠️ Using XML (FALLBACK) for filter '{filterName}', category '{category}' ({xmlZones.Count} zones) - SQLite returned 0 zones\n");
                    }
                    catch { }
                }
                result.AddRange(xmlZones);
            }
            catch (Exception ex)
            {
                Log($"[ClashZoneDataService] Error loading zones for filter '{filterName}', category '{category}': {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Load clash zones for multiple categories. Results are de-duplicated by clash zone Id.
        /// </summary>
        public List<ClashZone> LoadClashZones(string filterName, IEnumerable<string> categories, bool unresolvedOnly = false)
        {
            var allZones = new List<ClashZone>();
            if (categories == null)
                return allZones;

            var seen = new HashSet<Guid>();
            foreach (var category in categories)
            {
                var zones = LoadClashZonesForCategory(filterName, category, unresolvedOnly);
                foreach (var zone in zones)
                {
                    if (zone == null || seen.Contains(zone.Id))
                        continue;

                    seen.Add(zone.Id);
                    allZones.Add(zone);
                }
            }

            return allZones;
        }

        /// <summary>
        /// ✅ DATABASE OPTIMIZATION: Load clash zones with individual sleeves for clustering.
        /// Uses optimized database query that only returns zones with sleeves that need clustering.
        /// </summary>
        public List<ClashZone> LoadZonesWithSleevesForClustering(string filterName, string category)
        {
            var zones = new List<ClashZone>();

            if (string.IsNullOrWhiteSpace(filterName) || string.IsNullOrWhiteSpace(category))
                return zones;

            try
            {
                var sqliteZones = LoadZonesWithSleevesFromSqlite(filterName, category);
                if (sqliteZones.Count > 0)
                    return sqliteZones;

                Log($"[ClashZoneDataService] SQLite returned 0 zones with sleeves for clustering. Filter '{filterName}', category '{category}'.");
            }
            catch (Exception ex)
            {
                Log($"[ClashZoneDataService] Error loading zones with sleeves for clustering: {ex.Message}");
            }

            return zones;
        }

        private List<ClashZone> LoadZonesWithSleevesFromSqlite(string filterName, string category)
        {
            var zones = new List<ClashZone>();

            if (_document == null)
            {
                Log("[ClashZoneDataService] Document context unavailable for SQLite load.");
                return zones;
            }

            try
            {
                using (var context = new SleeveDbContext(_document, msg => Log($"[ClashZoneDataService][SQLite] {msg}")))
                {
                    var repository = new ClashZoneRepository(context, msg => Log($"[ClashZoneDataService][SQLite] {msg}"));
                    
                    // ✅ DATABASE OPTIMIZATION: Use optimized method for clustering
                    // This uses SleeveState index and only returns zones with individual sleeves
                    // ✅ Use GetClashZonesByFilter with SleeveState filter for clustering (zones with individual sleeves)
                    var dbZones = repository.GetClashZonesByFilter(filterName, category, unresolvedOnly: false)
                        .Where(z => z.SleeveInstanceId > 0 && z.ClusterSleeveInstanceId <= 0)
                        .ToList();

                    foreach (var zone in dbZones)
                    {
                        if (zone == null)
                            continue;

                        zone.EnsureSleevePlacementPointReconstructed();
                        zone.EnsureSleevePlacementPointActiveDocumentReconstructed();
                        zones.Add(zone);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"[ClashZoneDataService] SQLite load failed for clustering zones: {ex.Message}");
            }

            return zones;
        }

        private List<ClashZone> LoadFromSqlite(string filterName, string category, bool unresolvedOnly)
        {
            var zones = new List<ClashZone>();

            if (_document == null)
            {
                Log("[ClashZoneDataService] Document context unavailable for SQLite load.");
                return zones;
            }

            try
            {
                using (var context = new SleeveDbContext(_document, msg => Log($"[ClashZoneDataService][SQLite] {msg}")))
                {
                    var repository = new ClashZoneRepository(context, msg => Log($"[ClashZoneDataService][SQLite] {msg}"));
                    
                    // ✅ DATABASE OPTIMIZATION: Use optimized method for unresolved-only queries
                    // This uses UnresolvedClashZones view for better performance
                    List<ClashZone> dbZones;
                    if (unresolvedOnly)
                    {
                        Log($"[ClashZoneDataService][SQLite] Querying for UNRESOLVED zones only (unresolvedOnly=true)");
                        dbZones = repository.GetClashZonesByFilter(filterName, category, unresolvedOnly: true);
                    }
                    else
                    {
                        Log($"[ClashZoneDataService][SQLite] Querying for ALL zones (unresolvedOnly=false)");
                        dbZones = repository.GetClashZonesByFilter(filterName, category, unresolvedOnly: false) ?? new List<ClashZone>();
                    }
                    
                    Log($"[ClashZoneDataService][SQLite] Query returned {dbZones?.Count ?? 0} zones from database");

                    foreach (var zone in dbZones)
                    {
                        if (zone == null)
                            continue;

                        zone.EnsureSleevePlacementPointReconstructed();
                        zone.EnsureSleevePlacementPointActiveDocumentReconstructed();
                        zones.Add(zone);
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"[ClashZoneDataService] SQLite load failed for filter '{filterName}', category '{category}': {ex.Message}");
            }

            return zones;
        }

        private List<ClashZone> LoadFromXml(string filterName, string category)
        {
            var zones = new List<ClashZone>();

            try
            {
                var filtersDirectory = ProjectPathService.GetFiltersDirectory(_document);
                var normalizedFilter = FilterNameHelper.NormalizeBaseName(filterName, filterName, category);
                var categorySuffix = MepCategoryConstants.GetXmlSuffix(category);
                var expectedFile = Path.Combine(filtersDirectory, $"{normalizedFilter}_{categorySuffix}.xml");

                if (!File.Exists(expectedFile))
                {
                    Log($"[ClashZoneDataService] XML file not found for filter '{filterName}', category '{category}': {expectedFile}");
                    return zones;
                }

                var filterService = new FilterManagementService(_document, msg => Log($"[ClashZoneDataService][FilterMgmt] {msg}"), msg => Log($"[ClashZoneDataService][FilterMgmt] {msg}"));
                var filter = filterService.LoadFilterFromXmlFile(expectedFile);
                var storageZones = filter?.ClashZoneStorage?.AllZones;

                if (storageZones == null || storageZones.Count == 0)
                {
                    Log($"[ClashZoneDataService] XML file '{expectedFile}' contained no clash zones.");
                    return zones;
                }

                foreach (var zone in storageZones)
                {
                    if (zone == null)
                        continue;

                    if (!string.Equals(zone.MepElementCategory, category, StringComparison.OrdinalIgnoreCase))
                        continue;

                    zone.EnsureSleevePlacementPointReconstructed();
                    zone.EnsureSleevePlacementPointActiveDocumentReconstructed();
                    zones.Add(zone);
                }
            }
            catch (Exception ex)
            {
                Log($"[ClashZoneDataService] XML load failed for filter '{filterName}', category '{category}': {ex.Message}");
            }

            return zones;
        }
    }
}
