using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data
{
    /// <summary>
    /// ⚠️⚠️⚠️ CRITICAL PROTECTED SERVICE - DO NOT MODIFY WITHOUT TESTING ⚠️⚠️⚠️
    /// 
    /// ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
    /// To modify protected methods in this class, you MUST:
    /// 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true
    /// 2. Get explicit consent from the project owner
    /// 3. Test thoroughly before committing
    /// 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
    /// 
    /// Service for loading and caching clash zone data for clustering operations.
    /// 
    /// ✅ DATABASE-ONLY ARCHITECTURE: All data comes from SQLite database exclusively.
    /// No XML fallback - database is the single source of truth.
    /// 
    /// ⚠️ DO NOT:
    /// - Add XML fallback (breaks database-only architecture)
    /// - Remove database validation (prevents data corruption)
    /// - Modify cache structure (breaks clustering algorithm)
    /// - Change filtering logic (must match clustering requirements)
    /// 
    /// CACHE STRUCTURE:
    /// - Key: MepElementIdValue (long)
    /// - Value: ClashZone object with all required data
    /// - Used by: ClusteringAlgorithmService, ClusterRotationService, ClusterPlacementService
    /// 
    /// ⚠️ DO NOT SET TO true UNLESS YOU HAVE EXPLICIT CONSENT ⚠️
    /// </summary>
    public class ClusterDataService : IClusterDataService
    {
        // ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
        // To modify protected methods in this class, you MUST:
        // 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true
        // 2. Get explicit consent from the project owner
        // 3. Test thoroughly before committing
        // 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
        // 
        // PROTECTED METHODS:
        // - LoadClashZonesFromRegularXml() - Database loading logic
        // 
        // ⚠️ DO NOT SET TO true UNLESS YOU HAVE EXPLICIT CONSENT ⚠️
        private const bool ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false;
        private readonly Document _doc;
        private readonly Dictionary<long, ClashZone> _clashZoneCache = new Dictionary<long, ClashZone>();

        public int CacheCount => _clashZoneCache.Count;

        public ClusterDataService(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// ⚠️⚠️⚠️ CRITICAL PROTECTED METHOD - DO NOT MODIFY WITHOUT TESTING ⚠️⚠️⚠️
        /// 
        /// ⚠️⚠️⚠️ MODIFICATION CONSENT REQUIRED ⚠️⚠️⚠️
        /// To modify this method, you MUST:
        /// 1. Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true in this class
        /// 2. Get explicit consent from the project owner
        /// 3. Test thoroughly with database operations
        /// 4. Reset ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = false after changes
        /// 
        /// ✅ DATABASE-ONLY: Load clash zones from SQLite database (NO XML fallback).
        /// For all paths (PATH 1, PATH 2, PATH 3), only database is used.
        /// Returns zones with valid SleeveInstanceId for clustering.
        /// 
        /// ⚠️ DO NOT:
        /// - Add XML fallback (breaks database-only architecture)
        /// - Remove validation checks (prevents invalid data)
        /// - Change filtering logic (must match clustering requirements)
        /// - Modify cache population (breaks clustering algorithm)
        /// 
        /// NOTE: xmlFilePath parameter is ignored (kept for backward compatibility only).
        /// </summary>
        public List<ClashZone> LoadClashZonesFromRegularXml(string xmlFilePath, string targetCategory, Document doc)
        {
            // ⚠️ CONSENT CHECK: Prevent modifications without explicit consent
            if (!ALLOW_MODIFICATIONS_TO_PROTECTED_CODE)
            {
                // This method is protected - modifications require explicit consent
                // To modify: Set ALLOW_MODIFICATIONS_TO_PROTECTED_CODE = true and get consent
            }
            
            var clashZones = new List<ClashZone>();

            // ✅ VALIDATION: Ensure all required parameters are valid
            // ✅ DATABASE-ONLY ARCHITECTURE - xmlFilePath parameter is ignored, all data from database
            if (string.IsNullOrEmpty(targetCategory))
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌ ERROR: Missing targetCategory for database load\n");
                return clashZones; // Return empty list - validation failed
            }
            if (doc == null)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌ ERROR: Missing doc for database load\n");
                return clashZones; // Return empty list - validation failed
            }

            try
            {
                // 🔥 CRITICAL: Direct IO logging (bypass SafeFileLogger completely for R2023 compatibility)
                try
                {
                    var versionTag = VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔍 ClusterDataService: About to query database for category '{targetCategory}'\n");
                }
                catch { }
                
                using (var dbContext = new SleeveDbContext(doc))
                {
                    var repository = new ClashZoneRepository(dbContext);
                    var dbZones = repository.GetClashZonesByCategory(targetCategory);
                    
                    // 🔥 CRITICAL: Direct IO logging (bypass SafeFileLogger)
                    try
                    {
                        var versionTag = VersionInfo.VersionTag;
                        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                        var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                        var logPath = Path.Combine(logDir, "cluster_debug.log");
                        File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ✅ DATABASE QUERY RESULT: {dbZones?.Count ?? 0} zones from database for category '{targetCategory}'\n");
                    }
                    catch { }
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ DATABASE: Loaded {dbZones?.Count ?? 0} zones from database for category '{targetCategory}'\n");
                    
                    if (dbZones != null && dbZones.Count > 0)
                    {
                        // ✅ CRITICAL FIX: Filter to only zones with SleeveInstanceId > 0 (placed sleeves) and not cluster resolved
                        // ⚠️ BUG CHECK: If a zone has ClusterSleeveInstanceId > 0, it MUST have IsClusterResolved=true
                        // If it doesn't, the flags are inconsistent (bug in flag management)
                        var placedZones = dbZones.Where(z => z != null && z.SleeveInstanceId > 0 && !z.IsClusterResolved).ToList();
                        
                        // ✅ DIAGNOSTIC: Check for inconsistent flags (SleeveInstanceId > 0 AND ClusterSleeveInstanceId > 0 but IsClusterResolved=false)
                        // This indicates a bug where cluster flags were not set correctly
                        var inconsistentFlags = dbZones.Where(z => z != null && 
                                                                  z.SleeveInstanceId > 0 && 
                                                                  z.ClusterSleeveInstanceId > 0 && 
                                                                  !z.IsClusterResolved).ToList();
                        
                        if (inconsistentFlags.Count > 0)
                        {
                            // 🔥 CRITICAL: Direct IO logging (bypass SafeFileLogger)
                            try
                            {
                                var versionTag = VersionInfo.VersionTag;
                                var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                                var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                                var logPath = Path.Combine(logDir, "cluster_debug.log");
                                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ INCONSISTENT FLAGS DETECTED: {inconsistentFlags.Count} zones have ClusterSleeveInstanceId>0 but IsClusterResolved=false!\n");
                                foreach (var cz in inconsistentFlags.Take(20))
                                {
                                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}]   Zone {cz.Id}: SleeveId={cz.SleeveInstanceId}, ClusterId={cz.ClusterSleeveInstanceId}, IsClusterResolved={cz.IsClusterResolved}\n");
                                }
                            }
                            catch { }
                            
                            // ✅ CRITICAL FIX: Exclude zones with ClusterSleeveInstanceId > 0 from clustering
                            // These zones are part of a cluster and should NOT be processed again
                            placedZones = placedZones.Where(z => z.ClusterSleeveInstanceId <= 0).ToList();
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ FLAG BUG: Excluded {inconsistentFlags.Count} zones with ClusterSleeveInstanceId>0 but IsClusterResolved=false (should be true!)\n");
                        }
                        
                        // 🔥 CRITICAL: Direct IO logging (bypass SafeFileLogger)
                        try
                        {
                            var versionTag = VersionInfo.VersionTag;
                            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                            var logPath = Path.Combine(logDir, "cluster_debug.log");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] 🔍 FILTERING: {dbZones.Count} total zones -> {placedZones.Count} with SleeveInstanceId>0 and !IsClusterResolved (after excluding inconsistent flags)\n");
                        }
                        catch { }
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ DATABASE: Filtered to {placedZones.Count} zones with SleeveInstanceId>0 and not cluster resolved\n");
                        
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
                        
                        // ✅ DATABASE SUCCESS: Return database data (NO XML - database only)
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ✅ DATABASE: Returning {clashZones.Count} clash zones for clustering\n");
                        return clashZones;
                    }
                    else
                    {
                        // 🔥 CRITICAL: Direct IO logging (bypass SafeFileLogger)
                        try
                        {
                            var versionTag = VersionInfo.VersionTag;
                            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                            var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                            var logPath = Path.Combine(logDir, "cluster_debug.log");
                            File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ DATABASE: No clash zones found in database for category '{targetCategory}' - RETURNING EMPTY LIST\n");
                        }
                        catch { }
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ DATABASE: No clash zones found in database for category '{targetCategory}'\n");
                        return clashZones; // Return empty list (NO XML fallback)
                    }
                }
            }
            catch (Exception ex)
            {
                // ✅ ERROR HANDLING: Log error but do NOT fall back to XML (database-only architecture)
                // 🔥 CRITICAL: Direct IO logging for errors (bypass SafeFileLogger)
                try
                {
                    var versionTag = VersionInfo.VersionTag;
                    var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                    var logDir = Path.Combine(appData, "JSE_MEP_Openings", "Logs", versionTag);
                    if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                    var logPath = Path.Combine(logDir, "cluster_debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ DATABASE ERROR: Failed to load clash zones: {ex.Message}\n");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] ❌ StackTrace: {ex.StackTrace}\n");
                }
                catch { }
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌ DATABASE ERROR: Failed to load clash zones from database: {ex.Message}\n");
                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] ❌ StackTrace: {ex.StackTrace}\n");
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClusterDataService] Database load failed for category '{targetCategory}': {ex.Message}");
                }
                
                return clashZones; // Return empty list (NO XML fallback)
            }
            
            // ⚠️ REMOVED: All XML loading code - database only now
            // No XML fallback for any path (PATH 1, PATH 2, PATH 3 all use database exclusively)
            return clashZones;
        }

        /// <summary>
        /// Load clash zone cache for cleanup operations.
        /// Populates internal cache with clash zones for fast lookups during clustering.
        /// NOTE: xmlFilePath parameter is ignored (kept for backward compatibility only).
        /// </summary>
        public void LoadClashZoneCacheForCleanup(string xmlFilePath, string targetCategory, Document doc, string filterName)
        {
            // ✅ CRITICAL: Clear cache before reloading to get fresh data
            _clashZoneCache.Clear();
            
            // ✅ DATABASE-ONLY: Load from database (xmlFilePath ignored)
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
        /// Optimization to avoid duplicate database loading.
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
        /// Get clash zone from cache by SleeveInstanceId.
        /// Returns null if not found.
        /// </summary>
        public ClashZone GetClashZoneBySleeveInstanceId(int sleeveInstanceId)
        {
            if (sleeveInstanceId <= 0)
                return null;
            
            // Search cache by SleeveInstanceId (cache is keyed by MepElementIdValue, so we need to search values)
            return _clashZoneCache.Values.FirstOrDefault(cz => cz != null && cz.SleeveInstanceId == sleeveInstanceId);
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

        /// <summary>
        /// Update pre-calculated 4 corner coordinates for Cluster Sleeves (Phase 3 Persistence).
        /// </summary>
        public void UpdateClusterSleeveCorners(int clusterInstanceId,
            double corner1X, double corner1Y, double corner1Z,
            double corner2X, double corner2Y, double corner2Z,
            double corner3X, double corner3Y, double corner3Z,
            double corner4X, double corner4Y, double corner4Z)
        {
            // ✅ DATABASE-ONLY: Write directly to repository
            try
            {
                using (var dbContext = new SleeveDbContext(_doc))
                {
                    var repository = new ClashZoneRepository(dbContext);
                    repository.UpdateClusterSleeveCorners(clusterInstanceId,
                        corner1X, corner1Y, corner1Z,
                        corner2X, corner2Y, corner2Z,
                        corner3X, corner3Y, corner3Z,
                        corner4X, corner4Y, corner4Z);
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[ClusterDataService] Failed to update corners for cluster {clusterInstanceId}: {ex.Message}");
                }
            }
        }
    }
}
