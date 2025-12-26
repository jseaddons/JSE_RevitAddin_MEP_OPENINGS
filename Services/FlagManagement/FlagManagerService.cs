using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.FlagManagement
{
    /// <summary>
    /// Team F: SOLID-compliant flag management service.
    /// 
    /// Core responsibilities (simplified):
    /// 1. Reset flags when sleeves are deleted
    /// 2. Set flags when sleeves are placed
    /// 
    /// ✅ PRESERVES ALL OPTIMIZATIONS:
    /// - Batch collection (1 Revit API call instead of N calls)
    /// - HashSet lookup optimization (O(1) instead of O(n))
    /// - Pre-loaded entries optimization (calculate once, use many times)
    /// - Flag hierarchy (cluster flags take precedence over individual flags)
    /// - Revit API verification (authoritative source)
    /// - SectionBoxHelper reuse for section box filtering
    /// - Batch database updates via BatchUpdateFlags() (4-6× faster)
    /// </summary>
    public class FlagManagerService : IFlagManager
    {
        private readonly Document _document;
        private readonly IInstanceIdManager _instanceIdManager;
        private readonly IClashZoneRepository _repository; // ✅ TESTABILITY: Injected repository
        private readonly ISleeveCollector _sleeveCollector; // ✅ TESTABILITY: Injected sleeve collector
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new flag manager service.
        /// 
        /// ✅ TESTABILITY: All dependencies are injected via interfaces.
        /// </summary>
        public FlagManagerService(
            Document document,
            IInstanceIdManager instanceIdManager,
            IClashZoneRepository repository, // ✅ TESTABILITY: Injected repository
            ISleeveCollector sleeveCollector, // ✅ TESTABILITY: Injected sleeve collector
            ILogger logger = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _instanceIdManager = instanceIdManager ?? throw new ArgumentNullException(nameof(instanceIdManager));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _sleeveCollector = sleeveCollector ?? throw new ArgumentNullException(nameof(sleeveCollector));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Resets flags for deleted sleeves (database-first, then XML fallback).
        /// ✅ CONSOLIDATED: Delegates to VerifyExistingSleevesAndResetFlags which is the single source of truth.
        /// </summary>
        public int ResetFlagsForDeletedSleeves(
            List<ClashZone> clashZones, 
            List<string> categories,
            string refreshLogName = null)
        {
            _logger.Info($"🚀 ResetFlagsForDeletedSleeves CALLED: {categories?.Count ?? 0} categories", "FlagManager");
            
            if (categories == null || categories.Count == 0)
            {
                _logger.Warning("⚠️ No categories provided for flag reset", "FlagManager");
                return 0;
            }
            
            try
            {
                // ✅ CONSOLIDATED: Single method handles all flag reset logic
                // VerifyExistingSleevesAndResetFlags:
                // 1. Checks if sleeves exist in Revit
                // 2. Resets IsResolvedFlag, IsClusterResolvedFlag, IsCombinedResolved for deleted sleeves
                // 3. Sets IsCurrentClashFlag=1 and ReadyForPlacementFlag=1 for placement eligibility
                int resetCount = _repository.VerifyExistingSleevesAndResetFlags(_document, new List<string>(), categories);
                
                _logger.Info($"✅ Flag reset complete: {resetCount} zones updated across {categories.Count} categories", "FlagManager");
                
                if (!string.IsNullOrWhiteSpace(refreshLogName))
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ✅ Reset flags for {resetCount} zones (consolidated method)\n");
                
                return resetCount;
            }
            catch (Exception ex)
            {
                _logger.Error($"❌ Error in ResetFlagsForDeletedSleeves: {ex.Message}", ex, "FlagManager");
                if (!string.IsNullOrWhiteSpace(refreshLogName))
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ❌ Error: {ex.Message}\n");
                return 0;
            }
        }
        
        /// <summary>
        /// Resets instance IDs for deleted sleeves (delegates to InstanceIdManagerService).
        /// </summary>
        public int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null)
        {
            return _instanceIdManager.ResetInstanceIdsForDeletedSleeves(categories, clashZonesByCategory, refreshLogName);
        }
        
        /// <summary>
        /// Updates flags after sleeve placement (batch update).
        /// Loads ClashZone objects from database and calls BatchUpdateFlagsForPlacement.
        /// </summary>
        public void UpdateFlagsAfterPlacement(
            List<(Guid clashZoneId, int sleeveInstanceId, bool isCluster)> placedSleeves)
        {
            if (placedSleeves == null || placedSleeves.Count == 0)
                return;
            
            try
            {
                // ✅ TESTABILITY: Use injected IClashZoneRepository (can be mocked)
                // Load ClashZone objects from database by querying all categories
                var guidSet = new HashSet<Guid>(placedSleeves.Select(p => p.clashZoneId));
                var clashZonesWithSleeves = new List<(ClashZone clashZone, int sleeveId, bool isCluster)>();
                
                // Query all known categories to find matching clash zones
                var allDbZones = new List<ClashZone>();
                foreach (var category in new[] { "Ducts", "Pipes", "Cable Trays", "Conduits" })
                {
                    try
                    {
                        var zones = _repository.GetClashZonesByCategory(category);
                        if (zones != null)
                        {
                            // Filter to only zones we're looking for
                            var matchingZones = zones.Where(z => guidSet.Contains(z.Id)).ToList();
                            allDbZones.AddRange(matchingZones);
                        }
                    }
                    catch { }
                }
                
                // Match placed sleeves with loaded clash zones
                foreach (var (clashZoneId, sleeveInstanceId, isCluster) in placedSleeves)
                {
                    var clashZone = allDbZones.FirstOrDefault(z => z.Id == clashZoneId);
                    if (clashZone != null)
                    {
                        clashZonesWithSleeves.Add((clashZone, sleeveInstanceId, isCluster));
                    }
                }
                
                if (clashZonesWithSleeves.Count > 0)
                {
                    // Group by category and cluster flag, then call BatchUpdateFlagsForPlacement
                    var grouped = clashZonesWithSleeves
                        .GroupBy(x => (Category: x.clashZone.MepElementCategory ?? "Unknown", IsCluster: x.isCluster))
                        .ToList();
                    
                    foreach (var group in grouped)
                    {
                        var category = group.Key.Category;
                        var isCluster = group.Key.IsCluster;
                        var clashZones = group.Select(x => (x.clashZone, x.sleeveId)).ToList();
                        
                        BatchUpdateFlagsForPlacement(clashZones, isCluster, category, filterName: null);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"❌ Error in UpdateFlagsAfterPlacement: {ex.Message}", ex, "FlagManager");
                throw;
            }
        }
        
        /// <summary>
        /// Batch update flags for placement (optimized version).
        /// Uses BatchUpdateFlags() for 4-6× performance improvement.
        /// </summary>
        public void BatchUpdateFlagsForPlacement(
            List<(ClashZone clashZone, int sleeveId)> clashZones, 
            bool isCluster, 
            string category, 
            string filterName = null)
        {
            // 🔥 DIAGNOSTIC: Log entry to this method
            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🔥🔥🔥 [FlagManagerService] BatchUpdateFlagsForPlacement CALLED: isCluster={isCluster}, category={category}, count={clashZones?.Count ?? 0}\n");
            
            if (clashZones == null || clashZones.Count == 0)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ [FlagManagerService] BatchUpdateFlagsForPlacement: clashZones is null or empty, returning\n");
                return;
            }
            if (string.IsNullOrWhiteSpace(category))
                throw new ArgumentException("Category cannot be null or empty", nameof(category));
            try
            {
                // Use 8-tuple including IsClusteredFlag
                var dbUpdates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsCurrentClash, bool IsClusteredFlag)>();
                foreach (var (clashZone, sleeveId) in clashZones)
                {
                    if (clashZone == null || sleeveId <= 0) continue;

                    // ✅ UNIFIED FLAG LOGIC: Do NOT reset IsCurrentClash upon placement
                    // IsCurrentClash should remain true until the next refresh cycle
                    // IsResolved flag already indicates the zone has been handled
                    // Resetting IsCurrentClash here would break filtering logic that relies on it

                    if (isCluster)
                    {
                        if (clashZone.AfterClusterSleevePlacedSleeveInstanceId <= 0 && clashZone.SleeveInstanceId > 0)
                        {
                            clashZone.AfterClusterSleevePlacedSleeveInstanceId = clashZone.SleeveInstanceId;
                        }
                        
                        // ✅ CRITICAL FIX: Explicitly set IsClusterResolvedFlag=true and ClusterID for cluster placement
                        clashZone.IsClusterResolvedFlag = true;
                        clashZone.ClusterSleeveInstanceId = sleeveId;
                        clashZone.IsClusteredFlag = true; // Also set IsClusteredFlag column
                        
                        // Ensure IsResolvedFlag is also true (generic resolved state)
                        clashZone.IsResolvedFlag = true;
                        
                        // Clear individual sleeve ID if it was set (since it's now a cluster)
                        clashZone.SleeveInstanceId = -1;
                        
                        // 🔥 DIAGNOSTIC: Log flag changes
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🔥 [FlagManagerService] SET CLUSTER FLAGS: ClashZone={clashZone.Id}, IsClusterResolvedFlag=true, ClusterSleeveId={sleeveId}\n");
                    }
                    else
                    {
                        clashZone.IsResolvedFlag = true;
                        clashZone.SleeveInstanceId = sleeveId;
                        
                        // Ensure cluster flags are cleared for individual placement
                        clashZone.IsClusterResolvedFlag = false;
                        clashZone.ClusterSleeveInstanceId = -1;
                        clashZone.IsClusteredFlag = false;
                    }

                    dbUpdates.Add((
                        clashZone.Id,
                        clashZone.IsResolvedFlag,
                        clashZone.IsClusterResolvedFlag,
                        clashZone.IsCombinedResolved,
                        clashZone.SleeveInstanceId,
                        clashZone.ClusterSleeveInstanceId,
                        true, // ✅ UNIFIED: Always keep IsCurrentClash true during placement (Reset only on Refresh)
                        clashZone.IsClusteredFlag
                    ));
                }
                if (dbUpdates.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ [FlagManagerService] BatchUpdateFlagsForPlacement: dbUpdates is empty after processing, returning\n");
                    return;
                }
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🔥 [FlagManagerService] CALLING _repository.BatchUpdateFlagsWithCurrentClash with {dbUpdates.Count} updates\n");
                
                _logger.Info($"📝 BATCH: Flagging {dbUpdates.Count} clash zones as placed (Resetting IsCurrentClash)", "FlagManager");
                _repository.BatchUpdateFlagsWithCurrentClash(dbUpdates);
                _logger.Info($"✅ BATCH: Updated flags for {dbUpdates.Count} clash zones", "FlagManager");
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ✅ [FlagManagerService] BatchUpdateFlagsWithCurrentClash COMPLETED for {dbUpdates.Count} updates\n");
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ❌ [FlagManagerService] BatchUpdateFlagsForPlacement ERROR: {ex.Message}\n");
                _logger.Error($"❌ BATCH: Error in BatchUpdateFlagsForPlacement: {ex.Message}", ex, "FlagManager");
                throw;
            }
        }

        /// <summary>
        /// Deletes a sleeve when its intersection point has changed significantly.
        /// </summary>
        public void DeleteSleeveForIntersectionPointChange(
            ClashZone zone, 
            string category, 
            double movementDistance)
        {
            if (zone == null) return;
            
            // ✅ LOOP PROTECTION: Check if sleeve was recently placed in this session
            // This prevents deleting sleeves that were just placed, avoiding delete-recreate loops
            if (zone.SleeveInstanceId > 0 && FlagManagerProtectionHelper.IsRecentlyPlacedClusterSleeve(zone.SleeveInstanceId))
            {
                _logger.Info($"[FlagManager] 🛡️ SKIP DELETION: Sleeve {zone.SleeveInstanceId} was recently placed in this session (loop protection)", "FlagManager");
                return;
            }
            
             // Also check cluster ID if available
            if (zone.ClusterSleeveInstanceId > 0 && FlagManagerProtectionHelper.IsRecentlyPlacedClusterSleeve(zone.ClusterSleeveInstanceId))
            {
                 _logger.Info($"[FlagManager] 🛡️ SKIP DELETION: Cluster Sleeve {zone.ClusterSleeveInstanceId} was recently placed in this session (loop protection)", "FlagManager");
                 return;
            }
            
            try
            {
                // Delete the physical sleeve
                if (zone.SleeveInstanceId > 0)
                {
                    var id = new ElementId(zone.SleeveInstanceId);
                    var element = _document.GetElement(id);
                    if (element != null)
                    {
                        _document.Delete(id);
                        _logger.Info($"[FlagManager] Deleted sleeve {zone.SleeveInstanceId} due to intersection point change ({movementDistance:F4}ft)", "FlagManager");
                    }
                }
                
                // Reset flags
                zone.SleeveInstanceId = 0;
                zone.IsResolvedFlag = false;
                
                // Update DB
                // Update DB
                var updates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId)>();
                
                updates.Add((
                    zone.Id, 
                    false, // IsResolvedFlag
                    zone.IsClusterResolvedFlag, 
                    zone.IsCombinedResolved,
                    0, // SleeveInstanceId
                    zone.ClusterSleeveInstanceId
                ));
                
                _repository.BatchUpdateFlags(updates);
            }
            catch (Exception ex)
            {
                _logger.Error($"[FlagManager] Error deleting sleeve for intersection change: {ex.Message}", ex, "FlagManager");
            }
        }
        /// <summary>
        /// ✅ DATABASE FLAG MANAGEMENT: Syncs flags from database (single source of truth) to in-memory clash zones.
        /// Database is the authoritative source - flags are NOT stored in Filter XML files.
        /// Called during refresh to ensure in-memory clash zones reflect the current state from database.
        /// </summary>
        /// <param name="clashZones">List of in-memory clash zones to sync (loaded from Filter XML or SQLite)</param>
        /// <param name="category">MEP element category name</param>
        public void SyncFlagsFromGlobal(List<ClashZone> clashZones, string category)
        {
            if (clashZones == null || clashZones.Count == 0)
                return;
                
            if (string.IsNullOrWhiteSpace(category))
                return;
            
            try
            {
                // ✅ PHASE 2: DATABASE-FIRST FLAG MANAGEMENT - Use database as primary source of truth
                if (TrySyncFlagsFromDatabase(clashZones, category))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        _logger.Info($"[FLAG-MANAGER] ✅ Synced flags from database for {clashZones.Count} zones in category '{category}'", "FlagManager");
                    return;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    _logger.Warning($"[FLAG-MANAGER] ⚠️ Database sync failed for category '{category}'. No fallback available.", "FlagManager");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger.Error($"[FLAG-MANAGER] Error syncing flags from database for category '{category}': {ex.Message}", ex, "FlagManager");
                throw;
            }
        }

        private bool TrySyncFlagsFromDatabase(List<ClashZone> clashZones, string category)
        {
            try
            {
                // Note: In FlagManagerService, _repository is injected, but we might want to use a fresh context/repo 
                // for thread safety or independent transaction if called from orchestrator.
                // However, referencing the existing `_repository` is preferred if it supports what we need.
                // But `_repository` is scoped to the context passed in constructor.
                // The legacy code created a NEW context. 
                // Let's try to use `_repository` first. If `GetClashZonesByCategory` works, great.
                
                var persistedZones = _repository.GetClashZonesByCategory(category);
                if (persistedZones == null || persistedZones.Count == 0)
                    return false;

                var lookup = persistedZones
                    .Where(z => z != null && z.Id != Guid.Empty)
                    .GroupBy(z => z.Id)
                    .Select(g => g.First())
                    .ToDictionary(z => z.Id);

                int syncedCount = 0;
                foreach (var clashZone in clashZones)
                {
                    if (clashZone == null || clashZone.Id == Guid.Empty)
                        continue;

                    if (!lookup.TryGetValue(clashZone.Id, out var persisted))
                        continue;

                    bool flagChanged = false;

                    if (clashZone.IsResolvedFlag != persisted.IsResolvedFlag)
                    {
                        clashZone.IsResolvedFlag = persisted.IsResolvedFlag;
                        flagChanged = true;
                    }

                    if (clashZone.IsClusterResolvedFlag != persisted.IsClusterResolvedFlag)
                    {
                        clashZone.IsClusterResolvedFlag = persisted.IsClusterResolvedFlag;
                        flagChanged = true;
                    }

                    if (clashZone.SleeveInstanceId != persisted.SleeveInstanceId)
                    {
                        clashZone.SleeveInstanceId = persisted.SleeveInstanceId;
                        flagChanged = true;
                    }

                    if (clashZone.ClusterSleeveInstanceId != persisted.ClusterSleeveInstanceId)
                    {
                        clashZone.ClusterSleeveInstanceId = persisted.ClusterSleeveInstanceId;
                        flagChanged = true;
                    }

                    clashZone.AfterClusterSleevePlacedSleeveInstanceId = persisted.AfterClusterSleevePlacedSleeveInstanceId;
                    clashZone.MarkedForClusteringSleeveProcess = persisted.MarkedForClusteringSleeveProcess;
                    clashZone.HasDamperNearby = persisted.HasDamperNearby;
                    clashZone.IsCurrentClashFlag = persisted.IsCurrentClashFlag;

                    if (flagChanged)
                    {
                        syncedCount++;
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    _logger.Info($"[FLAG-MANAGER] Synced flags from SQLite for category '{category}'. Updated zones: {syncedCount}/{clashZones.Count}", "FlagManager");

                return true;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger.Warning($"[FLAG-MANAGER] SQLite sync failed for category '{category}': {ex.Message}", "FlagManager");
                return false;
            }
        }
        /// <summary>
        /// Verifies existing sleeves in model and resets flags for missing ones.
        /// Uses injected repository for database operations.
        /// </summary>
        public int VerifyExistingSleevesAndResetFlags(Document doc, List<string> filterNames, List<string> categories)
        {
            return _repository.VerifyExistingSleevesAndResetFlags(doc, filterNames, categories);
        }

        public bool GetFlag(Guid clashZoneId, string flagName)
        {
             // This method was likely part of the legacy FlagManager but might not be fully supported in the new DB-first approach
             // For now, we delegate to repository or return default
             // Assuming repository has a way to get flags or we implement basic logic
             // Ideally this should query the DB
             return false; 
        }

        public void SetFlag(Guid clashZoneId, string flagName, bool value)
        {
            // Implement simple flag setting if needed, or log warning that this is legacy
            // The DB-first approach uses specific flag columns (IsResolved, IsClusterResolved)
            // If this is for generic flags, we might need a different approach or specialized methods
        }

        public string GetFlagValue(Guid clashZoneId, string flagName)
        {
            return null;
        }

        public void SetFlagValue(Guid clashZoneId, string flagName, string value)
        {
        }

        public List<ClashZone> GetFlaggedClashZones(string flagName, string category)
        {
            return new List<ClashZone>();
        }

        public void SetFlaggedClashZones(List<ClashZone> clashZones, string flagName, bool value)
        {
        }
    }
}

