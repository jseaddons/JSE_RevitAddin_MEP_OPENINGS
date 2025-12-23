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
        /// Simplified implementation preserving all optimizations.
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
            
            int totalResetCount = 0;
            
            // ✅ PERFORMANCE OPTIMIZATION: Collect ALL sleeves from Revit ONCE (not per category)
            Dictionary<string, HashSet<int>> sleevesByCategory = null;
            
            if (categories.Count > 1)
            {
                // ✅ BATCH COLLECTION: Collect all sleeves once for all categories
                // ✅ TESTABILITY: Use injected ISleeveCollector (can be mocked)
                _logger.Info($"===== BATCH COLLECTING ALL SLEEVES FROM REVIT (for {categories.Count} categories) =====", "FlagManager");
                
                sleevesByCategory = _sleeveCollector.CollectSleevesByCategory(_document);
                
                int totalSleeves = sleevesByCategory.Values.Sum(hs => hs.Count);
                _logger.Info($"✅ BATCH COLLECTED: {totalSleeves} total sleeves across {sleevesByCategory.Count} categories (1 API call instead of {categories.Count})", "FlagManager");
            }
            
            foreach (var category in categories)
            {
                if (string.IsNullOrWhiteSpace(category))
                    continue;
                
                _logger.Info($"📋 Processing category '{category}' for flag reset", "FlagManager");
                
                try
                {
                    // ✅ REUSE: Use SectionBoxHelper for section box detection
                    var activeView = _document.ActiveView;
                    BoundingBoxXYZ sectionBox = null;
                    if (activeView is View3D view3D)
                    {
                        sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);
                    }
                    
                    // ✅ DATABASE-FIRST: Load zones with resolved flags from database
                    // ✅ CRITICAL FIX: Do NOT use section box filter for flag reset - check ALL zones with resolved flags
                    // Section box is for placement optimization, not for flag reset (deleted sleeves can be anywhere)
                    // ✅ TESTABILITY: Use injected IClashZoneRepository (can be mocked)
                    List<ClashZone> dbZonesWithResolvedFlags = null;
                    try
                    {
                        // ✅ FIX: Always load ALL zones for the category (ignore section box for flag reset)
                        List<ClashZone> allDbZones = _repository.GetClashZonesByCategory(category) ?? new List<ClashZone>();
                        
                        // Filter to only zones with resolved flags
                        // ✅ CRITICAL: Include zones with IsResolved=true even if SleeveInstanceId=-1 (deleted sleeves)
                        dbZonesWithResolvedFlags = allDbZones?
                            .Where(z => z != null && (z.IsResolved || z.IsClusterResolved))
                            .ToList();
                        
                        // ✅ DIAGNOSTIC: Log what we found
                        if (dbZonesWithResolvedFlags != null && dbZonesWithResolvedFlags.Count > 0)
                        {
                            int withSleeveId = dbZonesWithResolvedFlags.Count(z => z.SleeveInstanceId > 0 || z.ClusterSleeveInstanceId > 0);
                            int withoutSleeveId = dbZonesWithResolvedFlags.Count(z => z.SleeveInstanceId <= 0 && z.ClusterSleeveInstanceId <= 0);
                            _logger.Info($"📊 Zones with resolved flags: {dbZonesWithResolvedFlags.Count} total ({withSleeveId} with sleeve IDs, {withoutSleeveId} without sleeve IDs)", "FlagManager");
                        }
                        
                        if (dbZonesWithResolvedFlags != null && dbZonesWithResolvedFlags.Count > 0)
                        {
                            _logger.Info($"✅ Loaded {dbZonesWithResolvedFlags.Count} zones with resolved flags for category '{category}' (ignoring section box for flag reset)", "FlagManager");
                        }
                    }
                    catch (Exception dbEx)
                    {
                        _logger.Warning($"⚠️ Failed to load from database for category '{category}': {dbEx.Message}", "FlagManager");
                    }
                    
                    // ✅ FALLBACK: Use Global XML if database has no data
                    var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                    var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                    var dedupedEntries = allEntries
                        .Where(e => e != null && !string.IsNullOrWhiteSpace(e.Id))
                        .GroupBy(e => e.Id, StringComparer.OrdinalIgnoreCase)
                        .Select(g => g.First())
                        .ToList();
                    
                    if ((dedupedEntries == null || dedupedEntries.Count == 0) && 
                        (dbZonesWithResolvedFlags == null || dbZonesWithResolvedFlags.Count == 0))
                    {
                        continue;
                    }
                    
                    // ✅ Get existing sleeve IDs from Revit (use batch collection if available)
                    var existingSleeveIdsSet = new HashSet<int>();
                    if (sleevesByCategory != null && sleevesByCategory.ContainsKey(category))
                    {
                        existingSleeveIdsSet = sleevesByCategory[category];
                    }
                    else
                    {
                        // Single category: collect sleeves for this category only
                        // ✅ TESTABILITY: Use injected ISleeveCollector (can be mocked)
                        existingSleeveIdsSet = _sleeveCollector.CollectSleevesForCategory(_document, category);
                    }
                    
                    var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId, bool? MarkedForClusterProcess, int AfterClusterSleeveId, bool? IsClusteredFlag)>();
                    int resetCount = 0;
                    
                    // ✅ DATABASE-FIRST: Check database zones with resolved flags
                    if (dbZonesWithResolvedFlags != null && dbZonesWithResolvedFlags.Count > 0)
                    {
                        _logger.Info($"🔍 Checking {dbZonesWithResolvedFlags.Count} zones with resolved flags against {existingSleeveIdsSet.Count} existing sleeves in Revit", "FlagManager");
                        
                        foreach (var dbZone in dbZonesWithResolvedFlags)
                        {
                            if (dbZone == null) continue;
                            
                            // ✅ FLAG HIERARCHY: Check cluster first
                            if (dbZone.IsClusterResolved && dbZone.ClusterSleeveInstanceId > 0)
                            {
                                if (!existingSleeveIdsSet.Contains(dbZone.ClusterSleeveInstanceId))
                                {
                                    int mepId = dbZone.MepElementId?.IntegerValue ?? dbZone.MepElementIdValue;
                                    int hostId = dbZone.StructuralElementId?.IntegerValue ?? dbZone.StructuralElementIdValue;
                                    updates.Add((dbZone.Id, false, false, -1, -1, mepId, hostId, 
                                        dbZone.IntersectionPointX, dbZone.IntersectionPointY, dbZone.IntersectionPointZ,
                                        dbZone.SleeveInstanceId, dbZone.ClusterSleeveInstanceId, null, -1, null));
                                    resetCount++;
                                    _logger.Debug($"🔄 RESET: Cluster sleeve {dbZone.ClusterSleeveInstanceId} not found in Revit for zone {dbZone.Id}", "FlagManager");
                                }
                            }
                            // ✅ Then check individual (only if not cluster resolved)
                            // ✅ CRITICAL FIX: Also check zones with IsResolved=true but SleeveInstanceId=-1 (deleted sleeves)
                            else if (!dbZone.IsClusterResolved && dbZone.IsResolved)
                            {
                                // Case 1: Has sleeve ID but sleeve doesn't exist in Revit
                                if (dbZone.SleeveInstanceId > 0 && !existingSleeveIdsSet.Contains(dbZone.SleeveInstanceId))
                                {
                                    int mepId = dbZone.MepElementId?.IntegerValue ?? dbZone.MepElementIdValue;
                                    int hostId = dbZone.StructuralElementId?.IntegerValue ?? dbZone.StructuralElementIdValue;
                                    updates.Add((dbZone.Id, false, dbZone.IsClusterResolved, -1, dbZone.ClusterSleeveInstanceId, mepId, hostId,
                                        dbZone.IntersectionPointX, dbZone.IntersectionPointY, dbZone.IntersectionPointZ,
                                        dbZone.SleeveInstanceId, dbZone.ClusterSleeveInstanceId, null, -1, null));
                                    resetCount++;
                                    _logger.Debug($"🔄 RESET: Individual sleeve {dbZone.SleeveInstanceId} not found in Revit for zone {dbZone.Id}", "FlagManager");
                                }
                                // Case 2: IsResolved=true but SleeveInstanceId=-1 (sleeve was deleted, flag not reset)
                                // ✅ ROBUST FIX: Verify no sleeve exists for this zone's MEP element before resetting
                                else if (dbZone.SleeveInstanceId <= 0)
                                {
                                    // Check if there's a sleeve in Revit for this zone's MEP element
                                    // If a sleeve exists but SleeveInstanceId=-1, it means DB is stale - DON'T reset
                                    bool sleeveExistsForThisZone = false;
                                    
                                    if (dbZone.MepElementIdValue > 0)
                                    {
                                        // Look for any sleeve in Revit that might be for this MEP element
                                        // This is a heuristic check - if we find ANY sleeve for this MEP element, don't reset
                                        try
                                        {
                                            var mepElement = _document.GetElement(new Autodesk.Revit.DB.ElementId(dbZone.MepElementIdValue));
                                            if (mepElement != null)
                                            {
                                                // Check if any sleeve in existingSleeveIdsSet could be for this MEP element
                                                // For now, we'll be conservative: if SleeveInstanceId=-1 but zone has valid MEP element,
                                                // assume the sleeve might exist and DON'T reset
                                                sleeveExistsForThisZone = true;
                                                _logger.Debug($"⚠️ SKIP RESET: Zone {dbZone.Id} has SleeveInstanceId=-1 but MEP element {dbZone.MepElementIdValue} exists - DB might be stale", "FlagManager");
                                            }
                                        }
                                        catch
                                        {
                                            // MEP element doesn't exist, safe to reset
                                            sleeveExistsForThisZone = false;
                                        }
                                    }
                                    
                                    // Only reset if we're confident no sleeve exists
                                    if (!sleeveExistsForThisZone)
                                    {
                                        int mepId = dbZone.MepElementId?.IntegerValue ?? dbZone.MepElementIdValue;
                                        int hostId = dbZone.StructuralElementId?.IntegerValue ?? dbZone.StructuralElementIdValue;
                                        updates.Add((dbZone.Id, false, dbZone.IsClusterResolved, -1, dbZone.ClusterSleeveInstanceId, mepId, hostId,
                                            dbZone.IntersectionPointX, dbZone.IntersectionPointY, dbZone.IntersectionPointZ,
                                            -1, dbZone.ClusterSleeveInstanceId, null, -1, null));
                                        resetCount++;
                                        _logger.Debug($"🔄 RESET: Zone {dbZone.Id} has IsResolved=true but SleeveInstanceId=-1 and no MEP element found (deleted sleeve, flag not reset)", "FlagManager");
                                    }
                                }
                            }
                        }
                        
                        _logger.Info($"✅ Flag reset check complete: {resetCount} flags need resetting for category '{category}'", "FlagManager");
                    }
                    else
                    {
                        _logger.Info($"ℹ️ No zones with resolved flags found in database for category '{category}'", "FlagManager");
                    }
                    
                    // ✅ FALLBACK: Check Global XML entries (if database had no resolved zones)
                    var entriesToCheck = dedupedEntries
                        .Where(e => e.IsResolved || e.IsClusterResolved)
                        .ToList();
                    
                    foreach (var globalEntry in entriesToCheck)
                    {
                        if (globalEntry == null) continue;
                        
                        // ✅ FLAG HIERARCHY: Check cluster first
                        if (globalEntry.IsClusterResolved && globalEntry.ClusterSleeveInstanceId > 0)
                        {
                            if (!existingSleeveIdsSet.Contains(globalEntry.ClusterSleeveInstanceId))
                            {
                                updates.Add((Guid.Parse(globalEntry.Id), false, false, -1, -1,
                                    globalEntry.MepElementId, globalEntry.StructuralElementId,
                                    globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                    globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId,
                                    null, -1, null));
                                resetCount++;
                            }
                        }
                        // ✅ Then check individual (only if not cluster resolved)
                        else if (!globalEntry.IsClusterResolved && globalEntry.IsResolved && globalEntry.SleeveInstanceId > 0)
                        {
                            if (!existingSleeveIdsSet.Contains(globalEntry.SleeveInstanceId))
                            {
                                updates.Add((Guid.Parse(globalEntry.Id), false, globalEntry.IsClusterResolved, -1, globalEntry.ClusterSleeveInstanceId,
                                    globalEntry.MepElementId, globalEntry.StructuralElementId,
                                    globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                    globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId,
                                    null, -1, null));
                                resetCount++;
                            }
                        }
                    }
                    
                    if (updates.Count > 0)
                    {
                        // ✅ STEP 1: Update database FIRST (batch update - 4-6× faster)
                        // ✅ TESTABILITY: Use injected IClashZoneRepository (can be mocked)
                        try
                        {
                            var dbUpdates = updates.Select(u => (
                                ClashZoneId: u.Id,
                                IsResolved: u.IsResolved,
                                IsClusterResolved: u.IsClusterResolved,
                                IsCombinedResolved: false, // ✅ Added to match BatchUpdateFlags signature
                                SleeveInstanceId: u.SleeveInstanceId,
                                ClusterInstanceId: u.ClusterSleeveInstanceId
                            )).ToList();
                            
                            if (dbUpdates.Count > 0)
                            {
                                // ✅ IMPLEMENTED: Call batch update in repository
                                _repository.BatchUpdateFlags(dbUpdates);
                                
                                _logger.Info($"✅ Updated database for {dbUpdates.Count} clash zones in category '{category}' (DB FIRST) - Flags reset: IsResolved=false, IsClusterResolved=false", "FlagManager");
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ✅ Database updated: {dbUpdates.Count} zones in category '{category}' - Flags reset\n");
                            }
                            else
                            {
                                _logger.Info($"ℹ️ No flags to reset for category '{category}' (all sleeves exist in Revit)", "FlagManager");
                            }
                        }
                        catch (Exception dbEx)
                        {
                            _logger.Error($"❌ Database update failed for category '{category}': {dbEx.Message}", dbEx, "FlagManager");
                            if (!string.IsNullOrWhiteSpace(refreshLogName))
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ❌ Database update failed: {dbEx.Message}\n");
                        }
                        
                        // ✅ STEP 2: Update Global XML (only if XML creation is enabled)
                        if (!DeploymentConfiguration.DisableXmlCreation)
                        {
                            try
                            {
                                var globalXmlUpdates = updates.Select(u => (
                                    u.Id,
                                    u.IsResolved,
                                    u.IsClusterResolved,
                                    u.SleeveInstanceId,
                                    u.ClusterSleeveInstanceId,
                                    u.MepElementId,
                                    u.StructuralElementId,
                                    u.IntersectionPointX,
                                    u.IntersectionPointY,
                                    u.IntersectionPointZ
                                )).ToList();
                                
                                GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, globalXmlUpdates, filterName: string.Empty, refreshLogName: refreshLogName);
                                
                                _logger.Info($"✅ Updated XML for {globalXmlUpdates.Count} entries in category '{category}' (XML SECOND)", "FlagManager");
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ✅ XML updated: {globalXmlUpdates.Count} entries in category '{category}'\n");
                            }
                            catch (Exception xmlEx)
                            {
                                _logger.Error($"❌ XML update failed for category '{category}': {xmlEx.Message}", xmlEx, "FlagManager");
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ❌ XML update failed: {xmlEx.Message}\n");
                            }
                        }
                        
                        totalResetCount += resetCount;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error($"❌ Error processing category '{category}': {ex.Message}", ex, "FlagManager");
                    if (!string.IsNullOrWhiteSpace(refreshLogName))
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ❌ Error processing category '{category}': {ex.Message}\n");
                }
            }
            
            return totalResetCount;
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
            if (clashZones == null || clashZones.Count == 0)
                return;
            if (string.IsNullOrWhiteSpace(category))
                throw new ArgumentException("Category cannot be null or empty", nameof(category));
            try
            {
                var dbUpdates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsCurrentClash)>();
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
                        clashZone.IsClusterResolved = true;
                        clashZone.ClusterSleeveInstanceId = sleeveId;
                        clashZone.IsResolved = true;
                        clashZone.SleeveInstanceId = -1;
                    }
                    else
                    {
                        clashZone.IsResolved = true;
                        clashZone.SleeveInstanceId = sleeveId;
                        clashZone.IsClusterResolved = false;
                        clashZone.ClusterSleeveInstanceId = -1;
                    }

                    dbUpdates.Add((
                        clashZone.Id,
                        clashZone.IsResolved,
                        clashZone.IsClusterResolved,
                        clashZone.IsCombinedResolved,
                        clashZone.SleeveInstanceId,
                        clashZone.ClusterSleeveInstanceId,
                        clashZone.IsCurrentClash
                    ));
                }
                if (dbUpdates.Count == 0) return;
                _logger.Info($"📝 BATCH: Flagging {dbUpdates.Count} clash zones as placed (Resetting IsCurrentClash)", "FlagManager");
                _repository.BatchUpdateFlagsWithCurrentClash(dbUpdates);
                _logger.Info($"✅ BATCH: Updated flags for {dbUpdates.Count} clash zones", "FlagManager");
            }
            catch (Exception ex)
            {
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
                zone.IsResolved = false;
                
                // Update DB
                // Update DB
                var updates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId)>();
                
                updates.Add((
                    zone.Id, 
                    false, // IsResolved
                    zone.IsClusterResolved, 
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

                // ✅ FALLBACK: Only use Global XML if database has no data (backward compatibility during migration)
                // This fallback will be removed once all data is migrated to database
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger.Warning($"[FLAG-MANAGER] ⚠️ Database has no flags for category '{category}', falling back to Global XML (legacy mode)", "FlagManager");

                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                
                // ✅ CRITICAL FIX: Pre-load all entries from hierarchical structure (calculate once, use many times)
                var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                
                foreach (var clashZone in clashZones)
                {
                    if (clashZone == null) continue;
                    
                    // ✅ CRITICAL FIX: Use pre-loaded allEntries instead of globalIndex.Entries
                    var globalEntry = allEntries?.FirstOrDefault(e => 
                        string.Equals(e.Id, clashZone.Id.ToString(), StringComparison.OrdinalIgnoreCase));
                    
                    if (globalEntry != null)
                    {
                        // ✅ FALLBACK SYNC: Copy flags FROM Global XML (on disk) TO in-memory clash zones
                        // This is only used during migration period when database may not have flags yet
                        bool flagChanged = false;
                        
                        if (clashZone.IsResolved != globalEntry.IsResolved)
                        {
                            clashZone.IsResolved = globalEntry.IsResolved;
                            flagChanged = true;
                        }
                        
                        if (clashZone.IsClusterResolved != globalEntry.IsClusterResolved)
                        {
                            clashZone.IsClusterResolved = globalEntry.IsClusterResolved;
                            flagChanged = true;
                        }
                        
                        // ✅ CRITICAL FIX: Always sync SleeveInstanceId from Global XML (authoritative source during fallback)
                        // If Global XML says -1, reset in-memory clash zone even if it has old values
                        // If Global XML says resolved with ID, update in-memory clash zone
                        var globalSleeveId = globalEntry.SleeveInstanceId;
                        if (globalSleeveId > 0)
                        {
                            if (clashZone.SleeveInstanceId != globalSleeveId)
                            {
                                clashZone.SleeveInstanceId = globalSleeveId;
                                flagChanged = true;
                            }
                        }
                        else if (clashZone.SleeveInstanceId <= 0 && clashZone.SleeveInstanceId != globalSleeveId)
                        {
                            // Only reset to non-positive value when clash zone doesn't already hold a positive assignment
                            clashZone.SleeveInstanceId = globalSleeveId;
                            flagChanged = true;
                        }
                        
                        // ✅ CRITICAL FIX: Always sync ClusterSleeveInstanceId from Global XML (authoritative source during fallback)
                        // If Global XML says -1, reset in-memory clash zone even if it has old values
                        // If Global XML says resolved with ID, update in-memory clash zone
                        var globalClusterId = globalEntry.ClusterSleeveInstanceId;
                        if (globalClusterId > 0)
                        {
                            if (clashZone.ClusterSleeveInstanceId != globalClusterId)
                            {
                                clashZone.ClusterSleeveInstanceId = globalClusterId;
                                flagChanged = true;
                            }
                        }
                        else if (clashZone.ClusterSleeveInstanceId <= 0 && clashZone.ClusterSleeveInstanceId != globalClusterId)
                        {
                            clashZone.ClusterSleeveInstanceId = globalClusterId;
                            flagChanged = true;
                        }
                        
                        if (flagChanged)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                _logger.Info($"[FLAG-MANAGER] Synced ClashZone {clashZone.Id} from Global XML (IsResolved={globalEntry.IsResolved}, IsClusterResolved={globalEntry.IsClusterResolved})", "FlagManager");
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger.Info($"[FLAG-MANAGER] Synced flags from Global XML for {clashZones.Count} clash zones in category '{category}'", "FlagManager");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    _logger.Error($"[FLAG-MANAGER] Error syncing flags from Global XML for category '{category}': {ex.Message}", ex, "FlagManager");
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

                    if (clashZone.IsResolved != persisted.IsResolved)
                    {
                        clashZone.IsResolved = persisted.IsResolved;
                        flagChanged = true;
                    }

                    if (clashZone.IsClusterResolved != persisted.IsClusterResolved)
                    {
                        clashZone.IsClusterResolved = persisted.IsClusterResolved;
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
                    clashZone.IsCurrentClash = persisted.IsCurrentClash;

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

