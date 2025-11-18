using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    // ✅ DEPLOYMENT MODE: All logging wrapped - tested pattern for deployment
    // Pattern: if (!DeploymentConfiguration.DeploymentMode) { DebugLogger.Info(...); }
    /// <summary>
    /// Centralized flag management for clash zones.
    /// Eliminates redundant flag operations across multiple services.
    /// Uses GlobalIndexService for all Global XML operations.
    /// </summary>
    public class FlagManager
    {
        private readonly Document _document;
        
        public FlagManager(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }
        
        /// <summary>
        /// ✅ DATABASE-FIRST: Resets instance IDs for deleted sleeves (DB first, then XML).
        /// Sets SleeveInstanceId and ClusterSleeveInstanceId to -1 when sleeves are deleted.
        /// </summary>
        /// <param name="categories">List of categories to check</param>
        /// <param name="clashZonesByCategory">Optional dictionary of clash zones by category</param>
        /// <param name="refreshLogName">Optional refresh log file name</param>
        /// <returns>Number of instance IDs reset</returns>
        public int ResetInstanceIdsForDeletedSleeves(List<string> categories, Dictionary<string, List<ClashZone>> clashZonesByCategory = null, string refreshLogName = null)
        {
            if (categories == null || categories.Count == 0)
                return 0;
            
            int totalResetCount = 0;
            
            foreach (var category in categories)
            {
                if (string.IsNullOrWhiteSpace(category))
                    continue;
                
                try
                {
                    // ✅ PHASE 2: DATABASE-FIRST - Get clash zones with sleeve IDs from database
                    List<ClashZone> dbZones = null;
                    try
                    {
                        using (var context = new SleeveDbContext(_document))
                        {
                            var repository = new ClashZoneRepository(context);
                            dbZones = repository.GetClashZonesByCategory(category)
                                ?.Where(z => z != null && (z.SleeveInstanceId > 0 || z.ClusterSleeveInstanceId > 0))
                                .ToList();
                        }
                    }
                    catch (Exception dbEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Failed to load from database for category '{category}', falling back to Global XML: {dbEx.Message}");
                    }
                    
                    // ✅ FALLBACK: Use Global XML if database has no data
                    List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ)> allEntries = null;
                    
                    if (dbZones != null && dbZones.Count > 0)
                    {
                        // ✅ Use database data
                        allEntries = dbZones.Select(z => (
                            z.Id,
                            z.IsResolved,
                            z.IsClusterResolved,
                            z.SleeveInstanceId,
                            z.ClusterSleeveInstanceId,
                            z.MepElementId?.IntegerValue ?? z.MepElementIdValue,
                            z.StructuralElementId?.IntegerValue ?? z.StructuralElementIdValue,
                            z.IntersectionPointX,
                            z.IntersectionPointY,
                            z.IntersectionPointZ
                        )).ToList();
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] ✅ [INSTANCE-ID-RESET] Using database for category '{category}' ({allEntries.Count} zones with sleeve IDs)");
                    }
                    else
                    {
                        // ✅ FALLBACK: Use Global XML
                        var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                        var globalEntries = GlobalIndexService.GetAllEntries(globalIndex)
                            .Where(e => e != null && !string.IsNullOrWhiteSpace(e.Id) && (e.SleeveInstanceId > 0 || e.ClusterSleeveInstanceId > 0))
                            .ToList();
                        
                        if (globalEntries != null && globalEntries.Count > 0)
                        {
                            allEntries = globalEntries.Select(e => (
                                Guid.Parse(e.Id),
                                e.IsResolved,
                                e.IsClusterResolved,
                                e.SleeveInstanceId,
                                e.ClusterSleeveInstanceId,
                                e.MepElementId,
                                e.StructuralElementId,
                                e.IntersectionPointX,
                                e.IntersectionPointY,
                                e.IntersectionPointZ
                            )).ToList();
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ [INSTANCE-ID-RESET] Using Global XML fallback for category '{category}' ({allEntries.Count} entries)");
                        }
                    }
                    
                    if (allEntries == null || allEntries.Count == 0)
                        continue;
                    
                    // Collect sleeve IDs to check
                    var sleeveIdsToCheck = new HashSet<int>();
                    foreach (var entry in allEntries)
                    {
                        if (entry.SleeveInstanceId > 0)
                            sleeveIdsToCheck.Add(entry.SleeveInstanceId);
                        if (entry.ClusterSleeveInstanceId > 0)
                            sleeveIdsToCheck.Add(entry.ClusterSleeveInstanceId);
                    }
                    
                    // Check which sleeves exist in Revit
                    var existingSleeveIds = new HashSet<int>();
                    foreach (var sleeveId in sleeveIdsToCheck)
                    {
                        try
                        {
                            var element = _document.GetElement(new ElementId(sleeveId));
                            if (element != null && element is FamilyInstance sleeve)
                            {
                                bool isSleeve = (sleeve.Category?.Name == "Generic Models" || sleeve.Category?.Name == "Structural Connections") &&
                                                (sleeve.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                                 sleeve.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true);
                                if (isSleeve)
                                    existingSleeveIds.Add(sleeveId);
                            }
                        }
                        catch
                        {
                            // Sleeve doesn't exist
                        }
                    }
                    
                    // Build updates for deleted sleeves
                    var updates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId)>();
                    
                    foreach (var entry in allEntries)
                    {
                        bool needsUpdate = false;
                        int newSleeveId = entry.SleeveInstanceId;
                        int newClusterId = entry.ClusterSleeveInstanceId;
                        
                        // Check cluster sleeve
                        if (entry.ClusterSleeveInstanceId > 0 && !existingSleeveIds.Contains(entry.ClusterSleeveInstanceId))
                        {
                            newClusterId = -1;
                            needsUpdate = true;
                        }
                        
                        // Check individual sleeve (only if cluster is not resolved)
                        if (!entry.IsClusterResolved && entry.SleeveInstanceId > 0 && !existingSleeveIds.Contains(entry.SleeveInstanceId))
                        {
                            newSleeveId = -1;
                            needsUpdate = true;
                        }
                        
                        if (needsUpdate)
                        {
                            updates.Add((
                                entry.Id,
                                entry.IsResolved && newSleeveId > 0,
                                entry.IsClusterResolved && newClusterId > 0,
                                newSleeveId,
                                newClusterId,
                                entry.MepElementId,
                                entry.StructuralElementId,
                                entry.IntersectionPointX,
                                entry.IntersectionPointY,
                                entry.IntersectionPointZ,
                                entry.SleeveInstanceId,
                                entry.ClusterSleeveInstanceId
                            ));
                            totalResetCount++;
                        }
                    }
                    
                    if (updates.Count > 0)
                    {
                        // ✅ STEP 1: Update database FIRST
                        try
                        {
                            using (var context = new SleeveDbContext(_document, msg =>
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER][INSTANCE-ID-RESET][SQLite] {msg}");
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET][SQLite] {msg}\n");
                            }))
                            {
                                var repository = new ClashZoneRepository(context, msg =>
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER][INSTANCE-ID-RESET][SQLite] {msg}");
                                    if (!string.IsNullOrWhiteSpace(refreshLogName))
                                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET][SQLite] {msg}\n");
                                });
                                
                                var dbUpdates = updates.Select(u => (
                                    ClashZoneId: u.ClashZoneId,
                                    IsResolved: u.IsResolved,
                                    IsClusterResolved: u.IsClusterResolved,
                                    SleeveInstanceId: u.SleeveInstanceId,
                                    ClusterInstanceId: u.ClusterSleeveInstanceId,
                                    MepElementId: u.MepElementId,
                                    StructuralElementId: u.StructuralElementId,
                                    IntersectionPointX: u.IntersectionPointX,
                                    IntersectionPointY: u.IntersectionPointY,
                                    IntersectionPointZ: u.IntersectionPointZ,
                                    OldSleeveInstanceId: u.OldSleeveInstanceId,
                                    OldClusterInstanceId: u.OldClusterInstanceId
                                )).ToList();
                                
                                repository.BatchUpdateFlags(dbUpdates);
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] ✅ [INSTANCE-ID-RESET] Updated database for {dbUpdates.Count} clash zones in category '{category}' (DB FIRST)");
                                
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ✅ Database updated: {dbUpdates.Count} zones in category '{category}'\n");
                            }
                        }
                        catch (Exception dbEx)
                        {
                            DebugLogger.Error($"[FLAG-MANAGER] ❌ [INSTANCE-ID-RESET] Database update failed for category '{category}': {dbEx.Message}");
                            if (!string.IsNullOrWhiteSpace(refreshLogName))
                                SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ❌ Database update failed: {dbEx.Message}\n");
                            // Continue to XML update even if DB fails
                        }
                        
                        // ✅ PHASE 2: Update XML only if XML creation is enabled (for backward compatibility)
                        if (!DeploymentConfiguration.DisableXmlCreation)
                        {
                            try
                            {
                                var globalXmlUpdates = updates.Select(u => (
                                    u.ClashZoneId,
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
                                
                                // ✅ PHASE 2: Only update Global XML if XML creation is enabled
                                if (!DeploymentConfiguration.DisableXmlCreation)
                                {
                                    GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, globalXmlUpdates, filterName: null, refreshLogName: refreshLogName);
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✅ [INSTANCE-ID-RESET] Updated XML for {globalXmlUpdates.Count} entries in category '{category}' (XML SECOND)");
                                }
                                else
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ⚠️ [INSTANCE-ID-RESET] XML creation disabled - skipping Global XML update for {globalXmlUpdates.Count} entries in category '{category}' (database only mode)");
                                }
                                
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ✅ XML updated: {globalXmlUpdates.Count} entries in category '{category}'\n");
                            }
                            catch (Exception xmlEx)
                            {
                                DebugLogger.Error($"[FLAG-MANAGER] ❌ [INSTANCE-ID-RESET] XML update failed for category '{category}': {xmlEx.Message}");
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ❌ XML update failed: {xmlEx.Message}\n");
                                // Don't throw - XML is optional in Phase 2
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ⚠️ [INSTANCE-ID-RESET] Skipping XML update (XML creation disabled) for category '{category}'");
                        }
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Error($"[FLAG-MANAGER] ❌ [INSTANCE-ID-RESET] Error processing category '{category}': {ex.Message}");
                    if (!string.IsNullOrWhiteSpace(refreshLogName))
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ❌ Error processing category '{category}': {ex.Message}\n");
                }
            }
            
            return totalResetCount;
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
                        DebugLogger.Info($"[FLAG-MANAGER] ✅ Synced flags from database for {clashZones.Count} zones in category '{category}'");
                    return;
                }

                // ✅ FALLBACK: Only use Global XML if database has no data (backward compatibility during migration)
                // This fallback will be removed once all data is migrated to database
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Database has no flags for category '{category}', falling back to Global XML (legacy mode)");

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
                                DebugLogger.Info($"[FLAG-MANAGER] Synced ClashZone {clashZone.Id} from Global XML (IsResolved={globalEntry.IsResolved}, IsClusterResolved={globalEntry.IsClusterResolved})");
                        }
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Synced flags from Global XML for {clashZones.Count} clash zones in category '{category}'");
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error syncing flags from Global XML for category '{category}': {ex.Message}");
                throw;
            }
        }

        private bool TrySyncFlagsFromDatabase(List<ClashZone> clashZones, string category)
        {
            try
            {
                using (var context = new SleeveDbContext(_document, msg =>
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER][SQLite] {msg}");
                }))
                {
                    var repository = new ClashZoneRepository(context, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER][SQLite] {msg}");
                    });

                    var persistedZones = repository.GetClashZonesByCategory(category);
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
                        DebugLogger.Info($"[FLAG-MANAGER] Synced flags from SQLite for category '{category}'. Updated zones: {syncedCount}/{clashZones.Count}");

                    return true;
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[FLAG-MANAGER] SQLite sync failed for category '{category}': {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// ✅ UNIFIED RESET METHOD: Resets flags for deleted sleeves across all categories.
        /// Single method handles everything - loads Global XML entries, checks Revit, updates Global XML, optionally syncs Filter XML.
        /// 
        /// WORKING LOGIC:
        /// 1. For each category, loads ALL Global XML entries (hierarchical + flat)
        /// 2. Pre-collects sleeves once per category (optimized HashSet for O(1) lookup)
        /// 3. Checks Global XML entries against Revit (Revit is authoritative)
        /// 4. Follows flag hierarchy: Cluster flags checked FIRST, then individual flags
        /// 5. Updates Global XML entries with reset flags
        /// 6. Optionally syncs back to Filter XML clash zones if provided
        /// </summary>
        /// <param name="categories">List of categories to check</param>
        /// <param name="clashZonesByCategory">Optional dictionary of Filter XML clash zones to sync back (category -> clash zones)</param>
        /// <returns>Total number of flags reset</returns>
        public int ResetFlagsForDeletedSleeves(List<string> categories, Dictionary<string, List<ClashZone>> clashZonesByCategory = null, string refreshLogName = null)
        {
            // ✅ CRITICAL: UNMISTAKABLE MARKER - This confirms the NEW code is running
            // ✅ VERSION MARKER: v2.0 - Enhanced logging with build timestamps and verification
            // ✅ ALWAYS LOG TO REFRESH FILE FIRST (bypasses DebugLogger filtering)
            if (!string.IsNullOrWhiteSpace(refreshLogName))
            {
                try
                {
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ═══════════════════════════════════════════════════════════════════════════════\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ⭐⭐⭐ NEW CODE VERSION v2.0 IS RUNNING ⭐⭐⭐\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ✅ Enhanced logging with build timestamps and verification enabled\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ✅ Database-first flag updates with detailed tracking\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] ✅ Global XML update tracking with entry lookup logging\n");
                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] ═══════════════════════════════════════════════════════════════════════════════\n");
                }
                catch (Exception fileEx)
                {
                    // If file logging fails, try DebugLogger as fallback
                    try { DebugLogger.Error($"[FLAG-MANAGER] Error writing version marker to file: {fileEx.Message}"); } catch { }
                }
            }
            
            // ✅ ALSO LOG VIA DebugLogger (may be filtered in deployment mode, but file logging above ensures it's captured)
            try
            {
                DebugLogger.Info("═══════════════════════════════════════════════════════════════════════════════");
                DebugLogger.Info("[FLAG-MANAGER] ⭐⭐⭐ NEW CODE VERSION v2.0 IS RUNNING ⭐⭐⭐");
                DebugLogger.Info("[FLAG-MANAGER] ✅ Enhanced logging with build timestamps and verification enabled");
                DebugLogger.Info("[FLAG-MANAGER] ✅ Database-first flag updates with detailed tracking");
                DebugLogger.Info("[FLAG-MANAGER] ✅ Global XML update tracking with entry lookup logging");
                DebugLogger.Info("═══════════════════════════════════════════════════════════════════════════════");
            }
            catch (Exception markerEx)
            {
                // Even if logging fails, try to log the error
                try { DebugLogger.Error($"[FLAG-MANAGER] Error logging version marker: {markerEx.Message}"); } catch { }
            }
            
            if (categories == null || categories.Count == 0)
                return 0;
            
            int totalResetCount = 0;
            
            try
            {
                // ✅ BUILD TIMESTAMP: Always log build timestamp (even in deployment mode for troubleshooting)
                try
                {
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    string buildTimestamp = assembly.GetName().Version?.ToString() ?? "Unknown";
                    DateTime buildTime = DateTime.MinValue;
                    string location = assembly.Location;
                    if (!string.IsNullOrEmpty(location) && System.IO.File.Exists(location))
                    {
                        buildTime = System.IO.File.GetLastWriteTime(location);
                    }
                    else
                    {
                        // Fallback: Use current time if location unavailable
                        buildTime = DateTime.Now;
                    }
                    DebugLogger.Info($"[FLAG-MANAGER] ===== BUILD INFO: Version={buildTimestamp}, BuildTime={buildTime:yyyy-MM-dd HH:mm:ss}, Location={location ?? "N/A"}, Method=ResetFlagsForDeletedSleeves =====");
                }
                catch (Exception buildEx)
                {
                    DebugLogger.Info($"[FLAG-MANAGER] ===== BUILD INFO: Error getting build info: {buildEx.Message}, Method=ResetFlagsForDeletedSleeves =====");
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] ===== UNIFIED RESET: CHECKING FOR DELETED SLEEVES FOR {categories.Count} CATEGORIES =====");
                
                foreach (var category in categories)
                {
                    if (string.IsNullOrWhiteSpace(category))
                        continue;
                    
                    try
                    {
                        var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                        
                        // ✅ CRITICAL FIX: Use GetAllEntries to get entries from BOTH hierarchical and flat structures
                        var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                        
                        if (allEntries == null || allEntries.Count == 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] No Global XML entries found for category '{category}' - skipping reset check");
                            continue;
                        }
                        
                        // ✅ DEBUG: Log how many entries have resolved flags
                        var dedupedEntries = allEntries
                            .Where(e => e != null && !string.IsNullOrWhiteSpace(e.Id))
                            .GroupBy(e => e.Id, StringComparer.OrdinalIgnoreCase)
                            .Select(g => g.First())
                            .ToList();

                        int resolvedEntries = dedupedEntries.Count(e => e.IsResolved || e.IsClusterResolved);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] Found {allEntries.Count} total Global XML entries for category '{category}', {resolvedEntries} have resolved flags");
                        
                        // ✅ DATABASE-FIRST: Get sleeve IDs from database, fall back to Global XML only if needed
                        var sleeveIdsToCheck = new HashSet<int>();
                        bool useDatabase = false;
                        
                        try
                        {
                            using (var context = new SleeveDbContext(_document, msg =>
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER][DB] {msg}");
                            }))
                            {
                                var repository = new ClashZoneRepository(context, msg =>
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER][DB] {msg}");
                                });
                                
                                var dbZones = repository.GetClashZonesByCategory(category);
                                if (dbZones != null && dbZones.Count > 0)
                                {
                                    // ✅ DATABASE-FIRST: Get sleeve IDs from database
                                    foreach (var zone in dbZones)
                                    {
                                        if (zone.SleeveInstanceId > 0)
                                            sleeveIdsToCheck.Add(zone.SleeveInstanceId);
                                        if (zone.ClusterSleeveInstanceId > 0)
                                            sleeveIdsToCheck.Add(zone.ClusterSleeveInstanceId);
                                    }
                                    useDatabase = true;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        DebugLogger.Info($"[FLAG-MANAGER] ✅ DATABASE-FIRST: Found {sleeveIdsToCheck.Count} sleeve IDs from database for category '{category}'");
                                    }
                                }
                            }
                        }
                        catch (Exception dbEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Database lookup failed (non-blocking), falling back to Global XML: {dbEx.Message}");
                        }
                        
                        // ✅ FALLBACK: Use Global XML only if database didn't provide sleeve IDs
                        if (!useDatabase || sleeveIdsToCheck.Count == 0)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[FLAG-MANAGER] ⚠️ FALLBACK: Getting sleeve IDs from Global XML (database had {sleeveIdsToCheck.Count} IDs)");
                            }
                            
                            foreach (var entry in dedupedEntries)
                            {
                                if (entry.SleeveInstanceId > 0)
                                    sleeveIdsToCheck.Add(entry.SleeveInstanceId);
                                if (entry.ClusterSleeveInstanceId > 0)
                                    sleeveIdsToCheck.Add(entry.ClusterSleeveInstanceId);
                            }
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[FLAG-MANAGER] FALLBACK: Found {sleeveIdsToCheck.Count} sleeve IDs from Global XML for category '{category}'");
                            }
                        }
                        
                        // ✅ EFFICIENT: Use doc.GetElement() for O(1) lookup instead of scanning all sleeves
                        var existingSleeveIdsSet = new HashSet<int>();
                        var sleeveIdToCategory = new Dictionary<int, string>();
                        var guidToSleeveId = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        
                        foreach (var sleeveId in sleeveIdsToCheck)
                        {
                            try
                            {
                                // ✅ EFFICIENT: Direct ElementId lookup (O(1)) instead of scanning all elements
                                var element = _document.GetElement(new ElementId(sleeveId));
                                if (element != null && element is FamilyInstance sleeve)
                                {
                                    // Verify it's actually a sleeve (check category and family name)
                                    bool isSleeve = (sleeve.Category?.Name == "Generic Models" || sleeve.Category?.Name == "Structural Connections") &&
                                                    (sleeve.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                                     sleeve.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true ||
                                                     sleeve.Symbol?.FamilyName?.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) == true ||
                                                     sleeve.Symbol?.FamilyName?.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase) == true);
                                    
                                    if (isSleeve)
                                    {
                                        existingSleeveIdsSet.Add(sleeveId);
                                        
                                        string sleeveCategory = GetSleeveCategory(sleeve);
                                        if (!string.IsNullOrWhiteSpace(sleeveCategory))
                                        {
                                            sleeveCategory = sleeveCategory.Trim();
                                            sleeveIdToCategory[sleeveId] = sleeveCategory;
                                        }
                                        
                                        string clashGuid = GetClashZoneGuidValue(sleeve);
                                        if (!string.IsNullOrWhiteSpace(clashGuid) && !guidToSleeveId.ContainsKey(clashGuid))
                                        {
                                            guidToSleeveId[clashGuid] = sleeveId;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                // Element doesn't exist or error accessing it - sleeve was deleted
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] Sleeve ID {sleeveId} not found in Revit (likely deleted): {ex.Message}");
                            }
                        }
                        
                        // ✅ FALLBACK: Still collect all sleeves for GUID discovery (needed for TryResolveSleeveIdFromGuid)
                        // But only if we need to discover sleeves by GUID
                        var allSleeveInstances = new List<FamilyInstance>();
                        if (dedupedEntries.Any(e => (e.IsResolved && e.SleeveInstanceId <= 0) || (e.IsClusterResolved && e.ClusterSleeveInstanceId <= 0)))
                        {
                            // Only scan if we need GUID discovery (sleeve IDs are missing)
                            allSleeveInstances = new FilteredElementCollector(_document)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .Where(s =>
                            {
                                bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                                        s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                                string familyName = s.Symbol?.FamilyName ?? string.Empty;
                                bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                                     familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                                return (s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") &&
                                       (hasSleeveKeyword || isKnownFamily);
                            })
                            .ToList();

                            // Build GUID lookup from all sleeves (for discovery)
                        foreach (var sleeve in allSleeveInstances)
                        {
                                string clashGuid = GetClashZoneGuidValue(sleeve);
                                if (!string.IsNullOrWhiteSpace(clashGuid) && !guidToSleeveId.ContainsKey(clashGuid))
                                {
                                    guidToSleeveId[clashGuid] = sleeve.Id.IntegerValue;
                                    
                            string sleeveCategory = GetSleeveCategory(sleeve);
                            if (!string.IsNullOrWhiteSpace(sleeveCategory))
                            {
                                        sleeveIdToCategory[sleeve.Id.IntegerValue] = sleeveCategory.Trim();
                                }
                            }
                            }
                        }

                        void LogToRefresh(string message)
                        {
                            // ✅ CRITICAL: Always log to DebugLogger (even if refreshLogName is null)
                            DebugLogger.Info($"[FLAG-MANAGER] {message}");

                            // ✅ Also log to refresh log file if provided
                            if (!string.IsNullOrWhiteSpace(refreshLogName))
                            {
                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [FLAG-MANAGER] {message}\n");
                            }
                        }

                        if (existingSleeveIdsSet.Count > 0)
                        {
                            var sample = string.Join(", ", existingSleeveIdsSet.Take(10));
                            LogToRefresh($"Revit sleeve scan → Category='{category}', Count={existingSleeveIdsSet.Count}, Sample=[{sample}{(existingSleeveIdsSet.Count > 10 ? ", ..." : string.Empty)}]");
                        }
                        else
                        {
                            LogToRefresh($"Revit sleeve scan → Category='{category}', Count=0 (no sleeves detected in model for this category)");
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[FLAG-MANAGER] Pre-collected {existingSleeveIdsSet.Count} sleeve IDs for category '{category}' (O(1) lookup)");
                            if (existingSleeveIdsSet.Count > 0)
                            {
                                DebugLogger.Info($"[FLAG-MANAGER] Sleeve IDs found in Revit for category '{category}': {string.Join(", ", existingSleeveIdsSet.Take(10))}{(existingSleeveIdsSet.Count > 10 ? "..." : "")}");
                            }
                            else
                            {
                                // ✅ DEBUG: Log why no sleeves were found
                                int sleevesCheckedCount = allSleeveInstances.Count > 0 ? allSleeveInstances.Count : existingSleeveIdsSet.Count;
                                DebugLogger.Info($"[FLAG-MANAGER] Found {sleevesCheckedCount} sleeves checked (direct lookup: {existingSleeveIdsSet.Count}, full scan: {allSleeveInstances.Count})");
                                if (allSleeveInstances.Count > 0)
                                {
                                    var categoryBreakdown = allSleeveInstances
                                        .GroupBy(s => GetSleeveCategory(s) ?? "NO_MEP_CATEGORY")
                                        .ToDictionary(g => g.Key, g => g.Count());
                                    
                                    var breakdown = string.Join(", ", categoryBreakdown.Select(kvp => $"{kvp.Key}={kvp.Value}"));
                                    DebugLogger.Info($"[FLAG-MANAGER] Sleeve category breakdown: {breakdown}");
                                    LogToRefresh($"Sleeve category breakdown (all sleeves detected): {breakdown}");
                                }
                            }
                        }
                        
                        var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId)>();
                        int resetCount = 0;
                        
                        // ✅ DEBUG: Log which entries will be checked
                        var entriesToCheck = dedupedEntries
                            .Where(e => e != null && (e.IsResolved || e.IsClusterResolved))
                            .ToList();
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[FLAG-MANAGER] Checking {entriesToCheck.Count} entries with resolved flags:");
                            foreach (var entry in entriesToCheck.Take(5))
                            {
                                DebugLogger.Info($"[FLAG-MANAGER]   Entry {entry.Id}: ClusterResolved={entry.IsClusterResolved} (ID={entry.ClusterSleeveInstanceId}), Resolved={entry.IsResolved} (ID={entry.SleeveInstanceId})");
                            }
                            if (entriesToCheck.Count > 5)
                                DebugLogger.Info($"[FLAG-MANAGER]   ... and {entriesToCheck.Count - 5} more entries");
                        }
                        
                        int detailLogBudget = 30;

                        // ✅ CRITICAL DEBUG: Log all sleeve IDs that will be checked
                        DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': Will check {entriesToCheck.Count} entries with resolved flags");
                        var allSleeveIdsToCheck = new HashSet<int>();
                        foreach (var entry in entriesToCheck)
                        {
                            if (entry.ClusterSleeveInstanceId > 0)
                                allSleeveIdsToCheck.Add(entry.ClusterSleeveInstanceId);
                            if (entry.SleeveInstanceId > 0)
                                allSleeveIdsToCheck.Add(entry.SleeveInstanceId);
                        }
                        DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': Checking {allSleeveIdsToCheck.Count} unique sleeve IDs: {string.Join(", ", allSleeveIdsToCheck.Take(20))}{(allSleeveIdsToCheck.Count > 20 ? "..." : "")}");
                        DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': Sleeve IDs found in Revit: {string.Join(", ", existingSleeveIdsSet.Take(20))}{(existingSleeveIdsSet.Count > 20 ? "..." : "")}");

                        // Check ALL Global XML entries (even if not in Filter XML)
                        foreach (var globalEntry in dedupedEntries)
                        {
                            if (globalEntry == null) continue;
                            
                            // Only check entries with resolved flags (no need to check unresolved ones)
                            if (!globalEntry.IsResolved && !globalEntry.IsClusterResolved)
                                continue;
                            
                            // ✅ PROTECTED LOGIC: Flag Hierarchy - Check cluster FIRST (cluster flags take precedence)
                            if (globalEntry.IsClusterResolved && globalEntry.ClusterSleeveInstanceId > 0)
                            {
                                int clusterSleeveIdToCheck = globalEntry.ClusterSleeveInstanceId;
                                
                                // ✅ DATABASE VERIFICATION: Check if sleeve ID exists in database first
                                bool sleeveExistsInDb = false;
                                try
                                {
                                    using (var context = new SleeveDbContext(_document, msg =>
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[FLAG-MANAGER][DB-CHECK] {msg}");
                                    }))
                                    {
                                        var repository = new ClashZoneRepository(context, msg =>
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[FLAG-MANAGER][DB-CHECK] {msg}");
                                        });
                                        
                                        // Check if this sleeve ID exists in database for this category
                                        var dbZones = repository.GetClashZonesByCategory(category);
                                        sleeveExistsInDb = dbZones?.Any(z => z.ClusterSleeveInstanceId == clusterSleeveIdToCheck || z.SleeveInstanceId == clusterSleeveIdToCheck) == true;
                                        
                                        DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Cluster sleeve ID {clusterSleeveIdToCheck} - EXISTS IN DATABASE: {sleeveExistsInDb}");
                                    }
                                }
                                catch (Exception dbEx)
                                {
                                    DebugLogger.Warning($"[FLAG-MANAGER] Entry {globalEntry.Id}: Database check failed for cluster sleeve ID {clusterSleeveIdToCheck}: {dbEx.Message}");
                                }
                                
                                // Check Revit
                                bool clusterSleeveExists = existingSleeveIdsSet.Contains(clusterSleeveIdToCheck);
                                
                                // ✅ CRITICAL: Also verify the sleeve is actually a cluster sleeve (not just any sleeve)
                                // A cluster sleeve should have the cluster sleeve parameter set
                                bool isActualClusterSleeve = false;
                                if (clusterSleeveExists)
                                {
                                    try
                                    {
                                        var clusterElement = _document.GetElement(new ElementId(clusterSleeveIdToCheck));
                                        if (clusterElement is FamilyInstance clusterSleeve)
                                        {
                                            // Check if this sleeve has cluster metadata (MEP Element IDs parameter)
                                            var mepElementIdsParam = clusterSleeve.LookupParameter("MEP Element IDs");
                                            if (mepElementIdsParam != null && mepElementIdsParam.HasValue)
                                            {
                                                var mepElementIdsValue = mepElementIdsParam.AsString();
                                                // Cluster sleeves should have multiple MEP element IDs (comma-separated)
                                                isActualClusterSleeve = !string.IsNullOrWhiteSpace(mepElementIdsValue) && mepElementIdsValue.Contains(",");
                                            }
                                            
                                            if (!isActualClusterSleeve && !DeploymentConfiguration.DeploymentMode)
                                            {
                                                DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Entry {globalEntry.Id}: Sleeve ID {clusterSleeveIdToCheck} exists but is NOT a cluster sleeve (no cluster metadata) - resetting cluster flag");
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Warning($"[FLAG-MANAGER] Error verifying cluster sleeve {clusterSleeveIdToCheck}: {ex.Message}");
                                    }
                                }
                                
                                // ✅ CRITICAL DEBUG: Log complete sleeve existence check (DB + Revit + Cluster verification)
                                DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Checking cluster sleeve ID {clusterSleeveIdToCheck} - DB={sleeveExistsInDb}, Revit={clusterSleeveExists}, IsClusterSleeve={isActualClusterSleeve}, Final={clusterSleeveExists && isActualClusterSleeve}");
                                
                                // ✅ CRITICAL: If sleeve exists in DB but NOT in Revit, it was deleted - reset flags
                                // ✅ CRITICAL: If sleeve exists in Revit but is NOT a cluster sleeve, reset cluster flag
                                if (sleeveExistsInDb && !clusterSleeveExists)
                                {
                                    DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Entry {globalEntry.Id}: Cluster sleeve ID {clusterSleeveIdToCheck} EXISTS IN DATABASE but NOT IN REVIT - sleeve was deleted, resetting flags");
                                }
                                
                                if (clusterSleeveExists && !isActualClusterSleeve)
                                {
                                    DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Entry {globalEntry.Id}: Sleeve ID {clusterSleeveIdToCheck} exists but is NOT a cluster sleeve - resetting cluster flag (keeping individual flag if valid)");
                                    // Reset cluster flag but preserve individual flag if individual sleeve exists
                                    bool individualSleeveExists = globalEntry.SleeveInstanceId > 0 && existingSleeveIdsSet.Contains(globalEntry.SleeveInstanceId);
                                    updates.Add((Guid.Parse(globalEntry.Id), individualSleeveExists ? globalEntry.IsResolved : false, false,
                                                 individualSleeveExists ? globalEntry.SleeveInstanceId : -1, -1,
                                                 globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                 globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                 globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    resetCount++;
                                    LogToRefresh($"RESET cluster flag for entry {globalEntry.Id}: Sleeve {clusterSleeveIdToCheck} is not a cluster sleeve → IsClusterResolved=false");
                                    continue; // Skip individual check (already processed)
                                }
                                
                                // Update final check result
                                clusterSleeveExists = clusterSleeveExists && isActualClusterSleeve;

                                if (!clusterSleeveExists &&
                                    TryResolveSleeveIdFromGuid(globalEntry.Id, guidToSleeveId, sleeveIdToCategory, category, out var discoveredClusterId, out var discoveredClusterCategory))
                                {
                                    clusterSleeveExists = discoveredClusterId > 0;
                                    if (clusterSleeveExists)
                                    {
                                        existingSleeveIdsSet.Add(discoveredClusterId);
                                        DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Cluster sleeve ID {clusterSleeveIdToCheck} NOT found, but discovered via GUID: {discoveredClusterId}");

                                        if (clusterSleeveIdToCheck != discoveredClusterId)
                                        {
                                            updates.Add((Guid.Parse(globalEntry.Id), globalEntry.IsResolved, true,
                                                         globalEntry.SleeveInstanceId, discoveredClusterId,
                                                         globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                         globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                         globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                            LogToRefresh($"Entry {globalEntry.Id}: Cluster sleeve healed via GUID → {discoveredClusterId} (category='{discoveredClusterCategory ?? "UNKNOWN"}').");
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Cluster sleeve ID healed via GUID → {discoveredClusterId} (was {clusterSleeveIdToCheck})");
                                            clusterSleeveIdToCheck = discoveredClusterId;
                                        }
                                    }
                                }
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER]   Entry {globalEntry.Id}: Cluster sleeve {clusterSleeveIdToCheck} exists in Revit: {clusterSleeveExists}");
                                if (detailLogBudget-- > 0)
                                    LogToRefresh($"Entry {globalEntry.Id}: ClusterSleeveId={clusterSleeveIdToCheck}, Exists={clusterSleeveExists}, IsClusterResolved={globalEntry.IsClusterResolved} (BEFORE check)");
                                
                                if (!clusterSleeveExists)
                                {
                                    // ✅ CRITICAL: Log BEFORE and AFTER states for clarity
                                    DebugLogger.Info($"[FLAG-MANAGER] ✓✓✓ RESETTING Entry {globalEntry.Id}: Cluster sleeve {clusterSleeveIdToCheck} NOT FOUND in Revit");
                                    DebugLogger.Info($"[FLAG-MANAGER]   BEFORE: IsClusterResolved={globalEntry.IsClusterResolved}, IsResolved={globalEntry.IsResolved}, ClusterSleeveId={globalEntry.ClusterSleeveInstanceId}, SleeveId={globalEntry.SleeveInstanceId}");
                                    DebugLogger.Info($"[FLAG-MANAGER]   AFTER:  IsClusterResolved=false, IsResolved=false, ClusterSleeveId=-1, SleeveId=-1");
                                    
                                    updates.Add((Guid.Parse(globalEntry.Id), false, false, -1, -1,
                                                globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    resetCount++;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset ALL flags for Global XML entry {globalEntry.Id} - cluster sleeve {clusterSleeveIdToCheck} NOT FOUND");
                                    LogToRefresh($"RESET cluster entry {globalEntry.Id}: ClusterSleeveId={clusterSleeveIdToCheck} missing → IsClusterResolved: {globalEntry.IsClusterResolved}→false, IsResolved: {globalEntry.IsResolved}→false");
                                }
                                continue; // Skip individual check if cluster was checked
                            }
                            else if (globalEntry.IsClusterResolved && globalEntry.ClusterSleeveInstanceId <= 0)
                            {
                                // ✅ CRITICAL: IsClusterResolved=true but ClusterSleeveInstanceId is missing/invalid
                                // Try to discover via GUID, but if not found, reset cluster flag
                                bool discovered = TryResolveSleeveIdFromGuid(globalEntry.Id, guidToSleeveId, sleeveIdToCategory, category, out var discoveredClusterId, out var discoveredClusterCategory) &&
                                                 discoveredClusterId > 0;
                                
                                if (discovered)
                                {
                                    existingSleeveIdsSet.Add(discoveredClusterId);
                                    updates.Add((Guid.Parse(globalEntry.Id), globalEntry.IsResolved, true,
                                                 globalEntry.SleeveInstanceId, discoveredClusterId,
                                                 globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                 globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                 globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    LogToRefresh($"Entry {globalEntry.Id}: Cluster sleeve ID populated via GUID → {discoveredClusterId} (category='{discoveredClusterCategory ?? "UNKNOWN"}').");
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Cluster sleeve ID populated via GUID → {discoveredClusterId} (was missing)");
                                }
                                else
                                {
                                    // ✅ CRITICAL FIX: Cluster flag is true but ID is missing and can't be discovered → Reset cluster flag
                                    DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Entry {globalEntry.Id}: IsClusterResolved=true but ClusterSleeveInstanceId={globalEntry.ClusterSleeveInstanceId} and cannot be discovered via GUID → Resetting cluster flag");
                                    updates.Add((Guid.Parse(globalEntry.Id), globalEntry.IsResolved, false, globalEntry.SleeveInstanceId, -1,
                                                 globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                 globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                 globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    resetCount++;
                                    LogToRefresh($"RESET cluster flag for entry {globalEntry.Id}: ClusterSleeveInstanceId missing and not discoverable → IsClusterResolved=false");
                                }
                                continue; // Skip individual check (cluster was processed)
                            }
                            
                            // ✅ PROTECTED LOGIC: Then check individual (ONLY if cluster flag is false)
                            // ⚠️ CRITICAL: If IsClusterResolved=true, individual sleeves were deleted during clustering
                            // So we should NOT check individual sleeves if cluster is resolved
                            if (!globalEntry.IsClusterResolved && globalEntry.IsResolved && globalEntry.SleeveInstanceId > 0)
                            {
                                int sleeveIdToCheck = globalEntry.SleeveInstanceId;
                                
                                // ✅ DATABASE VERIFICATION: Check if sleeve ID exists in database first
                                bool sleeveExistsInDb = false;
                                try
                                {
                                    using (var context = new SleeveDbContext(_document, msg =>
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[FLAG-MANAGER][DB-CHECK] {msg}");
                                    }))
                                    {
                                        var repository = new ClashZoneRepository(context, msg =>
                                        {
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[FLAG-MANAGER][DB-CHECK] {msg}");
                                        });
                                        
                                        // Check if this sleeve ID exists in database for this category
                                        var dbZones = repository.GetClashZonesByCategory(category);
                                        sleeveExistsInDb = dbZones?.Any(z => z.SleeveInstanceId == sleeveIdToCheck || z.ClusterSleeveInstanceId == sleeveIdToCheck) == true;
                                        
                                        DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Individual sleeve ID {sleeveIdToCheck} - EXISTS IN DATABASE: {sleeveExistsInDb}");
                                    }
                                }
                                catch (Exception dbEx)
                                {
                                    DebugLogger.Warning($"[FLAG-MANAGER] Entry {globalEntry.Id}: Database check failed for individual sleeve ID {sleeveIdToCheck}: {dbEx.Message}");
                                }
                                
                                // Check Revit
                                bool individualSleeveExists = existingSleeveIdsSet.Contains(sleeveIdToCheck);
                                
                                // ✅ CRITICAL DEBUG: Log complete sleeve existence check (DB + Revit)
                                DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Checking individual sleeve ID {sleeveIdToCheck} - DB={sleeveExistsInDb}, Revit={individualSleeveExists}, Final={individualSleeveExists}");
                                
                                // ✅ CRITICAL: If sleeve exists in DB but NOT in Revit, it was deleted - reset flags
                                if (sleeveExistsInDb && !individualSleeveExists)
                                {
                                    DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Entry {globalEntry.Id}: Individual sleeve ID {sleeveIdToCheck} EXISTS IN DATABASE but NOT IN REVIT - sleeve was deleted, resetting flags");
                                }

                                if (!individualSleeveExists &&
                                    TryResolveSleeveIdFromGuid(globalEntry.Id, guidToSleeveId, sleeveIdToCategory, category, out var discoveredSleeveId, out var discoveredCategory))
                                {
                                    individualSleeveExists = discoveredSleeveId > 0;
                                    if (individualSleeveExists)
                                    {
                                        existingSleeveIdsSet.Add(discoveredSleeveId);
                                        DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Individual sleeve ID {sleeveIdToCheck} NOT found, but discovered via GUID: {discoveredSleeveId}");

                                        if (sleeveIdToCheck != discoveredSleeveId)
                                        {
                                            updates.Add((Guid.Parse(globalEntry.Id), true, globalEntry.IsClusterResolved, discoveredSleeveId, globalEntry.ClusterSleeveInstanceId,
                                                        globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                        globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                        globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                            LogToRefresh($"Entry {globalEntry.Id}: Individual sleeve healed via GUID → {discoveredSleeveId} (category='{discoveredCategory ?? "UNKNOWN"}').");
                                            if (!DeploymentConfiguration.DeploymentMode)
                                                DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Sleeve ID healed via GUID → {discoveredSleeveId} (was {sleeveIdToCheck})");
                                            sleeveIdToCheck = discoveredSleeveId;
                                        }
                                    }
                                }
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER]   Entry {globalEntry.Id}: Individual sleeve {sleeveIdToCheck} exists in Revit: {individualSleeveExists}");
                                if (detailLogBudget-- > 0)
                                    LogToRefresh($"Entry {globalEntry.Id}: SleeveId={sleeveIdToCheck}, Exists={individualSleeveExists}, IsResolved={globalEntry.IsResolved} (BEFORE check)");
                                
                                if (!individualSleeveExists)
                                {
                                    // ✅ EDGE CASE: Check if individual sleeve was deleted because it fell in a cluster zone
                                    // Look for cluster sleeves at/near this location
                                    int foundClusterSleeveId = -1;
                                    const double clusterProximityTolerance = 1.0; // 1 foot tolerance for cluster zone detection
                                    
                                    // Check other Global XML entries with same MEP+Structural+nearby point that have IsClusterResolved=true
                                    var nearbyClusterEntry = dedupedEntries
                                        .Where(e => e != null && 
                                                    e.Id != globalEntry.Id &&
                                                    e.IsClusterResolved &&
                                                    e.ClusterSleeveInstanceId > 0 &&
                                                    e.MepElementId == globalEntry.MepElementId &&
                                                    e.StructuralElementId == globalEntry.StructuralElementId)
                                        .Where(e =>
                                        {
                                            double dx = e.IntersectionPointX - globalEntry.IntersectionPointX;
                                            double dy = e.IntersectionPointY - globalEntry.IntersectionPointY;
                                            double dz = e.IntersectionPointZ - globalEntry.IntersectionPointZ;
                                            double distance = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                                            return distance <= clusterProximityTolerance;
                                        })
                                        .FirstOrDefault();
                                    
                                    if (nearbyClusterEntry != null)
                                    {
                                        foundClusterSleeveId = nearbyClusterEntry.ClusterSleeveInstanceId;
                                        // Verify cluster sleeve exists in Revit
                                        if (existingSleeveIdsSet.Contains(foundClusterSleeveId))
                                        {
                                            // ✅ EDGE CASE: Individual sleeve was deleted because it fell in cluster zone
                                            DebugLogger.Info($"[FLAG-MANAGER] ✅ EDGE CASE: Entry {globalEntry.Id}: Individual sleeve {sleeveIdToCheck} NOT FOUND, but cluster sleeve {foundClusterSleeveId} exists nearby");
                                            DebugLogger.Info($"[FLAG-MANAGER]   BEFORE: IsClusterResolved={globalEntry.IsClusterResolved}, IsResolved={globalEntry.IsResolved}, ClusterSleeveId={globalEntry.ClusterSleeveInstanceId}, SleeveId={globalEntry.SleeveInstanceId}");
                                            DebugLogger.Info($"[FLAG-MANAGER]   AFTER:  IsClusterResolved=true, IsResolved=false, ClusterSleeveId={foundClusterSleeveId}, SleeveId=-1");
                                            updates.Add((Guid.Parse(globalEntry.Id), false, true, -1, foundClusterSleeveId,
                                                        globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                        globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                        globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                            resetCount++;
                                            LogToRefresh($"EDGE CASE: Entry {globalEntry.Id}: Individual sleeve deleted (fell in cluster zone) → IsClusterResolved: {globalEntry.IsClusterResolved}→true, ClusterSleeveInstanceId={foundClusterSleeveId}");
                                            continue; // Skip to next entry
                                        }
                                    }
                                    
                                    // No cluster sleeve found nearby - individual was deleted manually
                                    DebugLogger.Info($"[FLAG-MANAGER] ✓✓✓ RESETTING Entry {globalEntry.Id}: Individual sleeve {sleeveIdToCheck} NOT FOUND in Revit");
                                    DebugLogger.Info($"[FLAG-MANAGER]   BEFORE: IsClusterResolved={globalEntry.IsClusterResolved}, IsResolved={globalEntry.IsResolved}, ClusterSleeveId={globalEntry.ClusterSleeveInstanceId}, SleeveId={globalEntry.SleeveInstanceId}");
                                    DebugLogger.Info($"[FLAG-MANAGER]   AFTER:  IsClusterResolved={globalEntry.IsClusterResolved}, IsResolved=false, ClusterSleeveId={globalEntry.ClusterSleeveInstanceId}, SleeveId=-1");
                                    updates.Add((Guid.Parse(globalEntry.Id), false, globalEntry.IsClusterResolved, -1, globalEntry.ClusterSleeveInstanceId,
                                                globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    resetCount++;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset individual flag for Global XML entry {globalEntry.Id} - individual sleeve {sleeveIdToCheck} NOT FOUND");
                                    LogToRefresh($"RESET individual entry {globalEntry.Id}: SleeveId={sleeveIdToCheck} missing → IsResolved: {globalEntry.IsResolved}→false, SleeveInstanceId=-1");
                                }
                            }
                            else if (!globalEntry.IsClusterResolved && globalEntry.IsResolved && globalEntry.SleeveInstanceId <= 0)
                            {
                                // ✅ CRITICAL: IsResolved=true but SleeveInstanceId is missing/invalid
                                // Try to discover via GUID, but if not found, reset individual flag
                                bool discovered = TryResolveSleeveIdFromGuid(globalEntry.Id, guidToSleeveId, sleeveIdToCategory, category, out var discoveredSleeveId, out var discoveredCategory) &&
                                                 discoveredSleeveId > 0;
                                
                                if (discovered)
                                {
                                    existingSleeveIdsSet.Add(discoveredSleeveId);
                                    updates.Add((Guid.Parse(globalEntry.Id), true, globalEntry.IsClusterResolved, discoveredSleeveId, globalEntry.ClusterSleeveInstanceId,
                                                globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    LogToRefresh($"Entry {globalEntry.Id}: Sleeve ID populated via GUID → {discoveredSleeveId} (category='{discoveredCategory ?? "UNKNOWN"}').");
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] Entry {globalEntry.Id}: Sleeve ID populated via GUID → {discoveredSleeveId} (was missing)");
                                }
                                else
                                {
                                    // ✅ CRITICAL FIX: Individual flag is true but ID is missing and can't be discovered → Reset individual flag
                                    DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Entry {globalEntry.Id}: IsResolved=true but SleeveInstanceId={globalEntry.SleeveInstanceId} and cannot be discovered via GUID → Resetting individual flag");
                                    updates.Add((Guid.Parse(globalEntry.Id), false, globalEntry.IsClusterResolved, -1, globalEntry.ClusterSleeveInstanceId,
                                                globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    resetCount++;
                                    LogToRefresh($"RESET individual flag for entry {globalEntry.Id}: SleeveInstanceId missing and not discoverable → IsResolved=false");
                                }
                            }
                        }
                        
                        // ✅ CRITICAL DEBUG: Verify updates list before persisting
                        // ✅ ALWAYS LOG TO FILE (bypasses DebugLogger filtering)
                        LogToRefresh($"===== VERIFICATION: Category '{category}' scan complete =====");
                        LogToRefresh($"updates.Count={updates.Count}, resetCount={resetCount}");
                        DebugLogger.Info($"[FLAG-MANAGER] ===== VERIFICATION: Category '{category}' scan complete =====");
                        DebugLogger.Info($"[FLAG-MANAGER] updates.Count={updates.Count}, resetCount={resetCount}");
                        
                        if (updates.Count != resetCount)
                        {
                            LogToRefresh($"❌ BUG DETECTED: updates.Count ({updates.Count}) != resetCount ({resetCount})! This indicates a logic error.");
                            DebugLogger.Error($"[FLAG-MANAGER] ❌ BUG DETECTED: updates.Count ({updates.Count}) != resetCount ({resetCount})! This indicates a logic error.");
                        }

                        if (updates.Count > 0)
                        {
                            LogToRefresh($"===== ABOUT TO PERSIST {updates.Count} UPDATES =====");
                            LogToRefresh($"First update: GUID={updates[0].Id}, IsResolved={updates[0].IsResolved}, IsClusterResolved={updates[0].IsClusterResolved}, SleeveId={updates[0].SleeveInstanceId}, ClusterId={updates[0].ClusterSleeveInstanceId}");
                            DebugLogger.Info($"[FLAG-MANAGER] ===== ABOUT TO PERSIST {updates.Count} UPDATES =====");
                            DebugLogger.Info($"[FLAG-MANAGER] First update: GUID={updates[0].Id}, IsResolved={updates[0].IsResolved}, IsClusterResolved={updates[0].IsClusterResolved}, SleeveId={updates[0].SleeveInstanceId}, ClusterId={updates[0].ClusterSleeveInstanceId}");
                            if (updates.Count > 1)
                            {
                                LogToRefresh($"Last update: GUID={updates[updates.Count - 1].Id}, IsResolved={updates[updates.Count - 1].IsResolved}, IsClusterResolved={updates[updates.Count - 1].IsClusterResolved}, SleeveId={updates[updates.Count - 1].SleeveInstanceId}, ClusterId={updates[updates.Count - 1].ClusterSleeveInstanceId}");
                                DebugLogger.Info($"[FLAG-MANAGER] Last update: GUID={updates[updates.Count - 1].Id}, IsResolved={updates[updates.Count - 1].IsResolved}, IsClusterResolved={updates[updates.Count - 1].IsClusterResolved}, SleeveId={updates[updates.Count - 1].SleeveInstanceId}, ClusterId={updates[updates.Count - 1].ClusterSleeveInstanceId}");
                            }
                            
                            // Log sample of updates for verification
                            for (int i = 0; i < Math.Min(5, updates.Count); i++)
                            {
                                var update = updates[i];
                                LogToRefresh($"Update #{i + 1}/{updates.Count}: GUID={update.Id}, IsResolved={update.IsResolved}, IsClusterResolved={update.IsClusterResolved}, SleeveId={update.SleeveInstanceId}, ClusterId={update.ClusterSleeveInstanceId}");
                                DebugLogger.Info($"[FLAG-MANAGER] Update #{i + 1}/{updates.Count}: GUID={update.Id}, IsResolved={update.IsResolved}, IsClusterResolved={update.IsClusterResolved}, SleeveId={update.SleeveInstanceId}, ClusterId={update.ClusterSleeveInstanceId}");
                            }
                        }
                        else if (resetCount > 0)
                        {
                            LogToRefresh($"❌ CRITICAL BUG: updates list is EMPTY but resetCount={resetCount}! Updates were detected but not added to list.");
                            DebugLogger.Error($"[FLAG-MANAGER] ❌ CRITICAL BUG: updates list is EMPTY but resetCount={resetCount}! Updates were detected but not added to list.");
                        }
                        else
                        {
                            LogToRefresh($"No updates needed - all sleeves exist or no resolved flags found.");
                            DebugLogger.Info($"[FLAG-MANAGER] No updates needed - all sleeves exist or no resolved flags found.");
                        }
                        
                        // ✅ OPTION 4: DATABASE-FIRST APPROACH - Update database first, then sync to Global XML
                        // ✅ CRITICAL: Always log this summary (even if refreshLogName is null)
                        var summaryMsg = $"Category '{category}' scan complete → updates={updates.Count}, resetCount={resetCount}";
                        DebugLogger.Info($"[FLAG-MANAGER] {summaryMsg}");
                        LogToRefresh(summaryMsg);
                        
                        // ✅ CRITICAL DEBUG: Always log these details (even in deployment mode for troubleshooting)
                        DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': Found {updates.Count} updates, {resetCount} flags to reset");
                        DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': Checked {dedupedEntries.Count} Global XML entries, {entriesToCheck.Count} had resolved flags");
                        DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': Found {existingSleeveIdsSet.Count} sleeves in Revit for this category");
                        int totalSleevesScanned = allSleeveInstances.Count > 0 ? allSleeveInstances.Count : existingSleeveIdsSet.Count;
                        DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': Sleeves checked (direct lookup: {existingSleeveIdsSet.Count}, full scan: {allSleeveInstances.Count})");
                        
                        // ✅ ENHANCED DEBUG: Log all sleeve IDs from Global XML that should be checked
                        var allGlobalSleeveIds = new HashSet<int>();
                        foreach (var entry in entriesToCheck)
                        {
                            if (entry.SleeveInstanceId > 0) allGlobalSleeveIds.Add(entry.SleeveInstanceId);
                            if (entry.ClusterSleeveInstanceId > 0) allGlobalSleeveIds.Add(entry.ClusterSleeveInstanceId);
                        }
                        DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': Global XML has {allGlobalSleeveIds.Count} unique sleeve IDs that should exist: {string.Join(", ", allGlobalSleeveIds.Take(20))}{(allGlobalSleeveIds.Count > 20 ? "..." : "")}");
                        
                        // ✅ ENHANCED DEBUG: Find which sleeve IDs from Global XML are missing in Revit
                        var missingSleeveIds = allGlobalSleeveIds.Where(id => !existingSleeveIdsSet.Contains(id)).ToList();
                        if (missingSleeveIds.Count > 0)
                        {
                            DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Category '{category}': {missingSleeveIds.Count} sleeve IDs from Global XML are MISSING in Revit: {string.Join(", ", missingSleeveIds.Take(20))}{(missingSleeveIds.Count > 20 ? "..." : "")}");
                        }
                        else if (allGlobalSleeveIds.Count > 0)
                        {
                            DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': All {allGlobalSleeveIds.Count} sleeve IDs from Global XML exist in Revit - no reset needed");
                        }
                        
                        if (updates.Count > 0)
                        {
                            var sampleUpdate = updates.First();
                            DebugLogger.Info($"[FLAG-MANAGER] Sample update: GUID={sampleUpdate.Id}, IsResolved={sampleUpdate.IsResolved}, IsClusterResolved={sampleUpdate.IsClusterResolved}, SleeveId={sampleUpdate.SleeveInstanceId}, ClusterId={sampleUpdate.ClusterSleeveInstanceId}");
                        }
                        else if (resetCount > 0)
                        {
                            DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Category '{category}': resetCount={resetCount} but updates.Count=0 - this shouldn't happen!");
                        }
                        else if (entriesToCheck.Count > 0 && missingSleeveIds.Count == 0)
                        {
                            DebugLogger.Info($"[FLAG-MANAGER] Category '{category}': All {entriesToCheck.Count} entries with resolved flags have valid sleeves in Revit - no reset needed");
                        }
                        else if (entriesToCheck.Count > 0 && missingSleeveIds.Count > 0)
                        {
                            DebugLogger.Error($"[FLAG-MANAGER] ❌ Category '{category}': BUG DETECTED - {missingSleeveIds.Count} sleeves are missing but updates.Count=0! This indicates a logic error in flag reset detection.");
                        }

                        if (updates.Count > 0)
                        {
                            // ✅ CRITICAL DEBUG: Log ALL updates before persisting (always log, even in deployment mode)
                            LogToRefresh($"===== PERSISTING {updates.Count} UPDATES FOR CATEGORY '{category}' =====");
                            DebugLogger.Info($"[FLAG-MANAGER] ===== PERSISTING {updates.Count} UPDATES FOR CATEGORY '{category}' =====");
                            for (int i = 0; i < Math.Min(10, updates.Count); i++)
                            {
                                var update = updates[i];
                                LogToRefresh($"Update #{i + 1}/{updates.Count}: GUID={update.Id}, IsResolved={update.IsResolved}, IsClusterResolved={update.IsClusterResolved}, SleeveId={update.SleeveInstanceId}, ClusterId={update.ClusterSleeveInstanceId}, MEP={update.MepElementId}, Host={update.StructuralElementId}, Point=({update.IntersectionPointX:F3}, {update.IntersectionPointY:F3}, {update.IntersectionPointZ:F3})");
                                DebugLogger.Info($"[FLAG-MANAGER] Update #{i + 1}/{updates.Count}: GUID={update.Id}, IsResolved={update.IsResolved}, IsClusterResolved={update.IsClusterResolved}, SleeveId={update.SleeveInstanceId}, ClusterId={update.ClusterSleeveInstanceId}, MEP={update.MepElementId}, Host={update.StructuralElementId}, Point=({update.IntersectionPointX:F3}, {update.IntersectionPointY:F3}, {update.IntersectionPointZ:F3})");
                            }
                            if (updates.Count > 10)
                            {
                                LogToRefresh($"... and {updates.Count - 10} more updates");
                            }
                            
                            // Step 1: Update database first (triggers will auto-compute SleeveState)
                            try
                            {
                                LogToRefresh($"===== DATABASE UPDATE: Starting database update for {updates.Count} clash zones =====");
                                using (var context = new SleeveDbContext(_document, msg =>
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER][SQLite] {msg}");
                                    // ✅ ALWAYS LOG TO FILE (bypasses DebugLogger filtering)
                                    if (!string.IsNullOrWhiteSpace(refreshLogName))
                                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER][SQLite] {msg}\n");
                                }))
                                {
                                    var repository = new ClashZoneRepository(context, msg =>
                                    {
                                        if (!DeploymentConfiguration.DeploymentMode)
                                            DebugLogger.Info($"[FLAG-MANAGER][SQLite] {msg}");
                                        // ✅ ALWAYS LOG TO FILE (bypasses DebugLogger filtering)
                                        if (!string.IsNullOrWhiteSpace(refreshLogName))
                                            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER][SQLite] {msg}\n");
                                    });
                                    
                                    // Convert updates to database format (include OLD SleeveInstanceId/ClusterInstanceId for direct matching)
                                    // ✅ CRITICAL: Match exact field names expected by BatchUpdateFlags
                                    var dbUpdates = updates.Select(u => (
                                        ClashZoneId: u.Id,
                                        IsResolved: u.IsResolved,
                                        IsClusterResolved: u.IsClusterResolved,
                                        SleeveInstanceId: u.SleeveInstanceId,
                                        ClusterInstanceId: u.ClusterSleeveInstanceId, // Note: BatchUpdateFlags uses ClusterInstanceId, not ClusterSleeveInstanceId
                                        MepElementId: u.MepElementId,
                                        StructuralElementId: u.StructuralElementId,
                                        IntersectionPointX: u.IntersectionPointX,
                                        IntersectionPointY: u.IntersectionPointY,
                                        IntersectionPointZ: u.IntersectionPointZ,
                                        OldSleeveInstanceId: u.OldSleeveInstanceId,
                                        OldClusterInstanceId: u.OldClusterInstanceId
                                    )).ToList();
                                    
                                    LogToRefresh($"✅ DATABASE UPDATE: Calling BatchUpdateFlags with {dbUpdates.Count} updates");
                                    DebugLogger.Info($"[FLAG-MANAGER] ✅ DATABASE UPDATE: Calling BatchUpdateFlags with {dbUpdates.Count} updates");
                                    repository.BatchUpdateFlags(dbUpdates);
                                    LogToRefresh($"✅ DATABASE UPDATE: BatchUpdateFlags completed successfully for {dbUpdates.Count} clash zones in category '{category}'");
                                    DebugLogger.Info($"[FLAG-MANAGER] ✅ DATABASE UPDATE: BatchUpdateFlags completed successfully for {dbUpdates.Count} clash zones in category '{category}'");
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✅ Updated database flags for {dbUpdates.Count} clash zones in category '{category}' (database-first)");
                                }
                            }
                            catch (Exception dbEx)
                            {
                                // ✅ NON-BLOCKING: Log database error but continue with Global XML sync
                                LogToRefresh($"❌ DATABASE UPDATE FAILED for category '{category}': {dbEx.Message}");
                                LogToRefresh($"❌ DATABASE UPDATE Stack trace: {dbEx.StackTrace}");
                                DebugLogger.Error($"[FLAG-MANAGER] ❌ DATABASE UPDATE FAILED for category '{category}': {dbEx.Message}");
                                DebugLogger.Error($"[FLAG-MANAGER] ❌ DATABASE UPDATE Stack trace: {dbEx.StackTrace}");
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Failed to update database flags for category '{category}': {dbEx.Message}");
                            }
                            
                            // Step 2: Sync to Global XML (for backward compatibility and cross-filter tracking)
                            // ✅ CRITICAL FIX: FilterName is preserved from Global XML entry (not overwritten)
                            try
                            {
                                // ✅ ENHANCED DEBUG: Log what we're about to update
                                LogToRefresh($"===== GLOBAL XML UPDATE: About to update Global XML for category '{category}' with {updates.Count} updates =====");
                                DebugLogger.Info($"[FLAG-MANAGER] ===== GLOBAL XML UPDATE: About to update Global XML for category '{category}' with {updates.Count} updates =====");
                                foreach (var update in updates.Take(5))
                                {
                                    LogToRefresh($"  Update: GUID={update.Id}, IsResolved={update.IsResolved}, IsClusterResolved={update.IsClusterResolved}, SleeveId={update.SleeveInstanceId}, ClusterId={update.ClusterSleeveInstanceId}");
                                    DebugLogger.Info($"[FLAG-MANAGER]   Update: GUID={update.Id}, IsResolved={update.IsResolved}, IsClusterResolved={update.IsClusterResolved}, SleeveId={update.SleeveInstanceId}, ClusterId={update.ClusterSleeveInstanceId}");
                                }
                                if (updates.Count > 5)
                                {
                                    LogToRefresh($"  ... and {updates.Count - 5} more updates");
                                    DebugLogger.Info($"[FLAG-MANAGER]   ... and {updates.Count - 5} more updates");
                                }
                                
                                LogToRefresh($"✅ GLOBAL XML UPDATE: Calling UpsertFlagsWithIdsAndClashZoneData");
                                DebugLogger.Info($"[FLAG-MANAGER] ✅ GLOBAL XML UPDATE: Calling UpsertFlagsWithIdsAndClashZoneData");
                                // ✅ Convert to format expected by GlobalIndexService (remove OldSleeveInstanceId and OldClusterInstanceId)
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
                                
                                // ✅ PHASE 2: Only update Global XML if XML creation is enabled
                                if (!DeploymentConfiguration.DisableXmlCreation)
                                {
                                    GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, globalXmlUpdates, filterName: null, refreshLogName: refreshLogName);
                                    LogToRefresh($"✅ GLOBAL XML UPDATE: UpsertFlagsWithIdsAndClashZoneData completed successfully");
                                    DebugLogger.Info($"[FLAG-MANAGER] ✅ GLOBAL XML UPDATE: UpsertFlagsWithIdsAndClashZoneData completed successfully");
                                    
                                    DebugLogger.Info($"[FLAG-MANAGER] ✅ Updated Global XML for {updates.Count} entries in category '{category}' (reset {resetCount} flags)");
                                    LogToRefresh($"✅ Updated Global XML: {updates.Count} entries, {resetCount} flags reset for category '{category}'");
                                }
                                else
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ⚠️ XML creation disabled - skipping Global XML update for {updates.Count} entries in category '{category}' (database only mode)");
                                    LogToRefresh($"⚠️ XML creation disabled - skipping Global XML update for category '{category}' (database only mode)");
                                }
                            }
                            catch (Exception xmlEx)
                            {
                                // ✅ CRITICAL: Log Global XML update failure - this is blocking!
                                DebugLogger.Error($"[FLAG-MANAGER] ❌ FAILED to update Global XML for category '{category}': {xmlEx.Message}");
                                DebugLogger.Error($"[FLAG-MANAGER] ❌ Stack trace: {xmlEx.StackTrace}");
                                LogToRefresh($"❌ ERROR: Failed to update Global XML for category '{category}': {xmlEx.Message}");
                                throw; // Re-throw - Global XML update is critical
                            }
                            
                            totalResetCount += resetCount;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✅ Reset {resetCount} Global XML entries in category '{category}' - sleeves were deleted");
                            LogToRefresh($"Reset {resetCount} entries for category '{category}' (deleted sleeves detected).");
                            
                            // ✅ OPTIONAL: Sync back to Filter XML clash zones if provided
                            if (clashZonesByCategory != null && clashZonesByCategory.ContainsKey(category))
                            {
                                var filterClashZones = clashZonesByCategory[category];
                                if (filterClashZones != null && filterClashZones.Count > 0)
                                {
                                    // Sync reset flags back to Filter XML clash zones
                                    var resetGuids = new HashSet<Guid>(updates.Select(u => u.Id));
                                    foreach (var clashZone in filterClashZones)
                                    {
                                        if (clashZone != null && resetGuids.Contains(clashZone.Id))
                                        {
                                            var resetUpdate = updates.FirstOrDefault(u => u.Id == clashZone.Id);
                                            if (resetUpdate.Id != Guid.Empty)
                                            {
                                                clashZone.IsResolved = resetUpdate.IsResolved;
                                                clashZone.IsClusterResolved = resetUpdate.IsClusterResolved;
                                                clashZone.SleeveInstanceId = resetUpdate.SleeveInstanceId;
                                                clashZone.ClusterSleeveInstanceId = resetUpdate.ClusterSleeveInstanceId;
                                                clashZone.LastUpdated = DateTime.Now;
                                            }
                                        }
                                    }
                                    
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] ✅ Synced reset flags back to {filterClashZones.Count} Filter XML clash zones for category '{category}'");
                                }
                            }
                        }
                        else
                        {
                            if (resetCount > 0)
                                LogToRefresh($"Category '{category}' WARNING: resetCount={resetCount} but updates list is empty – investigate merge logic.");
                            else
                                LogToRefresh($"Category '{category}' generated no flag updates.");
                        }
                    }
                    catch (Exception categoryEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[FLAG-MANAGER] Error checking Global XML for category '{category}': {categoryEx.Message}");
                    }
                }
                
                if (totalResetCount > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] ✅ Total {totalResetCount} Global XML entries reset due to deleted sleeves");
                    if (!string.IsNullOrWhiteSpace(refreshLogName))
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] Total resets across all categories: {totalResetCount}\n");
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] All sleeves exist - no flags reset");
                    if (!string.IsNullOrWhiteSpace(refreshLogName))
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [FLAG-MANAGER] No flags reset for any category – all tracked sleeves still present.\n");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error checking Global XML for deleted sleeves: {ex.Message}");
            }
            
            return totalResetCount;
        }
        
        /// <summary>
        /// ⚠️ DEPRECATED: Use ResetFlagsForDeletedSleeves(List&lt;string&gt; categories) instead.
        /// This method is kept for backward compatibility but will be removed in future versions.
        /// </summary>
        /// <param name="clashZones">List of clash zones to check</param>
        /// <param name="category">MEP element category name</param>
        [Obsolete("Use ResetFlagsForDeletedSleeves(List<string> categories) instead. This method will be removed in future versions.")]
        public void ResetFlagsForDeletedSleeves(List<ClashZone> clashZones, string category)
        {
            if (clashZones == null || clashZones.Count == 0)
                return;
                
            if (string.IsNullOrWhiteSpace(category))
                return;
            
            try
            {
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId)>();
                int resetCount = 0;
                
                // ✅ OPTIMIZATION: Pre-collect all sleeve IDs by category (calculate once, use many times)
                // Same optimization pattern as recovery method and ClashZoneService
                var allSleeves = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => 
                    {
                        // Check if it's a sleeve family
                        bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                               s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                        
                        string familyName = s.Symbol?.FamilyName ?? "";
                        bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                            familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                        
                        if (!((s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") &&
                               (hasSleeveKeyword || isKnownFamily)))
                            return false;
                        
                        // ✅ PERFORMANCE: Filter by MEP_Category parameter (no Revit API call needed)
                        var mepCategoryParam = s.LookupParameter("MEP_Category");
                        if (mepCategoryParam != null && !string.IsNullOrWhiteSpace(mepCategoryParam.AsString()))
                        {
                            string sleeveCategory = mepCategoryParam.AsString();
                            return string.Equals(sleeveCategory, category, StringComparison.OrdinalIgnoreCase);
                        }
                        
                        return false;
                    })
                    .ToList();
                
                // ✅ OPTIMIZATION: Build HashSet for O(1) lookup (calculate once, use many times)
                var existingSleeveIdsSet = new HashSet<int>(allSleeves.Select(s => s.Id.IntegerValue));
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[FLAG-MANAGER] ===== STARTING DELETED SLEEVE CHECK FOR {clashZones.Count} CLASH ZONES =====");
                    DebugLogger.Info($"[FLAG-MANAGER] Pre-collected {existingSleeveIdsSet.Count} sleeve IDs for category '{category}' (O(1) lookup)");
                }
                
                // ✅ CRITICAL FIX: Pre-load all entries from hierarchical structure (calculate once, use many times)
                var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                
                foreach (var clashZone in clashZones)
                {
                    if (clashZone == null) continue;
                    
                    // ✅ CRITICAL FIX: Use pre-loaded allEntries instead of globalIndex.Entries
                    // Entries are now stored in Filters → FileCombos → Entries, not just in flat Entries list
                    var globalEntry = allEntries?.FirstOrDefault(e => 
                        string.Equals(e.Id, clashZone.Id.ToString(), StringComparison.OrdinalIgnoreCase));
                    
                    bool globalSaysClusterResolved = globalEntry?.IsClusterResolved ?? false;
                    bool globalSaysResolved = globalEntry?.IsResolved ?? false;
                    
                    // ✅ DEBUG: Log clash zone state before checking (deployment mode wrapped)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[FLAG-MANAGER] Checking ClashZone {clashZone.Id}: IsClusterResolved={clashZone.IsClusterResolved}, IsResolved={clashZone.IsResolved}, ClusterSleeveId={clashZone.ClusterSleeveInstanceId}, SleeveId={clashZone.SleeveInstanceId}");
                        DebugLogger.Info($"[FLAG-MANAGER]   Global XML says: IsClusterResolved={globalSaysClusterResolved}, IsResolved={globalSaysResolved}, ClusterSleeveId={globalEntry?.ClusterSleeveInstanceId ?? -1}, SleeveId={globalEntry?.SleeveInstanceId ?? -1}");
                    }
                    
                    // ✅ PROTECTED LOGIC: Flag Hierarchy - Check cluster FIRST (cluster flags take precedence)
                    // ⚠️ DO NOT MODIFY THIS HIERARCHY - Cluster flags must be checked before individual flags
                    if (clashZone.IsClusterResolved)
                    {
                        // ✅ CRITICAL PROTECTED LOGIC: Always check Revit API FIRST (Global XML may be stale if sleeve was deleted)
                        // WORKING LOGIC: Revit is authoritative source - if sleeve doesn't exist in Revit → reset flags
                        // DO NOT trust Global XML blindly - verify the sleeve actually exists in Revit
                        // Use ClusterSleeveInstanceId from clashZone (synced from Global XML) or fallback to Global XML entry
                        int clusterSleeveId = clashZone.ClusterSleeveInstanceId > 0 
                            ? clashZone.ClusterSleeveInstanceId 
                            : (globalEntry?.ClusterSleeveInstanceId ?? -1);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → Checking cluster sleeve ID {clusterSleeveId} using pre-collected HashSet...");
                        // ✅ OPTIMIZATION: Use pre-collected HashSet for O(1) lookup instead of Revit API call
                        bool clusterSleeveExists = existingSleeveIdsSet.Contains(clusterSleeveId);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → HashSet check result: {clusterSleeveExists}");
                        
                        if (!clusterSleeveExists)
                        {
                            // ✅ Cluster sleeve deleted → Reset ALL flags (regardless of Global XML state)
                            // Global XML may still say IsClusterResolved=true, but we verify Revit is authoritative
                            clashZone.IsClusterResolved = false;
                            clashZone.IsResolved = false;
                            clashZone.ClusterSleeveInstanceId = -1;
                            clashZone.SleeveInstanceId = -1;
                            clashZone.SleeveFamilyName = string.Empty;
                            clashZone.LastUpdated = DateTime.Now;
                            
                            updates.Add((clashZone.Id, false, false, -1, -1));
                            resetCount++;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset ALL flags for ClashZone {clashZone.Id} - cluster sleeve {clusterSleeveId} NOT FOUND in Revit (Global XML was stale)");
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ Verified: ClashZone {clashZone.Id} - cluster sleeve {clusterSleeveId} EXISTS in Revit");
                        }
                    }
                    // ✅ PROTECTED LOGIC: Then check individual (only if cluster flag is false)
                    // ⚠️ DO NOT MODIFY THIS ORDER - Individual check must come AFTER cluster check
                    else if (clashZone.IsResolved)
                    {
                        // ✅ CRITICAL PROTECTED LOGIC: Always check Revit API FIRST (Global XML may be stale if sleeve was deleted)
                        // WORKING LOGIC: Revit is authoritative source - if sleeve doesn't exist in Revit → reset flags
                        // DO NOT trust Global XML blindly - verify the sleeve actually exists in Revit
                        // Use SleeveInstanceId from clashZone (synced from Global XML) or fallback to Global XML entry
                        int individualSleeveId = clashZone.SleeveInstanceId > 0 
                            ? clashZone.SleeveInstanceId 
                            : (globalEntry?.SleeveInstanceId ?? -1);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → Checking individual sleeve ID {individualSleeveId} using pre-collected HashSet...");
                        // ✅ OPTIMIZATION: Use pre-collected HashSet for O(1) lookup instead of Revit API call
                        bool individualSleeveExists = existingSleeveIdsSet.Contains(individualSleeveId);
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → HashSet check result: {individualSleeveExists}");
                        
                        if (!individualSleeveExists)
                        {
                            // ✅ Individual sleeve deleted → Reset individual flag (regardless of Global XML state)
                            // Global XML may still say IsResolved=true, but we verify Revit is authoritative
                            clashZone.IsResolved = false;
                            clashZone.SleeveInstanceId = -1;
                            clashZone.SleeveFamilyName = string.Empty;
                            clashZone.LastUpdated = DateTime.Now;
                            
                            updates.Add((clashZone.Id, false, clashZone.IsClusterResolved, -1, clashZone.ClusterSleeveInstanceId));
                            resetCount++;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset individual flags for ClashZone {clashZone.Id} - individual sleeve {individualSleeveId} NOT FOUND in Revit (Global XML was stale)");
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ Verified: ClashZone {clashZone.Id} - individual sleeve {individualSleeveId} EXISTS in Revit");
                        }
                    }
                    else
                    {
                        // Neither cluster nor individual flags are set - nothing to check
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER]   → Skipping ClashZone {clashZone.Id} - neither IsClusterResolved nor IsResolved is true");
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] ===== COMPLETED DELETED SLEEVE CHECK: {resetCount} flags reset =====");
                
                // Save updated flags to Global XML using existing service with MEP+Host+Point data
                if (updates.Count > 0)
                {
                    // Get MEP+Host+Point data for each clash zone
                    var updatesWithData = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ)>();
                    
                    foreach (var update in updates)
                    {
                        var clashZone = clashZones.FirstOrDefault(cz => cz.Id == update.Id);
                        if (clashZone != null)
                        {
                            int mepId = clashZone.MepElementId?.IntegerValue ?? clashZone.MepElementIdValue;
                            int hostId = clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue;
                            updatesWithData.Add((update.Id, update.IsResolved, update.IsClusterResolved, update.SleeveInstanceId, update.ClusterSleeveInstanceId, mepId, hostId, clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ));
                        }
                    }
                    
                    if (updatesWithData.Count > 0)
                    {
                        // ✅ PHASE 2: Only update Global XML if XML creation is enabled
                        if (!DeploymentConfiguration.DisableXmlCreation)
                    {
                        // ✅ CRITICAL FIX: Try to preserve FilterName from Global XML entries when resetting flags
                        // If FilterName exists in Global XML, use it; otherwise leave empty
                        GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, updatesWithData, filterName: null);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Updated Global XML for {updates.Count} clash zones with reset flags in category '{category}'");
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ⚠️ XML creation disabled - skipping Global XML update for {updates.Count} clash zones in category '{category}' (database only mode)");
                        }
                    }
                }
                
                if (resetCount > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Reset flags for {resetCount} clash zones in category '{category}'");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error resetting flags for deleted sleeves in category '{category}': {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// ⚠️ DEPRECATED: Use ResetFlagsForDeletedSleeves(List&lt;string&gt; categories) instead.
        /// This method is kept for backward compatibility but will be removed in future versions.
        /// </summary>
        /// <param name="categories">List of categories to check</param>
        /// <returns>Number of flags reset</returns>
        [Obsolete("Use ResetFlagsForDeletedSleeves(List<string> categories) instead. This method will be removed in future versions.")]
        public int ResetFlagsForDeletedSleevesFromGlobalXml(List<string> categories)
        {
            if (categories == null || categories.Count == 0)
                return 0;
            
            int totalResetCount = 0;
            
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] ===== CHECKING GLOBAL XML FOR DELETED SLEEVES FOR {categories.Count} CATEGORIES =====");
                
                foreach (var category in categories)
                {
                    if (string.IsNullOrWhiteSpace(category))
                        continue;
                    
                    try
                    {
                        var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                        
                        // ✅ CRITICAL FIX: Use GetAllEntries to get entries from BOTH hierarchical and flat structures
                        // Entries are now stored in Filters → FileCombos → Entries, not just in flat Entries list
                        var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                        
                        if (allEntries == null || allEntries.Count == 0)
                            continue;
                        
                        var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId)>();
                        int resetCount = 0;
                        
                        // Check ALL Global XML entries (even if not in Filter XML)
                        foreach (var globalEntry in allEntries)
                        {
                            if (globalEntry == null) continue;
                            
                            // Only check entries with resolved flags (no need to check unresolved ones)
                            if (!globalEntry.IsResolved && !globalEntry.IsClusterResolved)
                                continue;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] Checking Global XML entry {globalEntry.Id}: IsClusterResolved={globalEntry.IsClusterResolved}, IsResolved={globalEntry.IsResolved}, ClusterSleeveId={globalEntry.ClusterSleeveInstanceId}, SleeveId={globalEntry.SleeveInstanceId}");
                            
                            // Check cluster sleeve first (flag hierarchy)
                            if (globalEntry.IsClusterResolved && globalEntry.ClusterSleeveInstanceId > 0)
                            {
                                bool clusterSleeveExists = CheckClusterSleeveExists(globalEntry.ClusterSleeveInstanceId);
                                if (!clusterSleeveExists)
                                {
                                    // Cluster sleeve deleted → Reset ALL flags
                                    updates.Add((Guid.Parse(globalEntry.Id), false, false, -1, -1,
                                                globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    resetCount++;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset ALL flags for Global XML entry {globalEntry.Id} - cluster sleeve {globalEntry.ClusterSleeveInstanceId} NOT FOUND");
                                }
                                continue; // Skip individual check if cluster was checked
                            }
                            
                            // Check individual sleeve
                            if (globalEntry.IsResolved && globalEntry.SleeveInstanceId > 0)
                            {
                                bool individualSleeveExists = CheckIndividualSleeveExists(globalEntry.SleeveInstanceId);
                                if (!individualSleeveExists)
                                {
                                    // Individual sleeve deleted → Reset individual flag only
                                    updates.Add((Guid.Parse(globalEntry.Id), false, globalEntry.IsClusterResolved, -1, globalEntry.ClusterSleeveInstanceId,
                                                globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ,
                                                globalEntry.SleeveInstanceId, globalEntry.ClusterSleeveInstanceId)); // ✅ OLD values for matching
                                    resetCount++;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset individual flag for Global XML entry {globalEntry.Id} - individual sleeve {globalEntry.SleeveInstanceId} NOT FOUND");
                                }
                            }
                        }
                        
                        // Save updated flags to Global XML
                        if (updates.Count > 0)
                        {
                            // ✅ CRITICAL FIX: FilterName is preserved from Global XML entry (not overwritten)
                            // ✅ Convert to format expected by GlobalIndexService (remove OldSleeveInstanceId and OldClusterInstanceId)
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
                            
                            // ✅ PHASE 2: Only update Global XML if XML creation is enabled
                            if (!DeploymentConfiguration.DisableXmlCreation)
                            {
                                GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, globalXmlUpdates, filterName: null);
                            totalResetCount += resetCount;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✅ Reset {resetCount} Global XML entries in category '{category}' - sleeves were deleted");
                            }
                            else
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] ⚠️ XML creation disabled - skipping Global XML reset for {resetCount} entries in category '{category}' (database only mode)");
                            }
                        }
                    }
                    catch (Exception categoryEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[FLAG-MANAGER] Error checking Global XML for category '{category}': {categoryEx.Message}");
                    }
                }
                
                if (totalResetCount > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] ✅ Total {totalResetCount} Global XML entries reset due to deleted sleeves");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error checking Global XML for deleted sleeves: {ex.Message}");
            }
            
            return totalResetCount;
        }
        
        /// <summary>
        /// Updates both in-memory ClashZone object and Global XML.
        /// ✅ CRITICAL: Accepts FilterName to identify which Filter XML file contains placement data
        /// </summary>
        /// <param name="clashZone">The clash zone that had a sleeve placed</param>
        /// <param name="sleeveId">The Revit element ID of the placed sleeve</param>
        /// <param name="isCluster">True if this is a cluster sleeve, false for individual sleeve</param>
        /// <param name="category">MEP element category name</param>
        /// <param name="filterName">Filter name that contains this clash zone's placement data (optional)</param>
        public void UpdateFlagsForPlacement(ClashZone clashZone, int sleeveId, bool isCluster, string category, string filterName = null)
        {
            if (clashZone == null)
                throw new ArgumentNullException(nameof(clashZone));
                
            if (sleeveId <= 0)
                throw new ArgumentException("Sleeve ID must be greater than 0", nameof(sleeveId));
                
            if (string.IsNullOrWhiteSpace(category))
                throw new ArgumentException("Category cannot be null or empty", nameof(category));
            
            try
            {
                if (isCluster)
                {
                    // Cluster sleeve placed
                    clashZone.IsClusterResolved = true;
                    clashZone.ClusterSleeveInstanceId = sleeveId;
                    clashZone.IsResolved = true; // Keep individual flag true when clustered
                    clashZone.SleeveInstanceId = -1; // Clear individual ID (individual sleeve was deleted during clustering)
                }
                else
                {
                    // Individual sleeve placed
                    clashZone.IsResolved = true;
                    clashZone.SleeveInstanceId = sleeveId;
                    // Note: IsClusterResolved remains false for individual sleeves
                }
                
                // ✅ OPTION 4: DATABASE-FIRST APPROACH - Update database first (authoritative source), then sync to Global XML
                int mepId = clashZone.MepElementId?.IntegerValue ?? clashZone.MepElementIdValue;
                int hostId = clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue;
                
                // Step 1: Update database first (triggers will auto-compute SleeveState)
                try
                {
                    using (var context = new SleeveDbContext(_document, msg =>
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER][SQLite] {msg}");
                    }))
                    {
                        var repository = new ClashZoneRepository(context, msg =>
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER][SQLite] {msg}");
                        });
                        
                        // ✅ Use BatchUpdateFlags for single update (database-first approach)
                        // ✅ CRITICAL FIX: Get OLD values BEFORE updating clashZone properties
                        // The clashZone properties were already updated above (lines 1915-1926), so we need to track old values
                        int oldSleeveInstanceId = isCluster ? clashZone.SleeveInstanceId : clashZone.SleeveInstanceId; // For clusters, individual sleeve was deleted (should be -1 or original)
                        int oldClusterInstanceId = clashZone.ClusterSleeveInstanceId; // Old cluster ID before update
                        
                        // ✅ CRITICAL: For cluster placement, the individual sleeve ID was already set to -1 in MarkClashZonesAsClusterResolvedWithSleeveId
                        // But we need the ORIGINAL sleeve ID for database matching
                        // Check if clashZone has AfterClusterSleevePlacedSleeveInstanceId (stored before clearing)
                        if (isCluster && clashZone.AfterClusterSleevePlacedSleeveInstanceId > 0)
                        {
                            oldSleeveInstanceId = clashZone.AfterClusterSleevePlacedSleeveInstanceId; // Use original individual sleeve ID for matching
                        }
                        
                        var singleUpdate = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId)>
                        {
                            (
                                clashZone.Id,
                                clashZone.IsResolved,
                                clashZone.IsClusterResolved,
                                clashZone.SleeveInstanceId,
                                clashZone.ClusterSleeveInstanceId,
                                clashZone.MepElementIdValue,
                                clashZone.StructuralElementIdValue,
                                clashZone.IntersectionPointX,
                                clashZone.IntersectionPointY,
                                clashZone.IntersectionPointZ,
                                oldSleeveInstanceId, // ✅ CRITICAL: Use original sleeve ID for matching
                                oldClusterInstanceId // Old cluster ID before update
                            )
                        };
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[FLAG-MANAGER] 📝 Calling BatchUpdateFlags for ClashZone {clashZone.Id}: " +
                                $"IsResolved={clashZone.IsResolved}, IsClusterResolved={clashZone.IsClusterResolved}, " +
                                $"SleeveId={clashZone.SleeveInstanceId}, ClusterId={clashZone.ClusterSleeveInstanceId}, " +
                                $"OldSleeveId={oldSleeveInstanceId}, OldClusterId={oldClusterInstanceId}, " +
                                $"GUID={clashZone.Id}, MEP={clashZone.MepElementIdValue}, Host={clashZone.StructuralElementIdValue}");
                        }
                        
                        repository.BatchUpdateFlags(singleUpdate);
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] ✅ Updated database flags for ClashZone {clashZone.Id} (database-first)");
                    }
                }
                catch (Exception dbEx)
                {
                    // ✅ NON-BLOCKING: Log database error but continue with Global XML sync
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[FLAG-MANAGER] ⚠️ Failed to update database flags for ClashZone {clashZone.Id}: {dbEx.Message}");
                }
                
                // Step 2: Sync to Global XML (for backward compatibility and cross-filter tracking)
                // ✅ PHASE 2: Only update Global XML if XML creation is enabled
                if (!DeploymentConfiguration.DisableXmlCreation)
                {
                GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, new[] { 
                    (clashZone.Id, clashZone.IsResolved, clashZone.IsClusterResolved, 
                     clashZone.SleeveInstanceId, clashZone.ClusterSleeveInstanceId,
                     mepId, hostId, clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ) 
                }, filterName);
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Updated flags for ClashZone {clashZone.Id} after {(isCluster ? "cluster" : "individual")} sleeve placement (SleeveId={sleeveId}, Filter='{filterName ?? "N/A"}')");
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] ⚠️ XML creation disabled - skipping Global XML update for ClashZone {clashZone.Id} (database only mode)");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error updating flags for placement: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Checks if a cluster sleeve exists in the active Revit document.
        /// </summary>
        /// <param name="sleeveId">The Revit element ID of the cluster sleeve</param>
        /// <returns>True if the sleeve exists, false otherwise</returns>
        private bool CheckClusterSleeveExists(int sleeveId)
        {
            if (sleeveId <= 0) return false;
            
            try
            {
                var element = _document.GetElement(new ElementId(sleeveId));
                return element != null;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Checks if an individual sleeve exists in the active Revit document.
        /// </summary>
        /// <param name="sleeveId">The Revit element ID of the individual sleeve</param>
        /// <returns>True if the sleeve exists, false otherwise</returns>
        private bool CheckIndividualSleeveExists(int sleeveId)
        {
            if (sleeveId <= 0) return false;
            
            try
            {
                var element = _document.GetElement(new ElementId(sleeveId));
                return element != null;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// ✅ OPTIMIZED RECOVERY: Incrementally recovers sleeves only for missing file combos.
        /// Checks if Global XML exists and compares UI state file combos with Global XML combos.
        /// Recovery logic:
        /// 1. If Global XML doesn't exist → Full recovery (all sleeves for category)
        /// 2. If Global XML exists → Check UI state file combos vs Global XML combos
        ///    - If all combos exist → Skip recovery (already processed)
        ///    - If any combo missing → Partial recovery (only sleeves for missing combos)
        /// </summary>
        /// <param name="category">MEP element category name</param>
        /// <param name="uiStateFileCombos">List of file combos from UI state (LinkedFile, HostFile) pairs (optional, null = full recovery)</param>
        /// <returns>Number of Global XML entries created/updated with sleeve IDs</returns>
        public int RecoverSleeveFlagsFromRevitOptimized(string category, List<(string LinkedFile, string HostFile)> uiStateFileCombos = null)
        {
            if (string.IsNullOrWhiteSpace(category))
                return 0;
            
            try
            {
                // ✅ STEP 1: Check if Global XML exists (fresh check)
                string globalXmlPath = GlobalIndexService.GetCategoryIndexPath(_document, category);
                bool isGlobalXmlFresh = !System.IO.File.Exists(globalXmlPath);
                
                if (isGlobalXmlFresh)
                {
                    // ✅ Fresh Global XML → Full recovery (all sleeves for category)
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] [OPTIMIZED-RECOVERY] Global XML for category '{category}' doesn't exist - performing FULL recovery");
                    
                    return RecoverSleeveFlagsFromRevit(category, fileComboFilter: null);
                }
                
                // ✅ STEP 2: Global XML exists → Check which file combos are missing
                if (uiStateFileCombos == null || uiStateFileCombos.Count == 0)
                {
                    // No UI state provided → Skip recovery (can't determine what to recover)
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] [OPTIMIZED-RECOVERY] Global XML exists but no UI state file combos provided - skipping recovery");
                    return 0;
                }
                
                // Get processed file combos from Global XML
                var processedKeys = GlobalIndexService.GetProcessedFileComboKeys(_document, category);
                var processedKeysSet = new HashSet<string>(processedKeys, StringComparer.OrdinalIgnoreCase);
                
                // Check which UI state combos are missing from Global XML
                var missingCombos = new List<(string LinkedFile, string HostFile)>();
                foreach (var combo in uiStateFileCombos)
                {
                    var comboObj = new ProcessedFileCombo { LinkedFile = combo.LinkedFile, HostFile = combo.HostFile };
                    var normalizedKey = comboObj.GetNormalizedKey();
                    
                    if (!processedKeysSet.Contains(normalizedKey))
                    {
                        missingCombos.Add(combo);
                    }
                }
                
                if (missingCombos.Count == 0)
                {
                    // ✅ CRITICAL FIX: Even if all combos exist, entries might have been just created with IsResolved="false"
                    // Recovery must still run to update entries that need sleeve IDs
                    // Check if any entries in existing combos need updating (have IsResolved="false" or SleeveInstanceId="-1")
                    var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                    bool needsRecovery = false;
                    int unresolvedCount = 0;
                    int totalEntryCount = 0;
                    
                    // ✅ CRITICAL FIX: Check ALL entries, not just entries in matching file combos
                    // Entries might be in any filter group or file combo group, so we need to check all of them
                    var allEntries = new List<CategoryGlobalIndexEntry>();
                    
                    // Collect entries from hierarchical structure
                    if (globalIndex.Filters != null && globalIndex.Filters.Count > 0)
                    {
                        foreach (var filter in globalIndex.Filters)
                        {
                            if (filter.FileCombos != null)
                            {
                                foreach (var fileCombo in filter.FileCombos)
                                {
                                    if (fileCombo.Entries != null)
                                    {
                                        allEntries.AddRange(fileCombo.Entries);
                                    }
                                }
                            }
                        }
                    }
                    
                    // Also collect entries from flat structure
                    if (globalIndex.Entries != null && globalIndex.Entries.Count > 0)
                    {
                        allEntries.AddRange(globalIndex.Entries);
                    }
                    
                    totalEntryCount = allEntries.Count;
                    var unresolvedEntries = allEntries.Where(e => !e.IsResolved || e.SleeveInstanceId <= 0).ToList();
                    unresolvedCount = unresolvedEntries.Count;
                    
                    if (unresolvedEntries.Count > 0)
                    {
                        needsRecovery = true;
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] [OPTIMIZED-RECOVERY] Found {unresolvedEntries.Count} unresolved entries out of {totalEntryCount} total entries");
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] [OPTIMIZED-RECOVERY] Checked {totalEntryCount} total entries, {unresolvedCount} unresolved entries found");
                    
                    if (!needsRecovery)
                    {
                        // ✅ All combos exist AND all entries are resolved → Skip recovery
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] [OPTIMIZED-RECOVERY] All {uiStateFileCombos.Count} UI state file combos exist in Global XML with resolved entries ({totalEntryCount} entries, {unresolvedCount} unresolved) - skipping recovery");
                        return 0;
                    }
                    
                    // ✅ Entries exist but need updating → Run recovery for all UI state combos
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] [OPTIMIZED-RECOVERY] All combos exist but {unresolvedCount} entries need updating ({totalEntryCount} total) - performing recovery for {uiStateFileCombos.Count} file combos");
                    
                    return RecoverSleeveFlagsFromRevit(category, fileComboFilter: uiStateFileCombos);
                }
                
                // ✅ STEP 3: Some combos missing → Partial recovery (only for missing combos)
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] [OPTIMIZED-RECOVERY] {missingCombos.Count} of {uiStateFileCombos.Count} file combos missing in Global XML - performing PARTIAL recovery");
                
                return RecoverSleeveFlagsFromRevit(category, fileComboFilter: missingCombos);
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error in optimized recovery for category '{category}': {ex.Message}");
                return 0;
            }
        }
        
        /// <summary>
        /// ✅ INCREMENTAL RECOVERY MECHANISM: Scans ALL sleeves in Revit and ensures each has a Global XML entry.
        /// This handles cases where:
        /// 1. Global XML was deleted and recreated (fresh recovery)
        /// 2. New file combo added but sleeves already exist (incremental recovery)
        /// 3. Sleeves exist but Global XML entries are missing for some file combos
        /// 
        /// Uses MEP_Category parameter on sleeves to filter by category (no expensive Revit API calls).
        /// For each sleeve, checks if Global XML has a matching entry (by MEP+Host+Point).
        /// If no entry exists → CREATES entry in correct FilterGroup using sleeve's FilterName parameter.
        /// If entry exists → UPDATES flags with sleeve IDs.
        /// 
        /// FilterName format: "Plumbing_pipes.xml" → extracts "Plumbing" (before underscore) for FilterGroup.
        /// </summary>
        /// <param name="category">MEP element category name</param>
        /// <param name="fileComboFilter">Optional list of file combos to filter recovery (null = all sleeves)</param>
        /// <returns>Number of Global XML entries created/updated with sleeve IDs</returns>
        public int RecoverSleeveFlagsFromRevit(string category, List<(string LinkedFile, string HostFile)> fileComboFilter = null)
        {
            if (string.IsNullOrWhiteSpace(category))
                return 0;
            
            int recoveredCount = 0;
            int createdCount = 0;
            int updatedCount = 0;
            
            try
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] ===== INCREMENTAL RECOVERY: SCANNING ALL SLEEVES FOR CATEGORY '{category}' =====");
                
                // ✅ STEP 1: Filter sleeves by MEP_Category parameter (no expensive Revit API calls)
                // This is much faster than checking element IDs to determine category
                var categorySleeves = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s => 
                    {
                        // Check if it's a sleeve family
                        bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                               s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                        
                        string familyName = s.Symbol?.FamilyName ?? "";
                        bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                            familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                        
                        if (!((s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") &&
                               (hasSleeveKeyword || isKnownFamily)))
                            return false;
                        
                        // ✅ PERFORMANCE: Filter by MEP_Category parameter (no Revit API call needed)
                        var mepCategoryParam = s.LookupParameter("MEP_Category");
                        if (mepCategoryParam != null && !string.IsNullOrWhiteSpace(mepCategoryParam.AsString()))
                        {
                            string sleeveCategory = mepCategoryParam.AsString();
                            return string.Equals(sleeveCategory, category, StringComparison.OrdinalIgnoreCase);
                        }
                        
                        // Fallback: If MEP_Category parameter is missing, skip (category unknown)
                        return false;
                    })
                    .ToList();
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Found {categorySleeves.Count} sleeves for category '{category}' (filtered by MEP_Category parameter)");
                
                // ✅ STEP 2: Skip recovery if no sleeves found for this category
                if (categorySleeves.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] No sleeves found in Revit for category '{category}' - skipping recovery");
                    return 0;
                }
                
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                
                // ✅ STEP 3: Build index of existing entries by MEP+Host+Point for fast lookup
                var existingEntriesIndex = new Dictionary<(int MepId, int HostId, double X, double Y, double Z), CategoryGlobalIndexEntry>();
                var allEntries = new List<CategoryGlobalIndexEntry>();
                
                // ✅ HIERARCHICAL STRUCTURE: Collect entries from all filters → file combos
                if (globalIndex.Filters != null && globalIndex.Filters.Count > 0)
                {
                    foreach (var filter in globalIndex.Filters)
                    {
                        if (filter.FileCombos != null)
                        {
                            foreach (var fileCombo in filter.FileCombos)
                            {
                                if (fileCombo.Entries != null)
                                {
                                    allEntries.AddRange(fileCombo.Entries);
                                }
                            }
                        }
                    }
                }
                
                // ✅ FALLBACK: Also include entries from flat structure (backward compatibility)
                if (globalIndex.Entries != null && globalIndex.Entries.Count > 0)
                {
                    allEntries.AddRange(globalIndex.Entries);
                }
                
                // Build index for fast lookup (round to 0.1ft for tolerance - matches FindByMepHostAndPoint)
                // ✅ CRITICAL FIX: Tolerance must match FindByMepHostAndPoint default (0.1ft) to ensure recovery matches entries
                double pointTolerance = 0.1; // 0.1ft = ~30mm tolerance (matches GlobalIndexService.FindByMepHostAndPoint)
                foreach (var entry in allEntries)
                {
                    var key = (
                        entry.MepElementId,
                        entry.StructuralElementId,
                        Math.Round(entry.IntersectionPointX / pointTolerance) * pointTolerance,
                        Math.Round(entry.IntersectionPointY / pointTolerance) * pointTolerance,
                        Math.Round(entry.IntersectionPointZ / pointTolerance) * pointTolerance
                    );
                    if (!existingEntriesIndex.ContainsKey(key))
                    {
                        existingEntriesIndex[key] = entry;
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Built index of {existingEntriesIndex.Count} existing entries in Global XML for category '{category}'");
                
                var updatesToExisting = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ)>();
                var entriesToCreate = new List<(Guid Id, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, string FilterName, string LinkedFile, string HostFile)>();
                
                // ✅ STEP 4: OPTIMIZED RECOVERY - Calculate once, use many times
                // Pre-collect all sleeve IDs into HashSet for O(1) lookup (no Revit API calls needed)
                var existingSleeveIdsSet = new HashSet<int>(categorySleeves.Select(s => s.Id.IntegerValue));
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Pre-collected {existingSleeveIdsSet.Count} sleeve IDs for O(1) existence lookup");
                
                // ✅ STEP 4: For EACH SLEEVE, check if Global XML entry exists (INCREMENTAL RECOVERY)
                // Change from entry-driven to sleeve-driven: iterate through sleeves, not entries
                foreach (var sleeve in categorySleeves)
                {
                    try
                    {
                        // ✅ STEP 4a: Extract FilterName from sleeve parameter (format: "Plumbing_pipes.xml" → "Plumbing")
                        string filterName = string.Empty;
                        var filterNameParam = sleeve.LookupParameter("Filter Name");
                        if (filterNameParam != null && !string.IsNullOrWhiteSpace(filterNameParam.AsString()))
                        {
                            string filterNameValue = filterNameParam.AsString().Trim();
                            // Extract filter name before underscore (e.g., "Plumbing_pipes.xml" → "Plumbing")
                            int underscoreIndex = filterNameValue.IndexOf('_');
                            if (underscoreIndex > 0)
                            {
                                filterName = filterNameValue.Substring(0, underscoreIndex).Trim();
                            }
                            else
                            {
                                // No underscore - use entire value (fallback)
                                filterName = filterNameValue;
                            }
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Processing sleeve {sleeve.Id.IntegerValue}: FilterName='{filterName}'");
                        
                        // ✅ STEP 4b: Get MEP element from sleeve's MEP_ElementId parameter
                        var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
                        if (mepElementIdParam == null || !mepElementIdParam.HasValue)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] [RECOVERY] Sleeve {sleeve.Id.IntegerValue}: MEP_ElementId parameter missing - skipping");
                            continue;
                        }
                        
                        var sleeveMepElementId = mepElementIdParam.AsElementId();
                        if (sleeveMepElementId == null || sleeveMepElementId == ElementId.InvalidElementId)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] [RECOVERY] Sleeve {sleeve.Id.IntegerValue}: MEP_ElementId is invalid - skipping");
                            continue;
                        }
                        
                        // ✅ OPTIMIZATION: Use MEP ID directly from parameter (no expensive Revit API call to verify element exists)
                        // If sleeve exists in Revit, we trust the MEP_ElementId parameter is valid
                        int mepElementId = sleeveMepElementId.IntegerValue;
                        
                        // ✅ STEP 4c: Get Host element from sleeve.Host
                        if (sleeve.Host == null)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] [RECOVERY] Sleeve {sleeve.Id.IntegerValue}: Host is null - skipping");
                            continue;
                        }
                        
                        int hostElementId = sleeve.Host.Id.IntegerValue;
                        
                        // ✅ STEP 4d: Get intersection point from sleeve location
                        var sleeveLocation = (sleeve.Location as LocationPoint)?.Point ?? (sleeve.Location as LocationCurve)?.Curve?.GetEndPoint(0);
                        if (sleeveLocation == null)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] [RECOVERY] Sleeve {sleeve.Id.IntegerValue}: Cannot get location - skipping");
                            continue;
                        }
                        
                        double intersectionX = sleeveLocation.X;
                        double intersectionY = sleeveLocation.Y;
                        double intersectionZ = sleeveLocation.Z;
                        
                        // ✅ STEP 4e: Read LinkedFile and HostFile from sleeve parameters (performance optimization)
                        string linkedFile = string.Empty;
                        string hostFile = string.Empty;
                        
                        // ✅ OPTIMIZATION: Read LinkedFile and HostFile from sleeve parameters only (no Revit API calls for document lookup)
                        // Sleeve parameters are set during placement, so they should always be available
                        var linkedFileParam = sleeve.LookupParameter("LinkedFile");
                        if (linkedFileParam != null && !string.IsNullOrWhiteSpace(linkedFileParam.AsString()))
                        {
                            linkedFile = linkedFileParam.AsString().Trim();
                        }
                        
                        var hostFileParam = sleeve.LookupParameter("HostFile");
                        if (hostFileParam != null && !string.IsNullOrWhiteSpace(hostFileParam.AsString()))
                        {
                            hostFile = hostFileParam.AsString().Trim();
                        }
                        
                        // ✅ PERFORMANCE: Skip fallback document lookups (requires expensive Revit API calls)
                        // If parameters are missing, skip this sleeve (parameters should always be set during placement)
                        
                        if (string.IsNullOrWhiteSpace(linkedFile) || string.IsNullOrWhiteSpace(hostFile))
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] [RECOVERY] Sleeve {sleeve.Id.IntegerValue}: Missing file combo data (LinkedFile='{linkedFile}', HostFile='{hostFile}') - skipping");
                            continue;
                        }
                        
                        // ✅ OPTIMIZATION: If fileComboFilter is provided, skip sleeves not matching filter
                        if (fileComboFilter != null && fileComboFilter.Count > 0)
                        {
                            bool matchesFilter = false;
                            var sleeveCombo = new ProcessedFileCombo { LinkedFile = linkedFile, HostFile = hostFile };
                            string sleeveComboKey = sleeveCombo.GetNormalizedKey();
                            
                            foreach (var filterCombo in fileComboFilter)
                            {
                                var filterComboObj = new ProcessedFileCombo { LinkedFile = filterCombo.LinkedFile, HostFile = filterCombo.HostFile };
                                string filterComboKey = filterComboObj.GetNormalizedKey();
                                
                                if (sleeveComboKey == filterComboKey)
                                {
                                    matchesFilter = true;
                                    break;
                                }
                            }
                            
                            if (!matchesFilter)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Sleeve {sleeve.Id.IntegerValue}: File combo (LinkedFile='{linkedFile}', HostFile='{hostFile}') not in filter - skipping");
                                continue;
                            }
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Sleeve {sleeve.Id.IntegerValue}: LinkedFile='{linkedFile}', HostFile='{hostFile}' (from sleeve parameters)");
                        
                        // ✅ STEP 4f: Check if Global XML entry exists by MEP+Host+Point
                        var lookupKey = (
                            mepElementId,
                            hostElementId,
                            Math.Round(intersectionX / pointTolerance) * pointTolerance,
                            Math.Round(intersectionY / pointTolerance) * pointTolerance,
                            Math.Round(intersectionZ / pointTolerance) * pointTolerance
                        );
                        
                        CategoryGlobalIndexEntry existingEntry = null;
                        if (existingEntriesIndex.TryGetValue(lookupKey, out existingEntry))
                        {
                            // ✅ Entry exists - UPDATE flags
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] ✓ FOUND EXISTING ENTRY: Sleeve {sleeve.Id.IntegerValue} matches entry {existingEntry.Id}");
                            
                            // ✅ CRITICAL FIX: Check if entry is unresolved but sleeve exists in Revit
                            // If Global XML says unresolved but Revit has sleeve → update flags
                            bool needsUpdate = false;
                            if (!existingEntry.IsResolved || existingEntry.SleeveInstanceId <= 0)
                            {
                                // Entry is unresolved - check if this sleeve matches
                                needsUpdate = true;
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Entry {existingEntry.Id} is unresolved (IsResolved={existingEntry.IsResolved}, SleeveId={existingEntry.SleeveInstanceId}) - sleeve exists in Revit, updating flags");
                            }
                            
                            // Check if it's a cluster sleeve
                            bool isCluster = false;
                            var clusterParam = sleeve.LookupParameter("Cluster Sleeve Instance ID");
                            if (clusterParam != null && clusterParam.HasValue && clusterParam.AsInteger() > 0)
                                isCluster = true;
                            
                            var mepElementIdsParam = sleeve.LookupParameter("MEP_ElementIds");
                            if (!isCluster && mepElementIdsParam != null && !string.IsNullOrWhiteSpace(mepElementIdsParam.AsString()))
                                isCluster = true;
                            
                            if (isCluster)
                            {
                                int clusterSleeveId = sleeve.Id.IntegerValue;
                                // Update if entry doesn't already have this cluster sleeve ID OR if entry is unresolved
                                if (existingEntry.ClusterSleeveInstanceId != clusterSleeveId || needsUpdate)
                                {
                                    updatesToExisting.Add((Guid.Parse(existingEntry.Id), true, true, -1, clusterSleeveId,
                                        mepElementId, hostElementId, intersectionX, intersectionY, intersectionZ));
                                    updatedCount++;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✓ UPDATE: Entry {existingEntry.Id} → Cluster sleeve {clusterSleeveId} (was unresolved: {needsUpdate})");
                                }
                            }
                            else
                            {
                                int sleeveId = sleeve.Id.IntegerValue;
                                // Update if entry doesn't already have this sleeve ID OR if entry is unresolved
                                if (existingEntry.SleeveInstanceId != sleeveId || needsUpdate)
                                {
                                    updatesToExisting.Add((Guid.Parse(existingEntry.Id), true, false, sleeveId, -1,
                                        mepElementId, hostElementId, intersectionX, intersectionY, intersectionZ));
                                    updatedCount++;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✓ UPDATE: Entry {existingEntry.Id} → Individual sleeve {sleeveId} (was unresolved: {needsUpdate})");
                                }
                            }
                        }
                        else
                        {
                            // ✅ Entry does NOT exist - CREATE new entry with deterministic GUID
                            // Use deterministic GUID generation to ensure same intersection gets same GUID
                            Guid newEntryId;
                            try
                            {
                                // Generate deterministic GUID from MEP+Host+Point (matches CreateClashZone logic)
                                // This ensures recovery creates entries with same GUID that CreateClashZone would generate
                                var guidManager = new GuidManager(_document);
                                newEntryId = guidManager.GenerateDeterministicGuid(
                                    mepElementId,
                                    hostElementId,
                                    intersectionX,
                                    intersectionY,
                                    intersectionZ,
                                    tolerance: 0.1); // 0.1ft = ~30mm tolerance (matches GlobalIndexService.FindByMepHostAndPoint)
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Generated deterministic GUID {newEntryId} for MEP={mepElementId}, Host={hostElementId}, Point=({intersectionX:F3},{intersectionY:F3},{intersectionZ:F3})");
                            }
                            catch (Exception guidEx)
                            {
                                // Fallback to random GUID if deterministic generation fails
                                newEntryId = Guid.NewGuid();
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[FLAG-MANAGER] [RECOVERY] Error generating deterministic GUID: {guidEx.Message} - using random GUID {newEntryId}");
                            }
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] ✗ NO ENTRY FOUND: Creating new entry {newEntryId} for sleeve {sleeve.Id.IntegerValue}");
                            
                            // Check if it's a cluster sleeve
                            bool isCluster = false;
                            int sleeveId = sleeve.Id.IntegerValue;
                            var clusterParam = sleeve.LookupParameter("Cluster Sleeve Instance ID");
                            if (clusterParam != null && clusterParam.HasValue && clusterParam.AsInteger() > 0)
                                isCluster = true;
                            
                            var mepElementIdsParam = sleeve.LookupParameter("MEP_ElementIds");
                            if (!isCluster && mepElementIdsParam != null && !string.IsNullOrWhiteSpace(mepElementIdsParam.AsString()))
                                isCluster = true;
                            
                            entriesToCreate.Add((newEntryId, mepElementId, hostElementId, intersectionX, intersectionY, intersectionZ, filterName, linkedFile, hostFile));
                            createdCount++;
                            
                            // Add to updates list with sleeve ID and flags set to true
                            if (isCluster)
                            {
                                updatesToExisting.Add((newEntryId, true, true, -1, sleeveId,
                                    mepElementId, hostElementId, intersectionX, intersectionY, intersectionZ));
                            }
                            else
                            {
                                updatesToExisting.Add((newEntryId, true, false, sleeveId, -1,
                                    mepElementId, hostElementId, intersectionX, intersectionY, intersectionZ));
                            }
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✓ CREATE: New entry {newEntryId} for sleeve {sleeve.Id.IntegerValue} (Filter='{filterName}', LinkedFile='{linkedFile}', HostFile='{hostFile}')");
                        }
                    }
                    catch (Exception sleeveEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[FLAG-MANAGER] Error recovering sleeve {sleeve.Id.IntegerValue}: {sleeveEx.Message}");
                    }
                }
                
                // ✅ STEP 5: Verify unresolved entries against Revit (using pre-collected sleeve ID HashSet - O(1) lookup)
                // If Global XML says unresolved but Revit has sleeve → update flags
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Step 5: Verifying {allEntries.Count} entries against {existingSleeveIdsSet.Count} existing sleeve IDs");
                
                int step5UpdatedCount = 0;
                foreach (var entry in allEntries)
                {
                    // Only check unresolved entries
                    if (entry.IsResolved && entry.SleeveInstanceId > 0)
                        continue;
                    
                    // ✅ OPTIMIZED: Check if sleeve ID exists using pre-collected HashSet (O(1) lookup - no Revit API call)
                    bool sleeveExists = false;
                    int matchingSleeveId = -1;
                    
                    if (entry.SleeveInstanceId > 0 && existingSleeveIdsSet.Contains(entry.SleeveInstanceId))
                    {
                        sleeveExists = true;
                        matchingSleeveId = entry.SleeveInstanceId;
                    }
                    else if (entry.ClusterSleeveInstanceId > 0 && existingSleeveIdsSet.Contains(entry.ClusterSleeveInstanceId))
                    {
                        sleeveExists = true;
                        matchingSleeveId = entry.ClusterSleeveInstanceId;
                    }
                    else if (!entry.IsResolved && !entry.IsClusterResolved)
                    {
                        // Entry is unresolved - check if any sleeve exists at this MEP+Host+Point
                        var entryKey = (
                            MepId: entry.MepElementId,
                            HostId: entry.StructuralElementId,
                            X: Math.Round(entry.IntersectionPointX / pointTolerance) * pointTolerance,
                            Y: Math.Round(entry.IntersectionPointY / pointTolerance) * pointTolerance,
                            Z: Math.Round(entry.IntersectionPointZ / pointTolerance) * pointTolerance
                        );
                        
                        // Find matching sleeve by MEP+Host+Point (check if already processed in this recovery)
                        var matchingSleeve = categorySleeves.FirstOrDefault(s =>
                        {
                            var mepIdParam = s.LookupParameter("MEP_ElementId");
                            if (mepIdParam == null || !mepIdParam.HasValue) return false;
                            var mepId = mepIdParam.AsElementId();
                            if (mepId == null || mepId == ElementId.InvalidElementId) return false;
                            if (mepId.IntegerValue != entryKey.MepId) return false;
                            
                            if (s.Host == null) return false;
                            if (s.Host.Id.IntegerValue != entryKey.HostId) return false;
                            
                            var location = (s.Location as LocationPoint)?.Point ?? (s.Location as LocationCurve)?.Curve?.GetEndPoint(0);
                            if (location == null) return false;
                            
                            double roundedX = Math.Round(location.X / pointTolerance) * pointTolerance;
                            double roundedY = Math.Round(location.Y / pointTolerance) * pointTolerance;
                            double roundedZ = Math.Round(location.Z / pointTolerance) * pointTolerance;
                            
                            return Math.Abs(roundedX - entryKey.X) < pointTolerance &&
                                   Math.Abs(roundedY - entryKey.Y) < pointTolerance &&
                                   Math.Abs(roundedZ - entryKey.Z) < pointTolerance;
                        });
                        
                        if (matchingSleeve != null)
                        {
                            matchingSleeveId = matchingSleeve.Id.IntegerValue;
                            sleeveExists = true;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Entry {entry.Id} is unresolved but sleeve {matchingSleeveId} exists in Revit at MEP+Host+Point - updating flags");
                        }
                    }
                    
                    // ✅ Update flags if sleeve exists but entry says unresolved
                    if (sleeveExists && (!entry.IsResolved || entry.SleeveInstanceId <= 0))
                    {
                        // Check if it's a cluster sleeve
                        bool isCluster = false;
                        if (matchingSleeveId > 0)
                        {
                            var matchingSleeve = categorySleeves.FirstOrDefault(s => s.Id.IntegerValue == matchingSleeveId);
                            if (matchingSleeve != null)
                            {
                                var clusterParam = matchingSleeve.LookupParameter("Cluster Sleeve Instance ID");
                                if (clusterParam != null && clusterParam.HasValue && clusterParam.AsInteger() > 0)
                                    isCluster = true;
                                
                                var mepElementIdsParam = matchingSleeve.LookupParameter("MEP_ElementIds");
                                if (!isCluster && mepElementIdsParam != null && !string.IsNullOrWhiteSpace(mepElementIdsParam.AsString()))
                                    isCluster = true;
                            }
                        }
                        
                        if (isCluster)
                        {
                            updatesToExisting.Add((Guid.Parse(entry.Id), true, true, -1, matchingSleeveId,
                                entry.MepElementId, entry.StructuralElementId, entry.IntersectionPointX, entry.IntersectionPointY, entry.IntersectionPointZ));
                            updatedCount++;
                            step5UpdatedCount++;
                        }
                        else
                        {
                            updatesToExisting.Add((Guid.Parse(entry.Id), true, false, matchingSleeveId, -1,
                                entry.MepElementId, entry.StructuralElementId, entry.IntersectionPointX, entry.IntersectionPointY, entry.IntersectionPointZ));
                            updatedCount++;
                            step5UpdatedCount++;
                        }
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] ✓ VERIFIED: Entry {entry.Id} → {(isCluster ? "Cluster" : "Individual")} sleeve {matchingSleeveId} exists in Revit, updating flags");
                    }
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] [RECOVERY] Step 5 completed: {step5UpdatedCount} entries updated by verifying against Revit sleeves");
                
                // ✅ STEP 6: Create new entries first (so they exist before updating flags)
                if (entriesToCreate.Count > 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Creating {entriesToCreate.Count} new Global XML entries for sleeves without entries");
                    
                    foreach (var entryToCreate in entriesToCreate)
                    {
                        try
                        {
                            GlobalIndexService.EnsureEntriesWithClashZoneData(_document, category, new[]
                            {
                                (entryToCreate.Id, entryToCreate.MepElementId, entryToCreate.StructuralElementId,
                                 entryToCreate.IntersectionPointX, entryToCreate.IntersectionPointY, entryToCreate.IntersectionPointZ)
                            }, entryToCreate.FilterName, entryToCreate.LinkedFile, entryToCreate.HostFile);
                            
                            recoveredCount++;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✅ Created entry {entryToCreate.Id} in FilterGroup '{entryToCreate.FilterName}' for file combo (LinkedFile='{entryToCreate.LinkedFile}', HostFile='{entryToCreate.HostFile}')");
                        }
                        catch (Exception createEx)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Warning($"[FLAG-MANAGER] Error creating entry {entryToCreate.Id}: {createEx.Message}");
                        }
                    }
                }
                
                // ✅ STEP 6: Update existing entries with sleeve IDs and flags
                if (updatesToExisting.Count > 0)
                {
                    // ✅ PHASE 2: Only update Global XML if XML creation is enabled
                    if (!DeploymentConfiguration.DisableXmlCreation)
                {
                    GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, updatesToExisting, filterName: null);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] ✅ Updated {updatedCount} existing entries and created {createdCount} new entries with sleeve IDs from Revit for category '{category}'");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] ⚠️ XML creation disabled - skipping Global XML update for {updatesToExisting.Count} entries in category '{category}' (database only mode)");
                    }
                }
                else
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] No entries updated or created - all sleeves already have entries in Global XML for category '{category}'");
                }
                
                recoveredCount = createdCount + updatedCount;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error recovering sleeve flags from Revit for category '{category}': {ex.Message}");
            }
            
            return recoveredCount;
        }
        
        /// <summary>
        /// ✅ HELPER: Checks if Global XML file is fresh/newly created (needs recovery).
        /// A Global XML is considered "fresh" if:
        /// 1. File doesn't exist (will be created), OR
        /// 2. File exists but has entries with unresolved flags (all entries have IsResolved=false and no SleeveInstanceId)
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="category">MEP element category name</param>
        /// <returns>True if Global XML is fresh and needs recovery, false otherwise</returns>
        private bool IsGlobalXmlFresh(Document doc, string category)
        {
            try
            {
                string path = GlobalIndexService.GetCategoryIndexPath(doc, category);
                
                // ✅ Check 1: File doesn't exist - definitely fresh
                if (!System.IO.File.Exists(path))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Global XML file for category '{category}' does not exist - marked as fresh");
                    return true;
                }
                
                // ✅ Check 2: File exists - check if all entries are unresolved (no sleeves found yet)
                var globalIndex = GlobalIndexService.LoadOrCreate(doc, category);
                
                // Get all entries from hierarchical and flat structures
                var allEntries = new List<CategoryGlobalIndexEntry>();
                
                if (globalIndex.Filters != null && globalIndex.Filters.Count > 0)
                {
                    foreach (var filter in globalIndex.Filters)
                    {
                        if (filter.FileCombos != null)
                        {
                            foreach (var fileCombo in filter.FileCombos)
                            {
                                if (fileCombo.Entries != null)
                                {
                                    allEntries.AddRange(fileCombo.Entries);
                                }
                            }
                        }
                    }
                }
                
                if (globalIndex.Entries != null && globalIndex.Entries.Count > 0)
                {
                    allEntries.AddRange(globalIndex.Entries);
                }
                
                // Remove duplicates by ID
                allEntries = allEntries.GroupBy(e => e.Id, StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.First())
                    .ToList();
                
                if (allEntries.Count == 0)
                {
                    // No entries - fresh file
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Global XML for category '{category}' has no entries - marked as fresh");
                    return true;
                }
                
                // Check if ALL entries are unresolved (no sleeves found)
                bool allUnresolved = allEntries.All(e => 
                    !e.IsResolved && 
                    !e.IsClusterResolved && 
                    (e.SleeveInstanceId <= 0 || e.SleeveInstanceId == -1) && 
                    (e.ClusterSleeveInstanceId <= 0 || e.ClusterSleeveInstanceId == -1));
                
                if (allUnresolved)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Global XML for category '{category}' has {allEntries.Count} entries, all unresolved - marked as fresh");
                    return true;
                }
                
                // File exists and has resolved entries - not fresh
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Global XML for category '{category}' has resolved entries - not fresh, skipping recovery");
                return false;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[FLAG-MANAGER] Error checking if Global XML is fresh for category '{category}': {ex.Message}");
                // On error, assume not fresh (safer to skip recovery)
                return false;
            }
        }
        
        /// <summary>
        /// ✅ OOP HELPER: Deletes sleeve when intersection point changes beyond tolerance.
        /// If intersection point moves > tolerance (0.1ft), deletes old sleeve at old location and resets flags.
        /// This allows new sleeve to be placed at the new intersection point.
        /// </summary>
        /// <param name="clashZone">The clash zone with updated intersection point</param>
        /// <param name="category">MEP element category name</param>
        /// <param name="movementDistance">Distance the intersection point moved (in feet)</param>
        /// <param name="tolerance">Tolerance threshold for movement (default: 0.1ft = ~30mm)</param>
        /// <returns>True if sleeve was deleted, false otherwise</returns>
        public bool DeleteSleeveForIntersectionPointChange(ClashZone clashZone, string category, double movementDistance, double tolerance = 0.1)
        {
            if (clashZone == null)
                throw new ArgumentNullException(nameof(clashZone));
                
            if (string.IsNullOrWhiteSpace(category))
                throw new ArgumentException("Category cannot be null or empty", nameof(category));
            
            // Only delete if movement exceeds tolerance
            if (movementDistance <= tolerance)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Intersection point moved {movementDistance:F3}ft (≤{tolerance:F3}ft tolerance) - no sleeve deletion needed");
                return false;
            }
            
            try
            {
                bool deleted = false;
                
                // Find Global XML entry by GUID (clashZone.Id) - this gives us old sleeve IDs before point update
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
                
                // ✅ CRITICAL FIX: Use GetAllEntries to get entries from BOTH hierarchical and flat structures
                var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();
                var globalEntry = allEntries?.FirstOrDefault(e => 
                    string.Equals(e.Id, clashZone.Id.ToString(), StringComparison.OrdinalIgnoreCase));
                
                if (globalEntry == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[FLAG-MANAGER] No Global XML entry found for ClashZone {clashZone.Id} - cannot delete sleeve");
                    return false;
                }
                
                // Delete cluster sleeve if exists
                if (globalEntry.IsClusterResolved && globalEntry.ClusterSleeveInstanceId > 0)
                {
                    var clusterSleeve = _document.GetElement(new ElementId(globalEntry.ClusterSleeveInstanceId));
                    if (clusterSleeve != null)
                    {
                        try
                        {
                            _document.Delete(clusterSleeve.Id);
                            deleted = true;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✅ Deleted cluster sleeve {globalEntry.ClusterSleeveInstanceId} due to intersection point change (Δ={movementDistance:F3}ft > {tolerance:F3}ft tolerance)");
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[FLAG-MANAGER] Error deleting cluster sleeve {globalEntry.ClusterSleeveInstanceId}: {ex.Message}");
                            throw;
                        }
                    }
                }
                
                // Delete individual sleeve if exists (and not already deleted by cluster deletion)
                if (globalEntry.IsResolved && globalEntry.SleeveInstanceId > 0 && globalEntry.SleeveInstanceId != globalEntry.ClusterSleeveInstanceId)
                {
                    var individualSleeve = _document.GetElement(new ElementId(globalEntry.SleeveInstanceId));
                    if (individualSleeve != null)
                    {
                        try
                        {
                            _document.Delete(individualSleeve.Id);
                            deleted = true;
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✅ Deleted individual sleeve {globalEntry.SleeveInstanceId} due to intersection point change (Δ={movementDistance:F3}ft > {tolerance:F3}ft tolerance)");
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[FLAG-MANAGER] Error deleting individual sleeve {globalEntry.SleeveInstanceId}: {ex.Message}");
                            throw;
                        }
                    }
                }
                
                // Reset flags after deletion
                if (deleted)
                {
                    clashZone.IsResolved = false;
                    clashZone.IsClusterResolved = false;
                    clashZone.SleeveInstanceId = -1;
                    clashZone.ClusterSleeveInstanceId = -1;
                    
                    // Update Global XML with reset flags and new intersection point
                    int mepId = clashZone.MepElementId?.IntegerValue ?? clashZone.MepElementIdValue;
                    int hostId = clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue;
                    
                    // ✅ PHASE 2: Only update Global XML if XML creation is enabled
                    if (!DeploymentConfiguration.DisableXmlCreation)
                    {
                    // ✅ CRITICAL FIX: FilterName not available in this context, but flag reset will preserve existing FilterName
                    GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, new[] { 
                        (clashZone.Id, false, false, -1, -1,
                         mepId, hostId, clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ) 
                    }, filterName: null);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] ✅ Reset flags for ClashZone {clashZone.Id} after sleeve deletion due to intersection point change");
                    }
                    else
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info($"[FLAG-MANAGER] ⚠️ XML creation disabled - skipping Global XML reset for ClashZone {clashZone.Id} (database only mode)");
                    }
                }
                
                return deleted;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[FLAG-MANAGER] Error deleting sleeve for intersection point change: {ex.Message}");
                throw;
            }
        }

        private static string GetSleeveCategory(FamilyInstance sleeve)
        {
            try
            {
                var param = sleeve?.LookupParameter("MEP_Category");
                if (param == null) return null;
                var value = param.AsString();
                return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
            }
            catch
            {
                return null;
            }
        }

        private static string GetClashZoneGuidValue(FamilyInstance sleeve)
        {
            try
            {
                var param = sleeve?.LookupParameter("ClashZone_GUID");
                if (param == null) return null;
                var value = param.AsString();
                return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
            }
            catch
            {
                return null;
            }
        }

        private static bool TryResolveSleeveIdFromGuid(
            string entryGuid,
            Dictionary<string, int> guidToSleeveId,
            Dictionary<int, string> sleeveIdToCategory,
            string category,
            out int resolvedSleeveId,
            out string resolvedCategory)
        {
            resolvedSleeveId = -1;
            resolvedCategory = null;

            if (string.IsNullOrWhiteSpace(entryGuid) || guidToSleeveId == null)
                return false;

            if (guidToSleeveId.TryGetValue(entryGuid, out var candidateId))
            {
                resolvedCategory = sleeveIdToCategory != null && sleeveIdToCategory.TryGetValue(candidateId, out var catValue)
                    ? catValue
                    : null;

                if (string.IsNullOrWhiteSpace(resolvedCategory) ||
                    string.Equals(resolvedCategory, category, StringComparison.OrdinalIgnoreCase))
                {
                    resolvedSleeveId = candidateId;
                    return true;
                }
            }

            return false;
        }
    }
}

