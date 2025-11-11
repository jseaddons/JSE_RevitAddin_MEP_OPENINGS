using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
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
        /// ✅ GLOBAL XML FLAG MANAGEMENT: Syncs flags from Global XML (single source of truth) to in-memory clash zones.
        /// Global XML is the authoritative source - flags are NOT stored in Filter XML files.
        /// Called during refresh to ensure in-memory clash zones reflect the current state from Global XML.
        /// </summary>
        /// <param name="clashZones">List of in-memory clash zones to sync (loaded from Filter XML)</param>
        /// <param name="category">MEP element category name</param>
        public void SyncFlagsFromGlobal(List<ClashZone> clashZones, string category)
        {
            if (clashZones == null || clashZones.Count == 0)
                return;
                
            if (string.IsNullOrWhiteSpace(category))
                return;
            
            try
            {
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
                        // ✅ SYNC: Copy flags FROM Global XML (on disk) TO in-memory clash zones (loaded from Filter XML)
                        // Filter XML does NOT store flags ([XmlIgnore]) - flags default to false when loaded
                        // This sync ensures in-memory clash zones have correct flag values from Global XML (single source of truth)
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
                        
                        // ✅ CRITICAL FIX: Always sync SleeveInstanceId from Global XML (authoritative source)
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
                        
                        // ✅ CRITICAL FIX: Always sync ClusterSleeveInstanceId from Global XML (authoritative source)
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
            if (categories == null || categories.Count == 0)
                return 0;
            
            int totalResetCount = 0;
            
            try
            {
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
                        
                        // ✅ OPTIMIZATION: Pre-collect all sleeve IDs by category (calculate once, use many times)
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
                                    string sleeveCategory = mepCategoryParam.AsString().Trim();
                                    // ✅ CRITICAL: Exact case-insensitive match required - no fallback
                                    return string.Equals(sleeveCategory, category, StringComparison.OrdinalIgnoreCase);
                                }
                                
                                // ✅ CRITICAL: Sleeve without MEP_Category parameter cannot be matched to category - exclude it
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[FLAG-MANAGER] Sleeve {s.Id.IntegerValue}: Missing MEP_Category parameter - cannot match to category '{category}', excluding from reset check");
                                }
                                
                                return false;
                            })
                            .ToList();
                        
                        // ✅ OPTIMIZATION: Build HashSet for O(1) lookup (calculate once, use many times)
                        var existingSleeveIdsSet = new HashSet<int>(allSleeves.Select(s => s.Id.IntegerValue));

                        void LogToRefresh(string message)
                        {
                            if (string.IsNullOrWhiteSpace(refreshLogName))
                                return;

                            SafeFileLogger.SafeAppendText(refreshLogName,
                                $"[{DateTime.Now}] [FLAG-MANAGER] {message}\n");
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
                                var allSleevesWithoutCategoryFilter = new FilteredElementCollector(_document)
                                    .OfClass(typeof(FamilyInstance))
                                    .Cast<FamilyInstance>()
                                    .Where(s => 
                                    {
                                        bool hasSleeveKeyword = s.Symbol?.FamilyName?.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) == true ||
                                                               s.Symbol?.FamilyName?.Contains("Opening", StringComparison.OrdinalIgnoreCase) == true;
                                        string familyName = s.Symbol?.FamilyName ?? "";
                                        bool isKnownFamily = familyName.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) ||
                                                            familyName.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                                        return (s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") &&
                                               (hasSleeveKeyword || isKnownFamily);
                                    })
                                    .ToList();
                                
                                DebugLogger.Info($"[FLAG-MANAGER] Found {allSleevesWithoutCategoryFilter.Count} total sleeves (without category filter)");
                                if (allSleevesWithoutCategoryFilter.Count > 0)
                                {
                                    var categoryBreakdown = allSleevesWithoutCategoryFilter
                                        .GroupBy(s => 
                                        {
                                            var param = s.LookupParameter("MEP_Category");
                                            return param != null && !string.IsNullOrWhiteSpace(param.AsString()) 
                                                ? param.AsString().Trim() 
                                                : "NO_MEP_CATEGORY";
                                        })
                                        .ToDictionary(g => g.Key, g => g.Count());
                                    
                                var breakdown = string.Join(", ", categoryBreakdown.Select(kvp => $"{kvp.Key}={kvp.Value}"));
                                DebugLogger.Info($"[FLAG-MANAGER] Sleeve category breakdown: {breakdown}");
                                LogToRefresh($"Sleeve category breakdown (all sleeves detected): {breakdown}");
                                }
                            }
                        }
                        
                        var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ)>();
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
                                // ✅ OPTIMIZATION: Use pre-collected HashSet for O(1) lookup instead of Revit API call
                                bool clusterSleeveExists = existingSleeveIdsSet.Contains(globalEntry.ClusterSleeveInstanceId);
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER]   Entry {globalEntry.Id}: Cluster sleeve {globalEntry.ClusterSleeveInstanceId} exists in Revit: {clusterSleeveExists}");
                                if (detailLogBudget-- > 0)
                                    LogToRefresh($"Entry {globalEntry.Id}: ClusterSleeveId={globalEntry.ClusterSleeveInstanceId}, Exists={clusterSleeveExists}, IsClusterResolved={globalEntry.IsClusterResolved}");
                                
                                if (!clusterSleeveExists)
                                {
                                    // Cluster sleeve deleted → Reset ALL flags
                                    updates.Add((Guid.Parse(globalEntry.Id), false, false, -1, -1,
                                                globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ));
                                    resetCount++;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset ALL flags for Global XML entry {globalEntry.Id} - cluster sleeve {globalEntry.ClusterSleeveInstanceId} NOT FOUND");
                                    LogToRefresh($"RESET cluster entry {globalEntry.Id}: ClusterSleeveId={globalEntry.ClusterSleeveInstanceId} missing → Flags cleared.");
                                }
                                continue; // Skip individual check if cluster was checked
                            }
                            
                            // ✅ PROTECTED LOGIC: Then check individual (only if cluster flag is false)
                            if (globalEntry.IsResolved && globalEntry.SleeveInstanceId > 0)
                            {
                                // ✅ OPTIMIZATION: Use pre-collected HashSet for O(1) lookup instead of Revit API call
                                bool individualSleeveExists = existingSleeveIdsSet.Contains(globalEntry.SleeveInstanceId);
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Info($"[FLAG-MANAGER]   Entry {globalEntry.Id}: Individual sleeve {globalEntry.SleeveInstanceId} exists in Revit: {individualSleeveExists}");
                                if (detailLogBudget-- > 0)
                                    LogToRefresh($"Entry {globalEntry.Id}: SleeveId={globalEntry.SleeveInstanceId}, Exists={individualSleeveExists}, IsResolved={globalEntry.IsResolved}");
                                
                                if (!individualSleeveExists)
                                {
                                    // Individual sleeve deleted → Reset individual flag only
                                    updates.Add((Guid.Parse(globalEntry.Id), false, globalEntry.IsClusterResolved, -1, globalEntry.ClusterSleeveInstanceId,
                                                globalEntry.MepElementId, globalEntry.StructuralElementId,
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ));
                                    resetCount++;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[FLAG-MANAGER] ✓ Reset individual flag for Global XML entry {globalEntry.Id} - individual sleeve {globalEntry.SleeveInstanceId} NOT FOUND");
                                    LogToRefresh($"RESET individual entry {globalEntry.Id}: SleeveId={globalEntry.SleeveInstanceId} missing → IsResolved=false, SleeveInstanceId=-1");
                                }
                            }
                        }
                        
                        // Save updated flags to Global XML
                        LogToRefresh($"Category '{category}' scan complete → updates={updates.Count}, resetCount={resetCount}");

                        if (updates.Count > 0)
                        {
                            // ✅ CRITICAL FIX: FilterName is preserved from Global XML entry (not overwritten)
                            GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, updates, filterName: null);
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
                        // ✅ CRITICAL FIX: Try to preserve FilterName from Global XML entries when resetting flags
                        // If FilterName exists in Global XML, use it; otherwise leave empty
                        GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, updatesWithData, filterName: null);
                    }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] Updated Global XML for {updates.Count} clash zones with reset flags in category '{category}'");
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
                        
                        var updates = new List<(Guid Id, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ)>();
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
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ));
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
                                                globalEntry.IntersectionPointX, globalEntry.IntersectionPointY, globalEntry.IntersectionPointZ));
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
                            GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, updates, filterName: null);
                            totalResetCount += resetCount;
                            
                            if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Info($"[FLAG-MANAGER] ✅ Reset {resetCount} Global XML entries in category '{category}' - sleeves were deleted");
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
                
                // Update Global XML using existing service with MEP+Host+Point data
                // ✅ CRITICAL FIX: Pass FilterName to ensure Global XML knows which Filter XML file contains placement data
                int mepId = clashZone.MepElementId?.IntegerValue ?? clashZone.MepElementIdValue;
                int hostId = clashZone.StructuralElementId?.IntegerValue ?? clashZone.StructuralElementIdValue;
                
                GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, new[] { 
                    (clashZone.Id, clashZone.IsResolved, clashZone.IsClusterResolved, 
                     clashZone.SleeveInstanceId, clashZone.ClusterSleeveInstanceId,
                     mepId, hostId, clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ) 
                }, filterName);
                
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[FLAG-MANAGER] Updated flags for ClashZone {clashZone.Id} after {(isCluster ? "cluster" : "individual")} sleeve placement (SleeveId={sleeveId}, Filter='{filterName ?? "N/A"}')");
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
                    GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, updatesToExisting, filterName: null);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] ✅ Updated {updatedCount} existing entries and created {createdCount} new entries with sleeve IDs from Revit for category '{category}'");
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
                    
                    // ✅ CRITICAL FIX: FilterName not available in this context, but flag reset will preserve existing FilterName
                    GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, new[] { 
                        (clashZone.Id, false, false, -1, -1,
                         mepId, hostId, clashZone.IntersectionPointX, clashZone.IntersectionPointY, clashZone.IntersectionPointZ) 
                    }, filterName: null);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[FLAG-MANAGER] ✅ Reset flags for ClashZone {clashZone.Id} after sleeve deletion due to intersection point change");
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
    }
}

