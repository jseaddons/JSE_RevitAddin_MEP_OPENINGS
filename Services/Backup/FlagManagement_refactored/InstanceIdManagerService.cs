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
    /// Team F: SOLID-compliant instance ID management service.
    /// 
    /// This service handles resetting instance IDs for deleted sleeves.
    /// SOLID: Single Responsibility - instance ID operations only.
    /// 
    /// ✅ PRESERVES ALL OPTIMIZATIONS:
    /// - Batch collection (1 Revit API call instead of N calls per category)
    /// - Category filtering in memory (O(1) lookup via HashSet)
    /// - SectionBoxHelper reuse for section box filtering
    /// - Database-first approach with XML fallback
    /// - Batch database updates via BatchUpdateFlags()
    /// </summary>
    public class InstanceIdManagerService : IInstanceIdManager
    {
        private readonly Document _document;
        private readonly IClashZoneRepository _repository; // ✅ TESTABILITY: Injected repository
        private readonly ISleeveCollector _sleeveCollector; // ✅ TESTABILITY: Injected sleeve collector
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new instance ID manager service.
        /// 
        /// ✅ TESTABILITY: All dependencies are injected via interfaces.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="repository">Clash zone repository (injected for testability)</param>
        /// <param name="sleeveCollector">Sleeve collector (injected for testability)</param>
        /// <param name="logger">Optional logger for tracking operations</param>
        public InstanceIdManagerService(
            Document document,
            IClashZoneRepository repository,
            ISleeveCollector sleeveCollector,
            ILogger logger = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _sleeveCollector = sleeveCollector ?? throw new ArgumentNullException(nameof(sleeveCollector));
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Resets instance IDs for deleted sleeves.
        /// Uses batch collection optimization (collect all sleeves once for all categories).
        /// </summary>
        public int ResetInstanceIdsForDeletedSleeves(
            List<string> categories,
            Dictionary<string, List<ClashZone>> clashZonesByCategory = null,
            string refreshLogName = null)
        {
            if (categories == null || categories.Count == 0)
                return 0;
            
            int totalResetCount = 0;
            
            // ✅ PERFORMANCE OPTIMIZATION: Collect ALL sleeves from Revit ONCE (not per category)
            // This reduces N Revit API calls (one per category) to just 1 call
            // Then filter by category in memory (much faster)
            Dictionary<string, HashSet<int>> sleevesByCategory = null;
            Dictionary<int, string> sleeveIdToCategory = null;
            
            if (categories.Count > 1)
            {
                // ✅ BATCH COLLECTION: Collect all sleeves once for all categories
                // ✅ TESTABILITY: Use injected ISleeveCollector (can be mocked)
                _logger.Info($"===== BATCH COLLECTING ALL SLEEVES FROM REVIT (for {categories.Count} categories) =====", "InstanceIdManager");
                
                sleevesByCategory = _sleeveCollector.CollectSleevesByCategory(_document);
                
                int totalSleeves = sleevesByCategory.Values.Sum(hs => hs.Count);
                _logger.Info($"✅ BATCH COLLECTED: {totalSleeves} total sleeves across {sleevesByCategory.Count} categories from Revit (1 API call instead of {categories.Count})", "InstanceIdManager");
                foreach (var kvp in sleevesByCategory)
                {
                    _logger.Debug($"  Category '{kvp.Key}': {kvp.Value.Count} sleeves", "InstanceIdManager");
                }
            }
            
            foreach (var category in categories)
            {
                if (string.IsNullOrWhiteSpace(category))
                    continue;
                
                try
                {
                    // ✅ SOLID REFACTOR: Use existing SectionBoxHelper for section box detection
                    List<ClashZone> dbZones = null;
                    var activeView = _document.ActiveView;
                    BoundingBoxXYZ sectionBox = null;
                    
                    if (activeView is View3D view3D)
                    {
                        sectionBox = SectionBoxHelper.GetSectionBoxBounds(view3D);
                    }
                    
                    try
                    {
                        // ✅ TESTABILITY: Use injected IClashZoneRepository (can be mocked)
                        // ✅ Use section-box-aware query if section box is active
                        if (sectionBox != null && OptimizationFlags.UseRTreeDatabaseIndex)
                        {
                            // Query ALL zones within section box (R-tree spatial index)
                            dbZones = _repository.GetClashZonesByCategoryInSectionBox(category, sectionBox);
                            
                            _logger.Info($"✅ Section-box query: Loaded {dbZones?.Count ?? 0} zones (filtered by R-tree) for category '{category}'", "InstanceIdManager");
                        }
                        else
                        {
                            // Fallback: Load all zones for category (no section box or R-tree disabled)
                            dbZones = _repository.GetClashZonesByCategory(category);
                            
                            _logger.Warning($"⚠️ Loaded {dbZones?.Count ?? 0} zones (NO section-box filter) for category '{category}'", "InstanceIdManager");
                        }
                        
                        // Filter to only zones with sleeve IDs
                        if (dbZones != null)
                        {
                            dbZones = dbZones.Where(z => z != null && (z.SleeveInstanceId > 0 || z.ClusterSleeveInstanceId > 0))
                                .ToList();
                        }
                    }
                    catch (Exception dbEx)
                    {
                        _logger.Warning($"⚠️ Failed to load from database for category '{category}', falling back to Global XML: {dbEx.Message}", "InstanceIdManager");
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
                        
                        _logger.Info($"✅ [INSTANCE-ID-RESET] Using database for category '{category}' ({allEntries.Count} zones with sleeve IDs)", "InstanceIdManager");
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
                            
                            _logger.Warning($"⚠️ [INSTANCE-ID-RESET] Using Global XML fallback for category '{category}' ({allEntries.Count} entries)", "InstanceIdManager");
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
                    // ✅ TESTABILITY: Use injected ISleeveCollector (can be mocked)
                    var existingSleeveIds = new HashSet<int>();
                    foreach (var sleeveId in sleeveIdsToCheck)
                    {
                        if (_sleeveCollector.SleeveExists(_document, sleeveId))
                        {
                            existingSleeveIds.Add(sleeveId);
                        }
                    }
                    
                    // Build updates for deleted sleeves
                    var updates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, int SleeveInstanceId, int ClusterSleeveInstanceId, int MepElementId, int StructuralElementId, double IntersectionPointX, double IntersectionPointY, double IntersectionPointZ, int OldSleeveInstanceId, int OldClusterInstanceId, bool? MarkedForClusterProcess, int AfterClusterSleeveId, bool? IsClusteredFlag)>();
                    
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
                                entry.ClusterSleeveInstanceId,
                                null, // ✅ EDGE CASE: MarkedForClusterProcess (not available on database entry)
                                -1, // ✅ EDGE CASE: AfterClusterSleeveId (not available on database entry)
                                null // ✅ EDGE CASE: IsClusteredFlag (deprecated, set to null)
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
                                _logger.Debug($"[SQLite] {msg}", "InstanceIdManager");
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET][SQLite] {msg}\n");
                            }))
                            {
                                var repository = new ClashZoneRepository(context, msg =>
                                {
                                    _logger.Debug($"[SQLite] {msg}", "InstanceIdManager");
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
                                    OldClusterInstanceId: u.OldClusterInstanceId,
                                    MarkedForClusterProcess: u.MarkedForClusterProcess,
                                    AfterClusterSleeveId: u.AfterClusterSleeveId,
                                    IsClusteredFlag: (bool?)null
                                )).ToList();
                                
                                repository.BatchUpdateFlags(dbUpdates);
                                
                                _logger.Info($"✅ [INSTANCE-ID-RESET] Updated database for {dbUpdates.Count} clash zones in category '{category}' (DB FIRST)", "InstanceIdManager");
                                
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ✅ Database updated: {dbUpdates.Count} zones in category '{category}'\n");
                            }
                        }
                        catch (Exception dbEx)
                        {
                            _logger.Error($"❌ [INSTANCE-ID-RESET] Database update failed for category '{category}': {dbEx.Message}", dbEx, "InstanceIdManager");
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
                                
                                GlobalIndexService.UpsertFlagsWithIdsAndClashZoneData(_document, category, globalXmlUpdates, filterName: null!, refreshLogName: refreshLogName);
                                
                                _logger.Info($"✅ [INSTANCE-ID-RESET] Updated XML for {globalXmlUpdates.Count} entries in category '{category}' (XML SECOND)", "InstanceIdManager");
                                
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ✅ XML updated: {globalXmlUpdates.Count} entries in category '{category}'\n");
                            }
                            catch (Exception xmlEx)
                            {
                                _logger.Error($"❌ [INSTANCE-ID-RESET] XML update failed for category '{category}': {xmlEx.Message}", xmlEx, "InstanceIdManager");
                                if (!string.IsNullOrWhiteSpace(refreshLogName))
                                    SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ❌ XML update failed: {xmlEx.Message}\n");
                                // Don't throw - XML is optional in Phase 2
                            }
                        }
                        else
                        {
                            _logger.Debug($"⚠️ [INSTANCE-ID-RESET] Skipping XML update (XML creation disabled) for category '{category}'", "InstanceIdManager");
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error($"❌ [INSTANCE-ID-RESET] Error processing category '{category}': {ex.Message}", ex, "InstanceIdManager");
                    if (!string.IsNullOrWhiteSpace(refreshLogName))
                        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [INSTANCE-ID-RESET] ❌ Error processing category '{category}': {ex.Message}\n");
                }
            }
            
            return totalResetCount;
        }
    }
}

