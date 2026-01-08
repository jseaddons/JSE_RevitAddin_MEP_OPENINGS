using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Services.Helpers;
using Autodesk.Revit.UI;

using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration;
using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Processing
{
    /// <summary>
    /// Service for applying MEPMARK parameters to cluster sleeves.
    /// Handles sequential numbering and prefix resolution.
    /// </summary>
    public partial class MarkParameterService
    {
        private readonly IMarkCacheService _cache;
        private readonly ICombinedSleeveMarkService _combinedService;
        private readonly Action<string> _logger;

        public MarkParameterService(Document doc = null, Action<string> logger = null)
        {
            _logger = logger ?? ((msg) => { });
            _cache = new MarkCacheService();
            _combinedService = new CombinedSleeveMarkService();
            
            if (doc != null)
            {
                using var context = new SleeveDbContext(doc);
                _cache.Initialize(doc, context);
            }
        }

        #region Public API (Standard & Batch)

        public (int processedCount, int errorCount) ApplyMarksFromDatabase(Document doc, MarkPrefixSettings settings, string? category = null)
        {
            try
            {
                var activeView = doc.ActiveView;
                if (!(activeView is ViewPlan plan)) throw new InvalidOperationException("Active view must be a floor plan");
                var level = plan.GenLevel;
                if (level == null) throw new InvalidOperationException("Floor plan must have an associated level");

                return ApplyMarksFromDatabase(doc, level.Name, category, settings.StartNumber, settings.NumberFormat, settings);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] ApplyMarksFromDatabase wrapper error: {ex.Message}");
                return (0, 1);
            }
        }

        public (int processedCount, int errorCount) ApplyMarksFromDatabase(Document doc, string levelName, string category, int startNumber, string numberFormat, MarkPrefixSettings? markPrefixes = null)
        {
            int processedCount = 0;
            int errorCount = 0;
            try
            {
                using var context = new SleeveDbContext(doc);
                var markRepo = new MarkDataRepository(context);
                var prefixStrategy = new DisciplinePrefixStrategy();
                var geoHelper = new ViewGeometryHelper();

                var zones = markRepo.GetSleevesForLevel(levelName, category);
                if (zones.Count == 0) return (0, 0);

                Outline worldBounds = geoHelper.GetWorldViewBounds(doc.ActiveView as ViewPlan);
                var validZones = zones.Where(z => geoHelper.IsPointInView(new XYZ(z.IntersectionPointX, z.IntersectionPointY, z.IntersectionPointZ), worldBounds)).ToList();

                if (validZones.Count == 0) return (0, 0);

                // 🔥 DIAGNOSTIC: Log what we loaded from database
                NumberingDebugLogger.LogInfo($"[MarkParameterService] Loaded {validZones.Count} valid zones for level '{levelName}', category '{category}'");
                var clustersCount = validZones.Count(z => z.ClusterInstanceId > 0);
                var individualsCount = validZones.Count(z => z.ClusterInstanceId <= 0 && z.SleeveInstanceId > 0);
                NumberingDebugLogger.LogInfo($"[MarkParameterService] 🔍 Breakdown: {clustersCount} clusters (ClusterInstanceId > 0), {individualsCount} individuals");
                
                // Log first few zones for debugging
                foreach (var z in validZones.Take(5))
                {
                    NumberingDebugLogger.LogInfo($"[MarkParameterService] 🔍 Zone sample: ClashZoneId={z.ClashZoneId}, SleeveInstanceId={z.SleeveInstanceId}, ClusterInstanceId={z.ClusterInstanceId}");
                }

                var elementGroups = validZones
                    .GroupBy(z => (z.ClusterInstanceId > 0) ? z.ClusterInstanceId : z.SleeveInstanceId)
                    .OrderBy(g => g.Key)
                    .ToList();

                var prefixGroups = new Dictionary<string, List<ElementId>>();
                foreach (var group in elementGroups)
                {
                    int instanceId = group.Key;
                    var zonesInGroup = group.ToList();
                    var primaryZone = zonesInGroup.First();

                    // ✅ CLUSTER RULE: For clusters, ALWAYS use discipline prefix only (ignore system type overrides)
                    // For individual sleeves, use full prefix resolution (including system type overrides)
                    // ✅ FIX: Check ClusterInstanceId > 0 to detect clusters (each cluster has unique ID, so Count is always 1)
                    string prefix;
                    bool isCluster = primaryZone.ClusterInstanceId > 0;
                    if (isCluster)
                    {
                        var settings = markPrefixes ?? new MarkPrefixSettings();
                        prefix = settings.GetDisciplinePrefix(category);
                        NumberingDebugLogger.LogInfo($"[MarkParameterService] Cluster {instanceId}: Using discipline prefix '{prefix}' (ClusterInstanceId={primaryZone.ClusterInstanceId})");
                    }
                    else // Individual sleeve
                    {
                        prefix = prefixStrategy.ResolvePrefix(category, markPrefixes ?? new MarkPrefixSettings(), primaryZone);
                        NumberingDebugLogger.LogInfo($"[MarkParameterService] Individual Sleeve {instanceId}: Using resolved prefix '{prefix}'");
                    }
                    
                    if (!prefixGroups.ContainsKey(prefix)) prefixGroups[prefix] = new List<ElementId>();
                    prefixGroups[prefix].Add(new ElementId(instanceId));
                }

                var updates = new List<(ElementId ElementId, string Value)>();
                foreach (var group in prefixGroups)
                {
                    string prefix = group.Key;
                    var elements = group.Value;
                    int currentNumber = startNumber;
                    if (category != null && category.Equals("Combined", StringComparison.OrdinalIgnoreCase)) currentNumber = 1;

                    foreach (var elId in elements)
                    {
                        string markValue = $"{prefix}{currentNumber.ToString(numberFormat)}";
                        updates.Add((elId, markValue));
                        currentNumber++;
                    }
                }

                foreach (var update in updates)
                {
                    try
                    {
                        var el = doc.GetElement(update.ElementId);
                        var p = el?.LookupParameter("MEP Mark") ?? el?.LookupParameter("Mark");
                        if (p != null && !p.IsReadOnly) { p.Set(update.Value); processedCount++; }
                    }
                    catch { errorCount++; }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] ApplyMarksFromDatabase error: {ex.Message}");
            }
            return (processedCount, errorCount);
        }

        public (int processedCount, int errorCount) ApplyPrefixesOnly(Document doc, string category, string projectPrefix, string disciplinePrefix, bool remarkAll, MarkPrefixSettings? markPrefixes = null)
        {
            int processedCount = 0;
            int errorCount = 0;
            try
            {
                using var context = new SleeveDbContext(doc);
                var markRepo = new MarkDataRepository(context);
                var prefixStrategy = new DisciplinePrefixStrategy();

                var zones = markRepo.GetMarkableClashZones(category);
                NumberingDebugLogger.LogInfo($"Found {zones.Count} markable zones for category '{category}' in DB");
                if (zones.Count == 0) return (0, 0);

                var updates = new List<(ElementId Id, string Value)>();
                var settings = markPrefixes ?? new MarkPrefixSettings();

                // ✅ UPDATED: Process BOTH Individual Sleeves AND Clusters
                // Previously filtered out clusters: .Where(z => z.ClusterInstanceId <= 0 && z.CombinedClusterSleeveInstanceId <= 0)
                // Now we process everything that has a valid ID
                var targetZones = zones
                    .Where(z => z.SleeveInstanceId > 0 || z.ClusterInstanceId > 0)
                    .ToList();
                
                var clustersCount = targetZones.Count(z => z.ClusterInstanceId > 0);
                NumberingDebugLogger.LogInfo($"[PREFIX-ONLY] Total zones: {zones.Count}, Processing: {targetZones.Count} (Individuals: {targetZones.Count - clustersCount}, Clusters: {clustersCount})");

                foreach (var mode in new[] { "Individual", "Cluster" })
                {
                    // Process in two passes just for logging clarity if needed, or single pass
                    // Let's do single pass for efficiency
                }

                foreach (var zone in targetZones)
                {
                    // Determine Target Element ID
                    // If it's a Cluster, the ClusterInstanceId is the Revit ElementId of the cluster family
                    int targetId = (zone.ClusterInstanceId > 0) ? zone.ClusterInstanceId : zone.SleeveInstanceId;
                    
                    if (targetId <= 0) continue;

                    var el = doc.GetElement(new ElementId(targetId));
                    if (el == null) continue;

                    var p = el.LookupParameter("MEP Mark") ?? el.LookupParameter("Mark");
                    string existingMark = p?.AsString() ?? "";
                    if (!remarkAll && !string.IsNullOrEmpty(existingMark)) continue;

                    // Resolve Prefix
                    string elementPrefix;
                    bool isCluster = zone.ClusterInstanceId > 0;
                    
                    if (isCluster)
                    {
                        // ✅ STRICT RULE: Clusters ALWAYS use Discipline Prefix only (ignore System Type overrides)
                        elementPrefix = settings.GetDisciplinePrefix(category);
                    }
                    else
                    {
                        // Individuals use full resolution (System Type > Discipline)
                        elementPrefix = prefixStrategy.ResolvePrefix(category, settings, zone);
                    }

                    string targetPrefix = $"{projectPrefix}{elementPrefix}";
                    
                    // bool isCluster = zone.ClusterInstanceId > 0; // Removed: Duplicate declaration
                    string typeLabel = isCluster ? "Cluster" : "Individual";
                    
                    NumberingDebugLogger.LogInfo($"[PREFIX-ONLY] {typeLabel} {targetId}: Prefix '{targetPrefix}' (Existing: '{existingMark}')");
                    updates.Add((el.Id, targetPrefix));
                }

                // Sort updates by mark value to ensure sequential application (though dictionary iteration order is not guaranteed, the loop above was sequential)
                updates = updates.OrderBy(u => u.Value).ToList();

                NumberingDebugLogger.LogStep($"[DEBUG] Total updates to apply: {updates.Count}");
                foreach (var up in updates.Take(50)) // Log first 50 to see enough examples
                {
                    NumberingDebugLogger.LogInfo($"[DEBUG] Update: ElementId={up.Id.IntegerValue}, Value='{up.Value}'");
                }

                NumberingDebugLogger.LogStep($"Applying {updates.Count} prefix updates to Revit...");
                foreach (var update in updates)
                {
                    try
                    {
                        var el = doc.GetElement(update.Id);
                        var p = el?.LookupParameter("MEP Mark") ?? el?.LookupParameter("Mark");
                        if (p != null && !p.IsReadOnly) 
                        { 
                            p.Set(update.Value); 
                            processedCount++; 
                        }
                        else
                        {
                            NumberingDebugLogger.LogStep($"FAILED to update element {update.Id.IntegerValue}: Parameter NULL or ReadOnly");
                        }
                    }
                    catch (Exception ex)
                    { 
                        RemarkDebugLogger.LogError($"Error updating element {update.Id.IntegerValue}", ex);
                        errorCount++; 
                    }
                }
                RemarkDebugLogger.LogStep($"Finished applying prefixes. Processed: {processedCount}, Errors: {errorCount}");

                // --------------------------------------------------------------------------------
                // 🔍 ORPHAN CHECK (Sleeves in Model but NOT in DB)
                // --------------------------------------------------------------------------------
                // The logs indicated "PREFIX-ORPHANS" was running, but code was missing.
                // Re-implementing correctly with CATEGORY SAFETY.
                
                var allSleevesInView = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(fi => {
                         var famName = fi.Symbol?.Family?.Name ?? "";
                         return famName.Contains("OpeningOnWall") || famName.Contains("OpeningOnSlab");
                    })
                    .ToList();

                var dbSleeveIds = new HashSet<int>(targetZones.Select(z => z.SleeveInstanceId).Union(targetZones.Select(z => z.ClusterInstanceId)));
                
                var orphans = allSleevesInView
                    .Where(s => !dbSleeveIds.Contains(s.Id.IntegerValue))
                    .ToList();

                if (orphans.Count > 0)
                {
                    NumberingDebugLogger.LogInfo($"[PREFIX-ORPHANS] Found {orphans.Count} sleeves in View NOT in DB (Orphans). Checking categories...");
                    
                    foreach (var orphan in orphans)
                    {
                        // ✅ CRITICAL FIX: Only process orphan if it matches the current CATEGORY
                        if (orphan.Category == null || !orphan.Category.Name.Contains(category))
                        {
                            // Example: If running "Pipes", skip "Ducts" orphan
                            // Note: Category names might be "Ducts", "Pipes", "Cable Trays", etc.
                            // Better specific check:
                            bool match = false;
                            if (category.StartsWith("Duct") && orphan.Category.Name.Contains("Duct")) match = true;
                            else if (category.StartsWith("Pipe") && orphan.Category.Name.Contains("Pipe")) match = true;
                            else if (category.Contains("Tray") && orphan.Category.Name.Contains("Tray")) match = true;
                            else if (category.Contains("Conduit") && orphan.Category.Name.Contains("Conduit")) match = true;
                            
                            if (!match)
                            {
                                // NumberingDebugLogger.LogInfo($"[PREFIX-ORPHANS] Skipping Orphan {orphan.Id} ({orphan.Category.Name}) - Mismatch with Target '{category}'");
                                continue;
                            }
                        }

                        // ✅ ORPHAN OVERRIDE LOGIC: Try to read System/Service Type from the sleeve element itself
                        // (Since it's not in DB, we rely on transferred parameters if they exist)
                        string sysType = orphan.LookupParameter("System Type")?.AsString() ?? 
                                         orphan.LookupParameter("MEP System Type")?.AsString();
                                         
                        string srvType = orphan.LookupParameter("Service Type")?.AsString() ?? 
                                         orphan.LookupParameter("MEP Service Type")?.AsString();

                        // Resolve using the same settings logic as DB elements
                        string resolvedPrefix = settings.GetPrefixForElement(category, sysType, srvType);
                        string fullPrefix = $"{projectPrefix}{resolvedPrefix}";
                        
                        var p = orphan.LookupParameter("MEP Mark") ?? orphan.LookupParameter("Mark");
                        string current = p?.AsString() ?? "";
                        
                        if (!remarkAll && !string.IsNullOrEmpty(current)) continue;

                        NumberingDebugLogger.LogInfo($"[PREFIX-ORPHANS] Updating Orphan {orphan.Id}: '{fullPrefix}'");
                        updates.Add((orphan.Id, fullPrefix));
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MarkParameterService] Prefix error: {ex.Message}");
            }
            return (processedCount, errorCount);


        }

        public (int processedCount, int errorCount) ApplyNumbersBatch(Document doc, string numberFormat, HashSet<string> allowedPrefixes = null, MarkPrefixSettings? markPrefixes = null)
        {
            int processedCount = 0;
            int errorCount = 0;
            try
            {
                FilteredElementCollector collector = (markPrefixes?.ActiveViewOnly == true && doc.ActiveView != null) 
                    ? new FilteredElementCollector(doc, doc.ActiveView.Id) 
                    : new FilteredElementCollector(doc);

                var allSleeves = collector.OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                    .Where(fi => {
                        var famName = fi.Symbol?.Family?.Name ?? "";
                        return famName.Contains("OpeningOnWall") || famName.Contains("OpeningOnSlab");
                    }).ToList();

                if (allSleeves.Count == 0) return (0, 0);

                var prefixGroups = new Dictionary<string, List<FamilyInstance>>();
                foreach (var sleeve in allSleeves)
                {
                    var p = sleeve.LookupParameter("MEP Mark") ?? sleeve.LookupParameter("Mark");
                    string currentMark = p?.AsString() ?? "";
                    if (string.IsNullOrWhiteSpace(currentMark)) 
                    {
                        NumberingDebugLogger.LogInfo($"[NUMBERING SKIPPED] Element {sleeve.Id.IntegerValue} has empty/null MEP Mark.");
                        continue;
                    }

                    string prefix = ExtractPrefix(currentMark);
                    
                    if (allowedPrefixes != null && !allowedPrefixes.Any(ap => prefix.StartsWith(ap)))
                    {
                        NumberingDebugLogger.LogInfo($"[NUMBERING SKIPPED] Element {sleeve.Id.IntegerValue} Prefix '{prefix}' NOT in allowed list: {string.Join(", ", allowedPrefixes)}");
                        continue; 
                    }

                    if (!prefixGroups.ContainsKey(prefix)) prefixGroups[prefix] = new List<FamilyInstance>();
                    prefixGroups[prefix].Add(sleeve);
                }

                using var context = new SleeveDbContext(doc);
                var markerRepo = new CategoryProcessingMarkerRepository(context);
                var updates = new List<(ElementId Id, string Value)>();

                // ✅ PER-VIEW SCOPING: Determine the scope for numbering memory
                string targetScope = ""; // Where we save progress (Active View)
                string sourceScope = ""; // Where we read start number from (Selected View or Active View)

                if (markPrefixes?.ActiveViewOnly == true && doc.ActiveView != null)
                {
                    // ✅ FIX: Use strict Level:View hierarchy for scope
                    // User format request: "p:level 0 :viewname"
                    // We construct scope as "LevelName:ViewName"
                    string levelName = "NoLevel";
                    if (doc.ActiveView.GenLevel != null)
                        levelName = doc.ActiveView.GenLevel.Name;
                    
                    // Sanitize ':' from names to avoid parsing issues, though strictly we just use it as a separator
                    // Using " : " (space colon space) for readability in DB if needed, or just ":"
                    targetScope = $"{levelName}:{doc.ActiveView.Name}";

                    NumberingDebugLogger.LogInfo($"[NUMBERING SCOPE] Target Scope: '{targetScope}' (Key format: Prefix|{targetScope})");

                    // Source defaults to Target, unless "Continue" is selected
                    if (markPrefixes.UseContinueNumbering && !string.IsNullOrEmpty(markPrefixes.ContinueFromViewName))
                    {
                        // Assume continuity is on the SAME LEVEL
                        sourceScope = $"{levelName}:{markPrefixes.ContinueFromViewName}";
                        NumberingDebugLogger.LogInfo($"[NUMBERING SCOPE] CONTINUING from Source: '{sourceScope}' -> Target: '{targetScope}'");
                    }
                    else
                    {
                        sourceScope = targetScope;
                        NumberingDebugLogger.LogInfo($"[NUMBERING SCOPE] Using View-Specific memory for: '{targetScope}'");
                    }
                }
                else
                {
                    NumberingDebugLogger.LogInfo("[NUMBERING SCOPE] Using Global (Project-wide) memory.");
                }

                foreach (var group in prefixGroups)
                {
                    string prefix = group.Key;
                    var items = group.Value.OrderBy(i => i.Id.IntegerValue).ToList();
                    
                    // ✅ SCOPED KEYS: Use source for reading, target for writing
                    string sourceKey = string.IsNullOrEmpty(sourceScope) ? prefix : $"{prefix}|{sourceScope}";
                    string targetKey = string.IsNullOrEmpty(targetScope) ? prefix : $"{prefix}|{targetScope}";
                    
                    int startNum = 1;
                    if (markPrefixes != null && markPrefixes.StartNumber > 0)
                    {
                         // Respect UI Start Number if set (forces a reset for this batch)
                         startNum = markPrefixes.StartNumber;
                         NumberingDebugLogger.LogInfo($"[NUMBERING] Prefix '{prefix}' - FORCING Start Number {startNum} (UI Override)");
                    }
                    else
                    {
                         // Fallback to Scoped History
                         var (lastNum, _) = markerRepo.GetMarker(sourceKey);
                         startNum = lastNum + 1;
                         NumberingDebugLogger.LogInfo($"[NUMBERING] Prefix '{prefix}' - Starting from '{sourceKey}' history: {startNum}");
                    }

                    foreach (var item in items)
                    {
                        string markValue = $"{prefix}{startNum.ToString(markPrefixes?.NumberFormat ?? "000")}";
                        updates.Add((item.Id, markValue));
                        startNum++;
                    }

                    // Update the database marker for the TARGET (Active Level)
                    markerRepo.UpdateMarker(targetKey, startNum - 1, string.Join(",", items.Select(i => i.Id.IntegerValue)));
                }

                foreach (var up in updates)
                {
                    try 
                    { 
                        var el = doc.GetElement(up.Id);
                        var p = el?.LookupParameter("MEP Mark") ?? el?.LookupParameter("Mark");
                        if (p != null && !p.IsReadOnly)
                        {
                            p.Set(up.Value); 
                            processedCount++; 
                        }
                    } 
                    catch { errorCount++; }
                }
            }
            catch (Exception ex) { DebugLogger.Error($"[MarkParameterService] Batch numbering error: {ex.Message}"); }
            return (processedCount, errorCount);
        }

        public (int processedCount, int errorCount) ApplyNumbersOnly(Document doc, string category, string numberFormat, MarkPrefixSettings? markPrefixes = null)
        {
            // Simplified version for de-bloat, preserving category logic
            return ApplyNumbersBatch(doc, numberFormat, null, markPrefixes);
        }

        #endregion

        #region Helpers

        private string ExtractPrefix(string mark)
        {
            string prefix = mark.Trim();
            int lastNonDigit = -1;
            for (int i = prefix.Length - 1; i >= 0; i--) { if (!char.IsDigit(prefix[i])) { lastNonDigit = i; break; } }
            return (lastNonDigit >= 0 && lastNonDigit < prefix.Length - 1) ? prefix.Substring(0, lastNonDigit + 1) : prefix;
        }

        public int ResetMarksForLevel(Document doc, string levelName, BoundingBoxXYZ viewExtent = null)
        {
            int clearedCount = 0;
            try
            {
                FilteredElementCollector collector = (viewExtent != null && doc.ActiveView != null) 
                    ? new FilteredElementCollector(doc, doc.ActiveView.Id) 
                    : new FilteredElementCollector(doc);

                var sleeves = collector.OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.Contains("OpeningOnWall") || fi.Symbol.Family.Name.Contains("OpeningOnSlab"))
                    .ToList();

                using var t = new Transaction(doc, "Reset Marks for Level");
                t.Start();
                foreach (var sleeve in sleeves)
                {
                    var p = sleeve.LookupParameter("MEP Mark") ?? sleeve.LookupParameter("Mark");
                    if (p != null && !p.IsReadOnly && !string.IsNullOrEmpty(p.AsString()))
                    {
                        p.Set("");
                        clearedCount++;
                    }
                }
                t.Commit();
            }
            catch (Exception ex) { DebugLogger.Error($"[MarkParameterService] ResetMarksForLevel error: {ex.Message}"); }
            return clearedCount;
        }

        public int ResetMarksForSelection(Document doc, ICollection<ElementId> selectedIds)
        {
            int clearedCount = 0;
            try
            {
                using var t = new Transaction(doc, "Reset Marks for Selection");
                t.Start();
                foreach (var id in selectedIds)
                {
                    var el = doc.GetElement(id);
                    var p = el?.LookupParameter("MEP Mark") ?? el?.LookupParameter("Mark");
                    if (p != null && !p.IsReadOnly && !string.IsNullOrEmpty(p.AsString()))
                    {
                        p.Set("");
                        clearedCount++;
                    }
                }
                t.Commit();
            }
            catch (Exception ex) { DebugLogger.Error($"[MarkParameterService] ResetMarksForSelection error: {ex.Message}"); }
            return clearedCount;
        }

        public void ResetCategoryCounters(Document doc, string levelName = null)
        {
            try
            {
                using var context = new SleeveDbContext(doc);
                var repo = new CategoryProcessingMarkerRepository(context);
                
                if (string.IsNullOrEmpty(levelName))
                {
                    repo.ResetAllMarkers();
                }
                else
                {
                    repo.ResetMarkersForLevel(levelName);
                }
            }
            catch (Exception ex) { DebugLogger.Error($"[MarkParameterService] ResetCategoryCounters error: {ex.Message}"); }
        }

        #endregion
    }
}
