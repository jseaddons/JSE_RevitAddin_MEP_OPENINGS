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

namespace JSE_RevitAddin_MEP_OPENINGS.Services
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
                    string prefix;
                    if (instanceId > 0 && zonesInGroup.Count > 1) // It's a cluster
                    {
                        var settings = markPrefixes ?? new MarkPrefixSettings();
                        prefix = settings.GetDisciplinePrefix(category);
                        RemarkDebugLogger.LogInfo($"[MarkParameterService] Cluster {instanceId}: Using discipline prefix '{prefix}' (system type overrides ignored for clusters)");
                    }
                    else // Individual sleeve
                    {
                        prefix = prefixStrategy.ResolvePrefix(category, markPrefixes ?? new MarkPrefixSettings(), primaryZone);
                        RemarkDebugLogger.LogInfo($"[MarkParameterService] Individual Sleeve {instanceId}: Using resolved prefix '{prefix}'");
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
                RemarkDebugLogger.LogInfo($"Found {zones.Count} markable zones for category '{category}' in DB");
                if (zones.Count == 0) return (0, 0);

                var updates = new List<(ElementId Id, string Value)>();
                var settings = markPrefixes ?? new MarkPrefixSettings();

                // ✅ REMARK SELECTED: Only process INDIVIDUAL sleeves (skip clusters and combined sleeves)
                var individualZones = zones
                    .Where(z => z.ClusterInstanceId <= 0 && z.CombinedClusterSleeveInstanceId <= 0)
                    .ToList();

                RemarkDebugLogger.LogInfo($"[PREFIX-ONLY] Total zones: {zones.Count}, Individual sleeves: {individualZones.Count}, Skipped (clusters/combined): {zones.Count - individualZones.Count}");

                foreach (var zone in individualZones)
                {
                    int targetId = zone.SleeveInstanceId;
                    if (targetId <= 0) continue;

                    var el = doc.GetElement(new ElementId(targetId));
                    if (el == null) continue;

                    var p = el.LookupParameter("MEP Mark") ?? el.LookupParameter("Mark");
                    string existingMark = p?.AsString() ?? "";
                    if (!remarkAll && !string.IsNullOrEmpty(existingMark)) continue;

                    // Individual sleeves use full prefix resolution (including system type overrides)
                    string elementPrefix = prefixStrategy.ResolvePrefix(category, settings, zone);
                    string targetPrefix = $"{projectPrefix}{elementPrefix}";
                    
                    RemarkDebugLogger.LogInfo($"[PREFIX-ONLY] Individual {targetId}: Prefix '{targetPrefix}' (Existing: '{existingMark}')");
                    updates.Add((el.Id, targetPrefix));
                }

                RemarkDebugLogger.LogStep($"Applying {updates.Count} prefix updates to Revit...");
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
                            RemarkDebugLogger.LogStep($"FAILED to update element {update.Id.IntegerValue}: Parameter NULL or ReadOnly");
                        }
                    }
                    catch (Exception ex)
                    { 
                        RemarkDebugLogger.LogError($"Error updating element {update.Id.IntegerValue}", ex);
                        errorCount++; 
                    }
                }
                RemarkDebugLogger.LogStep($"Finished applying prefixes. Processed: {processedCount}, Errors: {errorCount}");
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
                    if (string.IsNullOrWhiteSpace(currentMark)) continue;

                    string prefix = ExtractPrefix(currentMark);
                    if (allowedPrefixes != null && !allowedPrefixes.Any(ap => prefix.StartsWith(ap))) continue;

                    if (!prefixGroups.ContainsKey(prefix)) prefixGroups[prefix] = new List<FamilyInstance>();
                    prefixGroups[prefix].Add(sleeve);
                }

                using var context = new SleeveDbContext(doc);
                var markerRepo = new CategoryProcessingMarkerRepository(context);
                var updates = new List<(ElementId Id, string Value)>();

                foreach (var group in prefixGroups)
                {
                    string prefix = group.Key;
                    var items = group.Value.OrderBy(i => i.Id.IntegerValue).ToList();
                    var (lastNum, _) = markerRepo.GetMarker(prefix);
                    int startNum = lastNum + 1;

                    foreach (var item in items)
                    {
                        string newVal = $"{prefix}{startNum.ToString(numberFormat)}";
                        updates.Add((item.Id, newVal));
                        startNum++;
                    }
                    markerRepo.UpdateMarker(prefix, startNum - 1);
                }

                foreach (var up in updates)
                {
                    try { doc.GetElement(up.Id)?.LookupParameter("MEP Mark")?.Set(up.Value); processedCount++; } catch { errorCount++; }
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

        public void ResetCategoryCounters(Document doc)
        {
            try
            {
                using var context = new SleeveDbContext(doc);
                var repo = new CategoryProcessingMarkerRepository(context);
                repo.ResetAllMarkers();
            }
            catch (Exception ex) { DebugLogger.Error($"[MarkParameterService] ResetCategoryCounters error: {ex.Message}"); }
        }

        #endregion
    }
}
