using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ DEDICATED OOP SERVICE: Saves clash zones to both Global XML and Filter XML
    /// Ensures data consistency by using the same clash zone objects for both
    /// Handles tree structure correctly for Global XML
    /// Called after intersection detection completes
    /// </summary>
    public class ClashZonePersistenceService
    {
        private readonly Document _document;
        private readonly GuidManager _guidManager;
        private readonly string _refreshLogName;

        public ClashZonePersistenceService(Document document, GuidManager guidManager, string refreshLogName = null)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _guidManager = guidManager ?? throw new ArgumentNullException(nameof(guidManager));
            _refreshLogName = refreshLogName ?? "refresh.log";
        }

        /// <summary>
        /// ✅ UNIFIED METHOD: Save clash zones to BOTH Global XML and Filter XML using SAME tree structure
        /// ONE method handles everything - groups by category, then by file combo, saves to both XML files
        /// </summary>
        /// <param name="allClashZones">All clash zones (new + existing merged)</param>
        /// <param name="baseFilterName">Base filter name (e.g., "Plumbing")</param>
        /// <param name="targetFilter">Target filter for Filter XML saving</param>
        /// <param name="existingClashZones">Existing clash zones from Filter XML (for merge logic)</param>
        public void SaveClashZones(
            List<ClashZone> allClashZones,
            string baseFilterName,
            OpeningFilter targetFilter,
            bool allowStructuralUpdates)
        {
            if (allClashZones == null || allClashZones.Count == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[CLASH-ZONE-PERSISTENCE] No clash zones to save");
                return;
            }

            var processingSummaries = new List<ProcessingStats>();
            baseFilterName = NormalizeBaseFilterName(baseFilterName);

            if (targetFilter != null)
            {
                EnsureFilterStorageInitialized(targetFilter);
                DeduplicateFilterStorage(targetFilter);
            }

            try
            {
                LogPlacement($"[PERSIST-ENTRY] Zones={allClashZones.Count}, Filter='{baseFilterName}', Sample=[{string.Join(", ", allClashZones.Where(z => z != null).Take(10).Select(z => $"{z.Id}:{z.SleeveInstanceId}"))}]");

                var clashZonesByCategory = allClashZones
                    .Where(z => z != null && !string.IsNullOrWhiteSpace(z.MepElementCategory))
                    .GroupBy(z => z.MepElementCategory, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[CLASH-ZONE-PERSISTENCE] Saving {allClashZones.Count} clash zones across {clashZonesByCategory.Count} categories");

                foreach (var categoryGroup in clashZonesByCategory)
                {
                    var stats = SaveCategory(
                        categoryGroup.Key,
                        categoryGroup.ToList(),
                        baseFilterName,
                        targetFilter,
                        allowStructuralUpdates);

                    processingSummaries.Add(stats);
                }

                LogAggregate(processingSummaries, baseFilterName);
            }
            catch (Exception ex)
            {
                HandleException("[CLASH-ZONE-PERSISTENCE] Error saving clash zones", ex);
                throw;
            }
        }

        /// <summary>
        /// ✅ UNIFIED SAVE METHOD: Saves to BOTH Global XML and Filter XML using SAME tree structure
        /// ONE method handles everything - groups by file combo, saves flags to Global XML, saves placement data to Filter XML
        /// </summary>
        private ProcessingStats SaveCategory(
            string category,
            List<ClashZone> categoryClashZones,
            string baseFilterName,
            OpeningFilter targetFilter,
            bool allowStructuralUpdates)
        {
            var stats = new ProcessingStats(category);

            try
            {
                var totalZones = categoryClashZones?.Count ?? 0;
                LogPlacement($"[PERSIST-CATEGORY] Category='{category}', Filter='{baseFilterName}', Zones={totalZones}");

                var validZones = categoryClashZones?
                    .Where(IsValidClashZone)
                    .ToList() ?? new List<ClashZone>();

                stats.TotalZones = totalZones;
                stats.ValidZones = validZones.Count;

                var invalidCount = totalZones - stats.ValidZones;
                if (invalidCount > 0)
                {
                    stats.InvalidZones = invalidCount;
                    LogRefresh($"[PERSIST-WARNING] Category='{category}' ignored {invalidCount} invalid clash zones");
                }

                if (validZones.Count == 0)
                {
                    LogRefresh($"[PERSIST-INFO] Category='{category}' has no valid clash zones to persist");
                    return stats;
                }

                // ✅ DEBUG: Log file combo keys before grouping to diagnose missing combos
                LogRefresh($"[PERSIST-DEBUG] Analyzing {validZones.Count} valid zones for file combo grouping");
                var comboKeysSample = validZones.Take(10).Select(cz => GetFileComboKey(cz)).ToList();
                foreach (var key in comboKeysSample)
                {
                    LogRefresh($"[PERSIST-DEBUG]   Sample combo key: Linked='{key.LinkedFile}', Host='{key.HostFile}'");
                }

                var combos = validZones
                    .GroupBy(GetFileComboKey)
                    .Where(g => !string.IsNullOrWhiteSpace(g.Key.LinkedFile) && !string.IsNullOrWhiteSpace(g.Key.HostFile))
                    .ToList();

                stats.FileComboCount = combos.Count;
                LogRefresh($"[PERSIST-DEBUG] Valid combos to persist for '{category}': {combos.Count}");
                foreach (var combo in combos)
                {
                    LogRefresh($"[PERSIST-DEBUG]   Combo: Linked='{combo.Key.LinkedFile}', Host='{combo.Key.HostFile}', Zones={combo.Count()}");
                }

                var filterName = BuildFilterFileName(baseFilterName, category);
                var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);

                foreach (var comboGroup in combos)
                {
                    var key = comboGroup.Key;
                    var comboClashZones = comboGroup.ToList();

                    LogRefresh($"[PERSIST-DEBUG]   Combo → Linked='{key.LinkedFile}', Host='{key.HostFile}', Zones={comboClashZones.Count}");
                    LogPlacement($"[PERSIST-COMBO] Linked='{key.LinkedFile}', Host='{key.HostFile}', Count={comboClashZones.Count}, Sample=[{string.Join(", ", comboClashZones.Take(5).Select(z => $"{z.Id}:{z.SleeveInstanceId}"))}]");

                    SaveToGlobalXml(globalIndex, comboClashZones, category, baseFilterName, filterName, key, stats, allowStructuralUpdates);

                    if (targetFilter != null)
                    {
                        SaveToFilterXml(comboClashZones, category, baseFilterName, targetFilter, key, stats, allowStructuralUpdates);
                    }
                }

                CleanupGlobalIndex(globalIndex, validZones, stats);
                GlobalIndexService.Save(_document, globalIndex);
            }
            catch (Exception ex)
            {
                HandleException($"[PERSIST-ERROR] Category='{category}'", ex);
                throw;
            }

            return stats;
        }

        private void SaveToGlobalXml(
            CategoryGlobalIndex globalIndex,
            List<ClashZone> comboClashZones,
            string category,
            string baseFilterName,
            string filterFileName,
            (string LinkedFile, string HostFile) comboKey,
            ProcessingStats stats,
            bool allowStructuralUpdates)
        {
            if (comboClashZones == null || comboClashZones.Count == 0)
            {
                LogRefresh($"[PERSIST-GLOBAL] Skipping SaveToGlobalXml - no clash zones for combo Linked='{comboKey.LinkedFile}', Host='{comboKey.HostFile}'");
                return;
            }

            if (globalIndex.Filters == null)
                globalIndex.Filters = new List<FilterGroup>();

            var globalFilterGroup = globalIndex.Filters
                .FirstOrDefault(f => string.Equals(f?.Name, baseFilterName, StringComparison.OrdinalIgnoreCase));

            if (globalFilterGroup == null)
            {
                globalFilterGroup = new FilterGroup
                {
                    Name = baseFilterName,
                    FileCombos = new List<FileComboGroup>()
                };
                globalIndex.Filters.Add(globalFilterGroup);
                LogRefresh($"[PERSIST-GLOBAL] Created new FilterGroup '{baseFilterName}' for category '{category}'");
            }

            globalFilterGroup.FileCombos ??= new List<FileComboGroup>();

            var processedCombo = new ProcessedFileCombo
            {
                LinkedFile = comboKey.LinkedFile,
                HostFile = comboKey.HostFile
            };
            var normalizedKey = processedCombo.GetNormalizedKey();

            LogRefresh($"[PERSIST-GLOBAL] Looking for FileComboGroup with normalized key '{normalizedKey}' (Linked='{comboKey.LinkedFile}', Host='{comboKey.HostFile}') in FilterGroup '{baseFilterName}'");
            LogRefresh($"[PERSIST-GLOBAL] Existing FileComboGroups in FilterGroup '{baseFilterName}': {globalFilterGroup.FileCombos.Count}");
            foreach (var existingCombo in globalFilterGroup.FileCombos)
            {
                LogRefresh($"[PERSIST-GLOBAL]   Existing combo: Linked='{existingCombo.LinkedFile}', Host='{existingCombo.HostFile}', NormalizedKey='{existingCombo.GetNormalizedKey()}'");
            }

            var globalFileCombo = globalFilterGroup.FileCombos
                .FirstOrDefault(fc => fc.GetNormalizedKey() == normalizedKey);

            if (globalFileCombo == null)
            {
                globalFileCombo = new FileComboGroup
                {
                    LinkedFile = comboKey.LinkedFile,
                    HostFile = comboKey.HostFile,
                    ProcessedAt = DateTime.Now,
                    IsProcessed = true,
                    Entries = new List<CategoryGlobalIndexEntry>()
                };
                globalFilterGroup.FileCombos.Add(globalFileCombo);
                LogRefresh($"[PERSIST-GLOBAL] ✅ Created NEW FileComboGroup: Linked='{comboKey.LinkedFile}', Host='{comboKey.HostFile}', NormalizedKey='{normalizedKey}', Zones={comboClashZones.Count}");
            }
            else
            {
                LogRefresh($"[PERSIST-GLOBAL] ✅ Found EXISTING FileComboGroup: Linked='{comboKey.LinkedFile}', Host='{comboKey.HostFile}', NormalizedKey='{normalizedKey}', Zones={comboClashZones.Count}");
            }

            globalFileCombo.Entries ??= new List<CategoryGlobalIndexEntry>();

            var entriesById = globalFileCombo.Entries
                .Where(e => e != null && !string.IsNullOrWhiteSpace(e.Id))
                .GroupBy(e => e.Id, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToDictionary(e => e.Id, StringComparer.OrdinalIgnoreCase);

            globalFileCombo.Entries = entriesById.Values.ToList();

            foreach (var clashZone in comboClashZones)
            {
                var entryId = clashZone.Id.ToString();
                var existingEntry = GlobalIndexService.GetAllEntries(globalIndex)
                    .FirstOrDefault(e => string.Equals(e.Id, entryId, StringComparison.OrdinalIgnoreCase));

                if (existingEntry == null ||
                    existingEntry.MepElementId == 0 ||
                    existingEntry.StructuralElementId == 0 ||
                    IsZeroIntersection(existingEntry))
                {
                    _guidManager.EnsureGlobalXmlEntry(clashZone, category, filterFileName);

                    if (existingEntry == null)
                        stats.GlobalCreated++;
                    else
                        stats.GlobalUpdated++;
                }

                if (!entriesById.TryGetValue(entryId, out var comboEntry))
                {
                    comboEntry = new CategoryGlobalIndexEntry
                    {
                        Id = entryId,
                        FilterName = filterFileName ?? string.Empty
                    };
                    globalFileCombo.Entries.Add(comboEntry);
                    entriesById[entryId] = comboEntry;
                }

                UpdateGlobalEntry(comboEntry, clashZone, filterFileName, allowStructuralUpdates);
            }
        }

        private void SaveToFilterXml(
            List<ClashZone> comboClashZones,
            string category,
            string baseFilterName,
            OpeningFilter targetFilter,
            (string LinkedFile, string HostFile) comboKey,
            ProcessingStats stats,
            bool allowStructuralUpdates)
        {
            if (comboClashZones == null || comboClashZones.Count == 0 || targetFilter == null)
                return;

            targetFilter.ClashZoneStorage ??= new ClashZoneStorage
            {
                CreatedAt = DateTime.Now,
                LastUpdated = DateTime.Now,
                DocumentPath = _document.PathName,
                DocumentHash = _document.PathName ?? "Unknown",
                AlgorithmVersion = "1.0",
                Filters = new List<FilterGroupForStorage>(),
                ClashZones = new List<ClashZone>()
            };

            targetFilter.ClashZoneStorage.Filters ??= new List<FilterGroupForStorage>();

            var placementLogPath = TryGetPlacementLogPath();
            var groupName = BuildFilterGroupName(baseFilterName, category);
            var normalizedKey = new ProcessedFileCombo
            {
                LinkedFile = comboKey.LinkedFile,
                HostFile = comboKey.HostFile
            }.GetNormalizedKey();

            PruneInvalidGroups(targetFilter.ClashZoneStorage);

            var filterGroup = targetFilter.ClashZoneStorage.Filters
                .FirstOrDefault(f => string.Equals(f?.Name, groupName, StringComparison.OrdinalIgnoreCase));

            if (filterGroup == null)
            {
                filterGroup = new FilterGroupForStorage
                {
                    Name = groupName,
                    FileCombos = new List<FilterFileComboGroup>()
                };
                targetFilter.ClashZoneStorage.Filters.Add(filterGroup);
            }

            MergeLegacyGroupsIntoTarget(targetFilter.ClashZoneStorage, filterGroup, normalizedKey, placementLogPath);

            filterGroup.FileCombos ??= new List<FilterFileComboGroup>();

            var filterFileCombo = filterGroup.FileCombos
                .FirstOrDefault(fc => fc.GetNormalizedKey() == normalizedKey);

            if (filterFileCombo == null)
            {
                filterFileCombo = new FilterFileComboGroup
                {
                    LinkedFile = comboKey.LinkedFile,
                    HostFile = comboKey.HostFile,
                    ProcessedAt = DateTime.Now,
                    ClashZones = new List<ClashZone>()
                };
                filterGroup.FileCombos.Add(filterFileCombo);
            }

            filterFileCombo.ClashZones ??= new List<ClashZone>();

            var existingById = filterFileCombo.ClashZones
                .Where(z => z != null)
                .GroupBy(z => z.Id)
                .Select(g => g.First())
                .ToDictionary(z => z.Id, z => z);

            filterFileCombo.ClashZones = existingById.Values.ToList();

            foreach (var newZone in comboClashZones)
            {
                if (!existingById.TryGetValue(newZone.Id, out var existingZone))
                {
                    if (!ShouldSkipAdd(newZone))
                    {
                        filterFileCombo.ClashZones.Add(newZone);
                        existingById[newZone.Id] = newZone;
                        stats.FilterAdded++;
                        LogPlacement($"[PERSIST-ADD] Zone={newZone.Id}, SleeveId={newZone.SleeveInstanceId}, W={newZone.SleeveWidth:F6}, H={newZone.SleeveHeight:F6}, D={newZone.SleeveDiameter:F6}");
                    }
                }
                else
                {
                    MergeZone(existingZone, newZone, placementLogPath, allowStructuralUpdates);
                    stats.FilterUpdated++;
                }
            }

            targetFilter.ClashZoneStorage.LastUpdated = DateTime.Now;
            targetFilter.LastModified = DateTime.Now;
        }

        private void CleanupGlobalIndex(CategoryGlobalIndex globalIndex, List<ClashZone> validZones, ProcessingStats stats)
        {
            if (globalIndex == null)
                return;

            var validIds = new HashSet<string>(
                validZones.Select(z => z.Id.ToString()),
                StringComparer.OrdinalIgnoreCase);

            var allEntries = GlobalIndexService.GetAllEntries(globalIndex).ToList();

            foreach (var entry in allEntries)
            {
                if (entry == null)
                    continue;

                if (IsOrphanEntry(entry) && !validIds.Contains(entry.Id))
                {
                    if (RemoveGlobalEntry(globalIndex, entry))
                        stats.GlobalRemoved++;
                }
            }
        }

        private void LogAggregate(IEnumerable<ProcessingStats> statsCollection, string baseFilterName)
        {
            var statsList = statsCollection?
                .Where(s => s != null)
                .ToList() ?? new List<ProcessingStats>();

            if (statsList.Count == 0)
            {
                LogRefresh("[CLASH-ZONE-PERSISTENCE] No categories persisted");
                return;
            }

            foreach (var stat in statsList)
            {
                LogRefresh($"[CLASH-ZONE-PERSISTENCE]   {stat}");
            }

            var summary = ProcessingStats.Combine(statsList);
            LogRefresh($"[CLASH-ZONE-PERSISTENCE] TOTAL → {summary}");

            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Info($"[CLASH-ZONE-PERSISTENCE] ✅ Successfully saved clash zones for '{baseFilterName}' → {summary}");
        }

        private void LogPlacement(string message)
        {
            try
            {
                var logPath = SafeFileLogger.GetLogFilePath("placement_debug.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss}] {message}\n");
            }
            catch { }
        }

        private void LogRefresh(string message)
        {
            SafeFileLogger.SafeAppendText(_refreshLogName, $"[{DateTime.Now}] {message}\n");
        }

        private void HandleException(string context, Exception ex)
        {
            LogRefresh($"{context}: {ex.Message}");
            if (!DeploymentConfiguration.DeploymentMode)
                DebugLogger.Error($"{context}: {ex}");
        }

        private static string NormalizeBaseFilterName(string rawName)
        {
            if (string.IsNullOrWhiteSpace(rawName))
                return string.Empty;

            var trimmed = rawName.Trim();
            return Path.GetFileNameWithoutExtension(trimmed);
        }

        private static void EnsureFilterStorageInitialized(OpeningFilter filter)
        {
            if (filter == null)
                return;

            filter.ClashZoneStorage ??= new ClashZoneStorage();
            filter.ClashZoneStorage.Filters ??= new List<FilterGroupForStorage>();
        }

        private static void DeduplicateFilterStorage(OpeningFilter filter)
        {
            if (filter?.ClashZoneStorage?.Filters == null)
                return;

            var storage = filter.ClashZoneStorage;
            var dedupedGroups = new List<FilterGroupForStorage>();
            var seenGroups = new Dictionary<string, FilterGroupForStorage>(StringComparer.OrdinalIgnoreCase);

            foreach (var group in storage.Filters)
            {
                if (group == null || string.IsNullOrWhiteSpace(group.Name))
                    continue;

                var name = group.Name.Trim();
                DeduplicateFilterFileCombos(group);

                if (!seenGroups.TryGetValue(name, out var existing))
                {
                    dedupedGroups.Add(group);
                    seenGroups[name] = group;
                }
                else
                {
                    MergeFilterGroup(existing, group);
                }
            }

            storage.Filters = dedupedGroups;
            storage.ClashZones = storage.AllZones?.ToList() ?? new List<ClashZone>();
            storage.LastUpdated = DateTime.Now;
            filter.LastModified = DateTime.Now;
        }

        private static void DeduplicateFilterFileCombos(FilterGroupForStorage group)
        {
            if (group?.FileCombos == null)
                return;

            var deduped = new List<FilterFileComboGroup>();
            var seen = new Dictionary<string, FilterFileComboGroup>(StringComparer.OrdinalIgnoreCase);

            foreach (var combo in group.FileCombos)
            {
                if (combo == null)
                    continue;

                var key = combo.GetNormalizedKey();
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                DeduplicateClashZones(combo);

                if (!seen.TryGetValue(key, out var existing))
                {
                    deduped.Add(combo);
                    seen[key] = combo;
                }
                else
                {
                    MergeFilterFileCombo(existing, combo);
                }
            }

            group.FileCombos = deduped;
        }

        private static void MergeFilterGroup(FilterGroupForStorage target, FilterGroupForStorage source)
        {
            if (target == null || source?.FileCombos == null)
                return;

            target.FileCombos ??= new List<FilterFileComboGroup>();

            var map = target.FileCombos
                .Where(fc => fc != null)
                .ToDictionary(fc => fc.GetNormalizedKey(), fc => fc, StringComparer.OrdinalIgnoreCase);

            foreach (var combo in source.FileCombos)
            {
                if (combo == null)
                    continue;

                var key = combo.GetNormalizedKey();
                if (string.IsNullOrWhiteSpace(key))
                    continue;

                DeduplicateClashZones(combo);

                if (!map.TryGetValue(key, out var existing))
                {
                    target.FileCombos.Add(combo);
                    map[key] = combo;
                }
                else
                {
                    MergeFilterFileCombo(existing, combo);
                }
            }
        }

        private static void MergeFilterFileCombo(FilterFileComboGroup target, FilterFileComboGroup source)
        {
            if (target == null || source?.ClashZones == null)
                return;

            target.ClashZones ??= new List<ClashZone>();

            var indexMap = new Dictionary<Guid, int>();
            for (int i = 0; i < target.ClashZones.Count; i++)
            {
                var existing = target.ClashZones[i];
                if (existing == null || existing.Id == Guid.Empty)
                    continue;
                indexMap[existing.Id] = i;
            }

            foreach (var zone in source.ClashZones)
            {
                if (zone == null || zone.Id == Guid.Empty)
                    continue;

                if (indexMap.TryGetValue(zone.Id, out var index))
                {
                    target.ClashZones[index] = zone;
                }
                else
                {
                    indexMap[zone.Id] = target.ClashZones.Count;
                    target.ClashZones.Add(zone);
                }
            }

            DeduplicateClashZones(target);
            if (source.ProcessedAt > target.ProcessedAt)
                target.ProcessedAt = source.ProcessedAt;
        }

        private static void DeduplicateClashZones(FilterFileComboGroup combo)
        {
            if (combo?.ClashZones == null)
                return;

            var seen = new HashSet<Guid>();
            var deduped = new List<ClashZone>();

            foreach (var zone in combo.ClashZones)
            {
                if (zone == null || zone.Id == Guid.Empty)
                    continue;

                if (seen.Add(zone.Id))
                {
                    deduped.Add(zone);
                }
            }

            combo.ClashZones = deduped;
        }

        private static string BuildFilterFileName(string baseFilterName, string category)
        {
            if (string.IsNullOrWhiteSpace(baseFilterName))
                return category?.ToLowerInvariant() switch
                {
                    null or "" => string.Empty,
                    _ => $"{category.ToLowerInvariant()}.xml"
                };

            if (string.IsNullOrWhiteSpace(category))
                return $"{baseFilterName}.xml";

            return $"{baseFilterName}_{category.ToLowerInvariant()}.xml";
        }

        private static string BuildFilterGroupName(string baseFilterName, string category)
        {
            if (string.IsNullOrWhiteSpace(category))
                return baseFilterName ?? string.Empty;

            var suffix = "_" + category.ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(baseFilterName))
                return suffix.TrimStart('_');

            return baseFilterName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                ? baseFilterName
                : baseFilterName + suffix;
        }

        private (string LinkedFile, string HostFile) GetFileComboKey(ClashZone clashZone)
        {
            if (clashZone == null)
                return (string.Empty, string.Empty);

            // ✅ FIX: Try to get RevitLinkInstance.Name first (matches UI display), then fall back to SourceDocKey/DocumentPath
            // This ensures file names match between UI selections and clash zone persistence
            var linkedFile = GetLinkInstanceName(clashZone.SourceDocKey) 
                ?? FirstNonEmptyNormalized(clashZone.SourceDocKey, clashZone.DocumentPath);

            var hostFile = GetLinkInstanceName(clashZone.HostDocKey)
                ?? FirstNonEmptyNormalized(clashZone.HostDocKey, clashZone.StructuralElementDocumentTitle);

            if (string.IsNullOrWhiteSpace(linkedFile))
            {
                linkedFile = "unknown-linked";
                LogRefresh($"[PERSIST-WARN] Clash zone {clashZone.Id} missing SourceDocKey/DocumentPath. Falling back to '{linkedFile}'.");
            }

            if (string.IsNullOrWhiteSpace(hostFile))
            {
                hostFile = "unknown-host";
                LogRefresh($"[PERSIST-WARN] Clash zone {clashZone.Id} missing HostDocKey/StructuralElementDocumentTitle. Falling back to '{hostFile}'.");
            }

            return (linkedFile, hostFile);
        }

        /// <summary>
        /// ✅ FIX: Get RevitLinkInstance.Name for a document title or path (matches UI display)
        /// Looks up the link instance in the main document and returns its Name if available
        /// Handles both Document.Title and full file paths (extracts filename from path)
        /// </summary>
        private string GetLinkInstanceName(string documentTitleOrPath)
        {
            if (string.IsNullOrWhiteSpace(documentTitleOrPath) || _document == null)
                return null;

            try
            {
                // Extract file name from path if it's a full path (e.g., "C:\...\PH-00001.rvt" -> "PH-00001")
                string searchFileName = documentTitleOrPath;
                if (documentTitleOrPath.Contains("\\") || documentTitleOrPath.Contains("/"))
                {
                    // It's a path - extract filename
                    searchFileName = System.IO.Path.GetFileNameWithoutExtension(documentTitleOrPath);
                }
                
                // Normalize for comparison
                var normalizedSearch = NormalizeFileName(searchFileName);
                
                // Find RevitLinkInstance in main document that links to a document with matching title/path
                var linkInstance = new FilteredElementCollector(_document)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .FirstOrDefault(link =>
                    {
                        var linkDoc = link.GetLinkDocument();
                        if (linkDoc == null) return false;
                        
                        // Match by title (normalized) - compare both Document.Title and filename from path
                        var linkTitle = NormalizeFileName(linkDoc.Title);
                        var linkPathName = NormalizeFileName(System.IO.Path.GetFileNameWithoutExtension(linkDoc.PathName ?? ""));
                        
                        return string.Equals(linkTitle, normalizedSearch, StringComparison.OrdinalIgnoreCase) ||
                               string.Equals(linkPathName, normalizedSearch, StringComparison.OrdinalIgnoreCase);
                    });

                if (linkInstance != null && !string.IsNullOrWhiteSpace(linkInstance.Name))
                {
                    // ✅ Return raw name (matches UI format) - normalization happens in GetNormalizedKey() for matching
                    return linkInstance.Name;
                }
            }
            catch (Exception ex)
            {
                // Silently fail - fall back to document title
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[PERSIST] Error looking up link instance name for '{documentTitleOrPath}': {ex.Message}");
            }

            return null; // Fall back to document title
        }

        private string FirstNonEmptyNormalized(params string[] candidates)
        {
            foreach (var candidate in candidates)
            {
                var normalized = NormalizeFileName(candidate);
                if (!string.IsNullOrWhiteSpace(normalized))
                    return normalized;
            }

            return string.Empty;
        }

        private static string TryGetPlacementLogPath()
        {
            try
            {
                return SafeFileLogger.GetLogFilePath("placement_debug.log");
            }
            catch
            {
                return null;
            }
        }

        private static void PruneInvalidGroups(ClashZoneStorage storage)
        {
            if (storage?.Filters == null)
                return;

            storage.Filters.RemoveAll(f =>
                f == null ||
                string.IsNullOrWhiteSpace(f.Name) ||
                string.Equals(f.Name, "Unknown", StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsValidClashZone(ClashZone zone)
        {
            if (zone == null || zone.Id == Guid.Empty)
                return false;

            var mepId = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
            var hostId = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
            if (mepId <= 0 || hostId <= 0)
                return false;

            var hasIntersection =
                Math.Abs(zone.IntersectionPointX) > 1e-9 ||
                Math.Abs(zone.IntersectionPointY) > 1e-9 ||
                Math.Abs(zone.IntersectionPointZ) > 1e-9;

            if (!hasIntersection)
                return false;

            var keys = new[]
            {
                zone.SourceDocKey,
                zone.DocumentPath,
                zone.HostDocKey,
                zone.StructuralElementDocumentTitle
            };

            return keys.Any(k => !string.IsNullOrWhiteSpace(NormalizeKey(k)));

            static string NormalizeKey(string key) => string.IsNullOrWhiteSpace(key) ? string.Empty : key.Trim();
        }

        private static bool IsZeroIntersection(CategoryGlobalIndexEntry entry)
        {
            return entry != null &&
                   Math.Abs(entry.IntersectionPointX) < 1e-9 &&
                   Math.Abs(entry.IntersectionPointY) < 1e-9 &&
                   Math.Abs(entry.IntersectionPointZ) < 1e-9;
        }

        private static bool IsOrphanEntry(CategoryGlobalIndexEntry entry)
        {
            if (entry == null) return false;

            return entry.MepElementId == 0 ||
                   entry.StructuralElementId == 0 ||
                   IsZeroIntersection(entry);
        }

        private static bool RemoveGlobalEntry(CategoryGlobalIndex globalIndex, CategoryGlobalIndexEntry entry)
        {
            if (globalIndex?.Filters != null)
            {
                foreach (var filter in globalIndex.Filters)
                {
                    if (filter?.FileCombos == null) continue;
                    foreach (var fileCombo in filter.FileCombos)
                    {
                        if (fileCombo?.Entries != null && fileCombo.Entries.Remove(entry))
                            return true;
                    }
                }
            }

            if (globalIndex?.Entries != null && globalIndex.Entries.Remove(entry))
                return true;

            return false;
        }

        private sealed class ProcessingStats
        {
            public ProcessingStats(string category)
            {
                Category = category ?? string.Empty;
            }

            public string Category { get; }
            public int TotalZones { get; set; }
            public int ValidZones { get; set; }
            public int InvalidZones { get; set; }
            public int FileComboCount { get; set; }
            public int GlobalCreated { get; set; }
            public int GlobalUpdated { get; set; }
            public int GlobalRemoved { get; set; }
            public int FilterAdded { get; set; }
            public int FilterUpdated { get; set; }

            public void Merge(ProcessingStats other)
            {
                if (other == null) return;

                TotalZones += other.TotalZones;
                ValidZones += other.ValidZones;
                InvalidZones += other.InvalidZones;
                FileComboCount += other.FileComboCount;
                GlobalCreated += other.GlobalCreated;
                GlobalUpdated += other.GlobalUpdated;
                GlobalRemoved += other.GlobalRemoved;
                FilterAdded += other.FilterAdded;
                FilterUpdated += other.FilterUpdated;
            }

            public override string ToString()
            {
                return $"Category='{Category}', Zones={ValidZones}/{TotalZones}, Invalid={InvalidZones}, FileCombos={FileComboCount}, Global({GlobalCreated}/{GlobalUpdated}/{GlobalRemoved}), Filter({FilterAdded}/{FilterUpdated})";
            }

            public static ProcessingStats Combine(IEnumerable<ProcessingStats> stats)
            {
                var aggregate = new ProcessingStats("ALL");
                foreach (var stat in stats)
                {
                    aggregate.Merge(stat);
                }
                return aggregate;
            }
        }

        /// <summary>
        /// Helper method to normalize file names (matches ProcessedFileCombo.GetNormalizedKey logic exactly)
        /// ✅ CRITICAL: Must match GetNormalizedKey() normalization to ensure file combos match correctly
        /// </summary>
        private string NormalizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return string.Empty;
            
            var trimmed = fileName.Trim();
            
            // ✅ FIX: Remove old format ": X : location Shared" pattern (e.g., ": 12 : location Shared")
            // This handles legacy file combo names from Global XML
            var locationMatch = System.Text.RegularExpressions.Regex.Match(trimmed, @":\s*\d+\s*:\s*location\s+Shared", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (locationMatch.Success)
            {
                trimmed = trimmed.Substring(0, locationMatch.Index).Trim();
            }
            
            var idxParen = trimmed.IndexOf('(');
            if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
            trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
            trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
            trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
            trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
            return trimmed.Trim();
        }

        private static bool ShouldSkipAdd(ClashZone zone)
        {
            if (zone == null)
                return true;

            // 🚫 DO NOT tighten this guard back to sleeve-dependent checks.
            // Refresh runs before placement, so SleeveInstanceId / dimensions / bbox will be zero.
            // If we require those fields here, the tree stays empty and the placer has nothing to work with.
            // We intentionally persist any clash that has a valid identity + intersection data.
            // Prior to placement we expect sleeve data to be zeroed out, so verify only the
            // core clash zone identity and intersection data. Anything with valid IDs and
            // intersection point should be persisted so the placer can work later.
            if (zone.Id == Guid.Empty)
                return true;

            var mepId = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
            var hostId = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
            if (mepId <= 0 || hostId <= 0)
                return true;

            // Require at least one non-zero intersection component so we ignore totally empty records
            var hasIntersection = Math.Abs(zone.IntersectionPointX) > 1e-9 ||
                                   Math.Abs(zone.IntersectionPointY) > 1e-9 ||
                                   Math.Abs(zone.IntersectionPointZ) > 1e-9;
            if (!hasIntersection)
                return true;

            return false;
        }

        private static void MergeZone(ClashZone target, ClashZone source, string logPath, bool allowStructuralUpdates)
        {
            if (target == null || source == null) return;

            if (allowStructuralUpdates)
            {
                if (!string.IsNullOrWhiteSpace(source.SourceDocKey))
                    target.SourceDocKey = source.SourceDocKey;
                if (!string.IsNullOrWhiteSpace(source.HostDocKey))
                    target.HostDocKey = source.HostDocKey;

                if (source.MepParameterValues != null && source.MepParameterValues.Count > 0)
                    target.MepParameterValues = source.MepParameterValues;
                if (source.HostParameterValues != null && source.HostParameterValues.Count > 0)
                    target.HostParameterValues = source.HostParameterValues;

                if (source.StructuralElementThickness > 0)
                    target.StructuralElementThickness = source.StructuralElementThickness;
                if (source.StructuralElementNormal != null)
                    target.StructuralElementNormal = source.StructuralElementNormal;

                target.IntersectionPointX = source.IntersectionPointX;
                target.IntersectionPointY = source.IntersectionPointY;
                target.IntersectionPointZ = source.IntersectionPointZ;
            }

            if (HasPlacementPoint(source))
            {
                target.SleevePlacementPointX = source.SleevePlacementPointX;
                target.SleevePlacementPointY = source.SleevePlacementPointY;
                target.SleevePlacementPointZ = source.SleevePlacementPointZ;
            }

            if (HasActivePlacementPoint(source))
            {
                target.SleevePlacementPointActiveDocumentX = source.SleevePlacementPointActiveDocumentX;
                target.SleevePlacementPointActiveDocumentY = source.SleevePlacementPointActiveDocumentY;
                target.SleevePlacementPointActiveDocumentZ = source.SleevePlacementPointActiveDocumentZ;
            }

            if (source.SleeveWidth > 0)
                target.SleeveWidth = source.SleeveWidth;
            if (source.SleeveHeight > 0)
                target.SleeveHeight = source.SleeveHeight;
            if (source.SleeveDiameter > 0)
                target.SleeveDiameter = source.SleeveDiameter;

            if (HasBoundingBox(source))
            {
                target.SleeveBoundingBoxMinX = source.SleeveBoundingBoxMinX;
                target.SleeveBoundingBoxMinY = source.SleeveBoundingBoxMinY;
                target.SleeveBoundingBoxMinZ = source.SleeveBoundingBoxMinZ;
                target.SleeveBoundingBoxMaxX = source.SleeveBoundingBoxMaxX;
                target.SleeveBoundingBoxMaxY = source.SleeveBoundingBoxMaxY;
                target.SleeveBoundingBoxMaxZ = source.SleeveBoundingBoxMaxZ;
            }

            if (source.SleeveInstanceId > 0)
                target.SleeveInstanceId = source.SleeveInstanceId;
            if (source.ClusterSleeveInstanceId > 0)
                target.ClusterSleeveInstanceId = source.ClusterSleeveInstanceId;
            if (source.AfterClusterSleevePlacedSleeveInstanceId > 0)
                target.AfterClusterSleevePlacedSleeveInstanceId = source.AfterClusterSleevePlacedSleeveInstanceId;

            target.IsResolved = source.IsResolved || target.IsResolved;
            target.IsClusterResolved = source.IsClusterResolved || target.IsClusterResolved;
            target.MarkedForClusteringSleeveProcess = source.MarkedForClusteringSleeveProcess ?? target.MarkedForClusteringSleeveProcess;

            if (!string.IsNullOrWhiteSpace(source.SleeveFamilyName))
                target.SleeveFamilyName = source.SleeveFamilyName;

            try
            {
                if (!string.IsNullOrWhiteSpace(logPath))
                {
                    File.AppendAllText(logPath,
                        $"[{DateTime.Now:HH:mm:ss}] [PERSIST-UPDATE] Zone={source.Id}, SleeveId={source.SleeveInstanceId}, W={source.SleeveWidth:F6}, H={source.SleeveHeight:F6}, D={source.SleeveDiameter:F6}\n");
                }
            }
            catch { }
        }

        private static bool HasPlacementPoint(ClashZone zone)
        {
            if (zone == null) return false;
            return Math.Abs(zone.SleevePlacementPointX) > 1e-9 || Math.Abs(zone.SleevePlacementPointY) > 1e-9 || Math.Abs(zone.SleevePlacementPointZ) > 1e-9;
        }

        private static bool HasActivePlacementPoint(ClashZone zone)
        {
            if (zone == null) return false;
            return Math.Abs(zone.SleevePlacementPointActiveDocumentX) > 1e-9 || Math.Abs(zone.SleevePlacementPointActiveDocumentY) > 1e-9 || Math.Abs(zone.SleevePlacementPointActiveDocumentZ) > 1e-9;
        }

        private static bool HasBoundingBox(ClashZone zone)
        {
            if (zone == null) return false;
            return Math.Abs(zone.SleeveBoundingBoxMinX) > 1e-9 || Math.Abs(zone.SleeveBoundingBoxMinY) > 1e-9 || Math.Abs(zone.SleeveBoundingBoxMinZ) > 1e-9 ||
                   Math.Abs(zone.SleeveBoundingBoxMaxX) > 1e-9 || Math.Abs(zone.SleeveBoundingBoxMaxY) > 1e-9 || Math.Abs(zone.SleeveBoundingBoxMaxZ) > 1e-9;
        }

        private static void UpdateGlobalEntry(CategoryGlobalIndexEntry entry, ClashZone zone, string filterName, bool allowStructuralUpdates)
        {
            if (entry == null || zone == null) return;

            entry.FilterName = filterName ?? entry.FilterName ?? string.Empty;

            entry.IsResolved = zone.IsResolved;
            entry.IsClusterResolved = zone.IsClusterResolved;
            entry.SleeveInstanceId = zone.SleeveInstanceId;
            entry.ClusterSleeveInstanceId = zone.ClusterSleeveInstanceId;

            if (allowStructuralUpdates)
            {
                entry.MepElementId = zone.MepElementId?.IntegerValue ?? zone.MepElementIdValue;
                entry.StructuralElementId = zone.StructuralElementId?.IntegerValue ?? zone.StructuralElementIdValue;
                entry.IntersectionPointX = zone.IntersectionPointX;
                entry.IntersectionPointY = zone.IntersectionPointY;
                entry.IntersectionPointZ = zone.IntersectionPointZ;
            }
        }

        private void ConsolidateFileCombos(FilterGroupForStorage filterGroup, string logPath)
        {
            if (filterGroup?.FileCombos == null || filterGroup.FileCombos.Count <= 1)
                return;

            var combos = filterGroup.FileCombos
                .Where(fc => fc != null)
                .GroupBy(fc => fc.GetNormalizedKey())
                .ToList();

            var duplicatesToRemove = new List<FilterFileComboGroup>();

            foreach (var grouping in combos)
            {
                var ordered = grouping
                    .OrderByDescending(fc => fc.ProcessedAt)
                    .ToList();

                var keeper = ordered.First();
                bool mergedAny = false;

                foreach (var duplicate in ordered.Skip(1))
                {
                    mergedAny = true;

                    if (duplicate.ClashZones != null && duplicate.ClashZones.Count > 0)
                    {
                        if (keeper.ClashZones == null)
                            keeper.ClashZones = new List<ClashZone>();

                        foreach (var zone in duplicate.ClashZones)
                        {
                            if (zone == null) continue;
                            var existing = keeper.ClashZones.FirstOrDefault(z => z != null && z.Id == zone.Id);
                            if (existing == null)
                            {
                                keeper.ClashZones.Add(zone);
                            }
                            else
                            {
                                MergeZone(existing, zone, logPath, allowStructuralUpdates: true);
                            }
                        }
                    }

                    duplicatesToRemove.Add(duplicate);
                }

                if (mergedAny)
                {
                    keeper.ProcessedAt = ordered.Max(fc => fc.ProcessedAt);
                    keeper.ClashZones = keeper.ClashZones?
                        .Where(z => z != null)
                        .GroupBy(z => z.Id)
                        .Select(g => g.First())
                        .ToList();

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(logPath))
                        {
                            File.AppendAllText(logPath,
                                $"[{DateTime.Now:HH:mm:ss}] [PERSIST-CONSOLIDATE] Deduped combo '{grouping.Key}' – kept {keeper.ClashZones?.Count ?? 0} zones, removed {ordered.Count - 1} duplicates\n");
                        }
                    }
                    catch { }
                }
            }

            if (duplicatesToRemove.Count > 0)
            {
                foreach (var duplicate in duplicatesToRemove)
                {
                    filterGroup.FileCombos.Remove(duplicate);
                }
            }
        }

        private void MergeLegacyGroupsIntoTarget(
            ClashZoneStorage storage,
            FilterGroupForStorage targetGroup,
            string normalizedKey,
            string logPath)
        {
            if (storage?.Filters == null || targetGroup == null)
                return;

            var filtersToProcess = storage.Filters.ToList();
            foreach (var group in filtersToProcess)
            {
                if (group == null || ReferenceEquals(group, targetGroup))
                    continue;

                if (group.FileCombos == null || group.FileCombos.Count == 0)
                {
                    storage.Filters.Remove(group);
                    continue;
                }

                var combosToMove = group.FileCombos
                    .Where(fc => fc != null &&
                                 (string.IsNullOrWhiteSpace(normalizedKey) ||
                                  string.Equals(fc.GetNormalizedKey(), normalizedKey, StringComparison.OrdinalIgnoreCase)))
                    .ToList();

                if (combosToMove.Count == 0 &&
                    string.Equals(group.Name, targetGroup.Name, StringComparison.OrdinalIgnoreCase))
                {
                    combosToMove = group.FileCombos.ToList();
                }

                foreach (var combo in combosToMove)
                {
                    group.FileCombos.Remove(combo);
                    targetGroup.FileCombos.Add(combo);

                    try
                    {
                        if (!string.IsNullOrWhiteSpace(logPath))
                        {
                            File.AppendAllText(logPath,
                                $"[{DateTime.Now:HH:mm:ss}] [PERSIST-GROUP-MERGE] Moved combo {combo.LinkedFile}|{combo.HostFile} from '{group.Name}' to '{targetGroup.Name}'\n");
                        }
                    }
                    catch { }
                }

                if (group.FileCombos == null || group.FileCombos.Count == 0)
                {
                    storage.Filters.Remove(group);
                }
            }

            ConsolidateFileCombos(targetGroup, logPath);

            storage.Filters.RemoveAll(f =>
                !ReferenceEquals(f, targetGroup) &&
                (f == null || f.FileCombos == null || f.FileCombos.Count == 0));

            if (!storage.Filters.Contains(targetGroup))
            {
                storage.Filters.Add(targetGroup);
            }
        }
    }
}

