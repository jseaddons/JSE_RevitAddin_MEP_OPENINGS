using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Utility class for analysing and deduplicating Global XML indexes (Phase B plan).
    /// The helper is deliberately isolated so it can be invoked manually (dry run / apply)
    /// without touching the existing placement pipeline.
    /// </summary>
    public static class GlobalIndexMaintenance
    {
        public sealed class DedupeResult
        {
            public string Category { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public int OriginalFilterCount { get; set; }
            public int OriginalComboCount { get; set; }
            public int OriginalEntryCount { get; set; }
            public int FiltersRemoved { get; set; }
            public int CombosMerged { get; set; }
            public int EntriesMerged { get; set; }
            public bool ChangesApplied { get; set; }
            public bool Success => string.IsNullOrEmpty(ErrorMessage);
            public string ErrorMessage { get; set; } = string.Empty;
        }

        public static DedupeResult Deduplicate(Document document, string categoryName, bool dryRun, StringBuilder diagnosticLog = null)
        {
            var result = new DedupeResult
            {
                Category = categoryName ?? string.Empty,
                FilePath = document == null ? string.Empty : GlobalIndexService.GetCategoryIndexPath(document, categoryName)
            };

            try
            {
                if (document == null)
                {
                    result.ErrorMessage = "Document is null.";
                    return result;
                }

                var indexPath = GlobalIndexService.GetCategoryIndexPath(document, categoryName);
                if (!File.Exists(indexPath))
                {
                    AppendLine(diagnosticLog, $"[DEDUPER] File not found: {indexPath}");
                    return result;
                }

                var categoryIndex = GlobalIndexService.LoadOrCreate(document, categoryName);
                if (categoryIndex == null)
                {
                    result.ErrorMessage = "Failed to load CategoryGlobalIndex.";
                    return result;
                }

                int originalFilterCount = categoryIndex.Filters?.Count ?? 0;
                int originalComboCount = CountCombos(categoryIndex);
                int originalEntryCount = CountEntries(categoryIndex);

                var (filtersRemoved, combosMerged, entriesMerged) = DeduplicateInMemory(categoryIndex, categoryName, diagnosticLog);

                result.OriginalFilterCount = originalFilterCount;
                result.OriginalComboCount = originalComboCount;
                result.OriginalEntryCount = originalEntryCount;
                result.FiltersRemoved = filtersRemoved;
                result.CombosMerged = combosMerged;
                result.EntriesMerged = entriesMerged;

                if (filtersRemoved == 0 && combosMerged == 0 && entriesMerged == 0)
                {
                    AppendLine(diagnosticLog, $"[DEDUPER] No duplicates detected for '{categoryName}'.");
                    return result;
                }

                if (dryRun)
                {
                    AppendLine(diagnosticLog,
                        $"[DEDUPER] Dry run: FiltersRemoved={filtersRemoved}, CombosMerged={combosMerged}, EntriesMerged={entriesMerged}");
                    return result;
                }

                var backupPath = CreateBackup(indexPath, diagnosticLog);

                GlobalIndexService.Save(document, categoryIndex);

                AppendLine(diagnosticLog,
                    $"[DEDUPER] Applied cleanup to '{indexPath}'. Backup: '{backupPath}'");

                result.ChangesApplied = true;
                return result;
            }
            catch (Exception ex)
            {
                AppendLine(diagnosticLog, $"[DEDUPER] ERROR: {ex}");
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        private static (int filtersRemoved, int combosMerged, int entriesMerged) DeduplicateInMemory(
            CategoryGlobalIndex index,
            string categoryName,
            StringBuilder diagnosticLog)
        {
            int filtersRemoved = 0;
            int combosMerged = 0;
            int entriesMerged = 0;

            if (index?.Filters == null || index.Filters.Count == 0)
                return (filtersRemoved, combosMerged, entriesMerged);

            foreach (var filterGroup in index.Filters)
            {
                if (filterGroup?.FileCombos == null)
                    continue;

                combosMerged += DeduplicateFileCombos(filterGroup.FileCombos, diagnosticLog);
            }

            var filters = index.Filters
                .Where(f => f != null && (f.FileCombos?.Count ?? 0) > 0)
                .ToList();

            var groupedByBase = filters
                .GroupBy(f => NormalizeFilterBaseName(f.Name, categoryName), StringComparer.OrdinalIgnoreCase)
                .ToList();

            var dedupedFilters = new List<FilterGroup>();

            foreach (var group in groupedByBase)
            {
                var list = group.ToList();
                if (list.Count == 0)
                    continue;

                var primary = SelectPrimaryFilterGroup(list);

                foreach (var duplicate in list.Where(f => !ReferenceEquals(f, primary)))
                {
                    combosMerged += MergeFilterGroup(primary, duplicate, diagnosticLog);
                    filtersRemoved++;
                }

                NormalizeFilterGroup(primary, group.Key, categoryName);
                dedupedFilters.Add(primary);
            }

            index.Filters = dedupedFilters;

            if (index.ProcessedFileCombos != null && index.ProcessedFileCombos.Count > 0)
            {
                var uniqueCombos = new Dictionary<string, ProcessedFileCombo>(StringComparer.OrdinalIgnoreCase);
                foreach (var combo in index.ProcessedFileCombos)
                {
                    if (combo == null)
                        continue;

                    var key = combo.GetNormalizedKey();
                    if (string.IsNullOrWhiteSpace(key))
                        continue;

                    if (!uniqueCombos.TryGetValue(key, out var existing) || existing.ProcessedAt < combo.ProcessedAt)
                    {
                        uniqueCombos[key] = combo;
                    }
                }

                index.ProcessedFileCombos = uniqueCombos.Values.ToList();
            }

            entriesMerged += NormalizeEntryFilterNames(index, categoryName, diagnosticLog);

            AppendLine(diagnosticLog,
                $"[DEDUPER] Summary → FiltersRemoved={filtersRemoved}, CombosMerged={combosMerged}, EntriesNormalized={entriesMerged}");

            return (filtersRemoved, combosMerged, entriesMerged);
        }

        private static int DeduplicateFileCombos(List<FileComboGroup> combos, StringBuilder diagnosticLog)
        {
            int combosMerged = 0;
            if (combos == null || combos.Count <= 1)
                return combosMerged;

            var map = new Dictionary<string, FileComboGroup>(StringComparer.OrdinalIgnoreCase);
            var deduped = new List<FileComboGroup>();

            foreach (var combo in combos)
            {
                if (combo == null)
                    continue;

                var key = combo.GetNormalizedKey();
                if (string.IsNullOrWhiteSpace(key))
                    key = $"unknown::{Guid.NewGuid()}";

                if (!map.TryGetValue(key, out var existing))
                {
                    map[key] = combo;
                    deduped.Add(combo);
                    DeduplicateEntriesByGuid(combo);
                }
                else
                {
                    combosMerged++;
                    MergeFileCombo(existing, combo);
                    DeduplicateEntriesByGuid(existing);
                }
            }

            combos.Clear();
            combos.AddRange(deduped);

            return combosMerged;
        }

        private static void MergeFileCombo(FileComboGroup target, FileComboGroup source)
        {
            if (target == null || source == null)
                return;

            if (target.Entries == null)
                target.Entries = new List<CategoryGlobalIndexEntry>();

            if (source.Entries == null)
                return;

            var map = target.Entries
                .Where(e => e != null && Guid.TryParse(e.Id, out _))
                .ToDictionary(e => e.Id, e => e, StringComparer.OrdinalIgnoreCase);

            foreach (var entry in source.Entries)
            {
                if (entry == null)
                    continue;

                if (string.IsNullOrWhiteSpace(entry.Id) || !Guid.TryParse(entry.Id, out _))
                {
                    target.Entries.Add(entry);
                    continue;
                }

                if (!map.TryGetValue(entry.Id, out var existing))
                {
                    target.Entries.Add(entry);
                    map[entry.Id] = entry;
                }
                else
                {
                    if (Prefer(entry, existing))
                    {
                        target.Entries.Remove(existing);
                        target.Entries.Add(entry);
                        map[entry.Id] = entry;
                    }
                }
            }
        }

        private static void DeduplicateEntriesByGuid(FileComboGroup combo)
        {
            if (combo?.Entries == null || combo.Entries.Count <= 1)
                return;

            var map = new Dictionary<string, CategoryGlobalIndexEntry>(StringComparer.OrdinalIgnoreCase);
            var deduped = new List<CategoryGlobalIndexEntry>();

            foreach (var entry in combo.Entries)
            {
                if (entry == null)
                    continue;

                if (string.IsNullOrWhiteSpace(entry.Id) || !Guid.TryParse(entry.Id, out _))
                {
                    deduped.Add(entry);
                    continue;
                }

                if (!map.TryGetValue(entry.Id, out var existing))
                {
                    map[entry.Id] = entry;
                    deduped.Add(entry);
                }
                else if (Prefer(entry, existing))
                {
                    map[entry.Id] = entry;
                    var index = deduped.IndexOf(existing);
                    if (index >= 0)
                        deduped[index] = entry;
                }
            }

            combo.Entries = deduped;
        }

        private static int MergeFilterGroup(FilterGroup target, FilterGroup source, StringBuilder diagnosticLog)
        {
            if (target == null || source == null || source.FileCombos == null)
                return 0;

            if (target.FileCombos == null)
                target.FileCombos = new List<FileComboGroup>();

            int combosMerged = 0;

            var map = target.FileCombos
                .Where(fc => fc != null)
                .ToDictionary(fc => fc.GetNormalizedKey(), fc => fc, StringComparer.OrdinalIgnoreCase);

            foreach (var combo in source.FileCombos)
            {
                if (combo == null)
                    continue;

                var key = combo.GetNormalizedKey();
                if (string.IsNullOrWhiteSpace(key))
                    key = $"unknown::{Guid.NewGuid()}";

                if (!map.TryGetValue(key, out var existing))
                {
                    target.FileCombos.Add(combo);
                    map[key] = combo;
                    DeduplicateEntriesByGuid(combo);
                }
                else
                {
                    combosMerged++;
                    MergeFileCombo(existing, combo);
                    DeduplicateEntriesByGuid(existing);
                }
            }

            AppendLine(diagnosticLog,
                $"[DEDUPER] Merged filter group '{source.Name}' into '{target.Name}' → combos merged {combosMerged}");

            return combosMerged;
        }

        private static void NormalizeFilterGroup(FilterGroup filterGroup, string baseFilterName, string categoryName)
        {
            if (filterGroup == null)
                return;

            var normalizedName = baseFilterName;
            if (string.IsNullOrWhiteSpace(normalizedName))
                normalizedName = filterGroup.Name;

            filterGroup.Name = normalizedName;

            if (filterGroup.FileCombos == null)
                return;

            var normalizedFilterFileName = $"{normalizedName}_{Sanitize(categoryName)}.xml";
            foreach (var combo in filterGroup.FileCombos)
            {
                if (combo?.Entries == null)
                    continue;

                foreach (var entry in combo.Entries)
                {
                    if (entry != null)
                        entry.FilterName = normalizedFilterFileName;
                }
            }
        }

        private static int NormalizeEntryFilterNames(CategoryGlobalIndex index, string categoryName, StringBuilder diagnosticLog)
        {
            if (index == null)
                return 0;

            int updated = 0;
            var sanitizedCategory = Sanitize(categoryName);

            foreach (var filter in index.Filters ?? Enumerable.Empty<FilterGroup>())
            {
                if (filter?.FileCombos == null)
                    continue;

                var normalizedFilterName = filter.Name;
                var normalizedFileName = $"{normalizedFilterName}_{sanitizedCategory}.xml";

                foreach (var combo in filter.FileCombos)
                {
                    if (combo?.Entries == null)
                        continue;

                    foreach (var entry in combo.Entries)
                    {
                        if (entry == null)
                            continue;

                        if (!string.Equals(entry.FilterName, normalizedFileName, StringComparison.OrdinalIgnoreCase))
                        {
                            entry.FilterName = normalizedFileName;
                            updated++;
                        }
                    }
                }
            }

            AppendLine(diagnosticLog, $"[DEDUPER] Normalized {updated} entry FilterName attributes.");
            return updated;
        }

        private static FilterGroup SelectPrimaryFilterGroup(List<FilterGroup> groups)
        {
            if (groups == null || groups.Count == 0)
                return null;

            groups.Sort((a, b) =>
            {
                int lengthCompare = (a?.Name?.Length ?? int.MaxValue).CompareTo(b?.Name?.Length ?? int.MaxValue);
                if (lengthCompare != 0)
                    return lengthCompare;
                return string.Compare(a?.Name, b?.Name, StringComparison.OrdinalIgnoreCase);
            });

            return groups[0];
        }

        private static bool Prefer(CategoryGlobalIndexEntry candidate, CategoryGlobalIndexEntry existing)
        {
            if (candidate == null)
                return false;
            if (existing == null)
                return true;

            if (candidate.SleeveInstanceId > 0 && existing.SleeveInstanceId <= 0)
                return true;
            if (existing.SleeveInstanceId > 0 && candidate.SleeveInstanceId <= 0)
                return false;

            if (!candidate.IsResolved && existing.IsResolved)
                return false;
            if (candidate.IsResolved && !existing.IsResolved)
                return true;

            if (candidate.IsClusterResolved && !existing.IsClusterResolved)
                return true;
            if (!candidate.IsClusterResolved && existing.IsClusterResolved)
                return false;

            var candidateMagnitude = Math.Abs(candidate.SleeveBoundingBoxMaxX - candidate.SleeveBoundingBoxMinX) +
                                     Math.Abs(candidate.SleeveBoundingBoxMaxY - candidate.SleeveBoundingBoxMinY) +
                                     Math.Abs(candidate.SleeveBoundingBoxMaxZ - candidate.SleeveBoundingBoxMinZ);
            var existingMagnitude = Math.Abs(existing.SleeveBoundingBoxMaxX - existing.SleeveBoundingBoxMinX) +
                                    Math.Abs(existing.SleeveBoundingBoxMaxY - existing.SleeveBoundingBoxMinY) +
                                    Math.Abs(existing.SleeveBoundingBoxMaxZ - existing.SleeveBoundingBoxMinZ);

            return candidateMagnitude > existingMagnitude;
        }

        private static string NormalizeFilterBaseName(string filterName, string categoryName)
        {
            if (string.IsNullOrWhiteSpace(filterName))
                return string.Empty;

            var name = filterName.Trim();
            if (string.IsNullOrWhiteSpace(categoryName))
                return name;

            var suffix = "_" + Sanitize(categoryName);
            if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
            {
                return name.Substring(0, name.Length - suffix.Length);
            }

            return name;
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            return value.Trim().Replace(" ", "_").ToLowerInvariant();
        }

        private static int CountCombos(CategoryGlobalIndex index)
        {
            if (index?.Filters == null)
                return 0;
            return index.Filters.Sum(f => f?.FileCombos?.Count ?? 0);
        }

        private static int CountEntries(CategoryGlobalIndex index)
        {
            if (index == null)
                return 0;

            int count = 0;

            foreach (var filter in index.Filters ?? Enumerable.Empty<FilterGroup>())
            {
                foreach (var combo in filter?.FileCombos ?? Enumerable.Empty<FileComboGroup>())
                {
                    count += combo?.Entries?.Count ?? 0;
                }
            }

            count += index.Entries?.Count ?? 0;
            return count;
        }

        private static string CreateBackup(string filePath, StringBuilder diagnosticLog)
        {
            try
            {
                var directory = Path.GetDirectoryName(filePath);
                var filename = Path.GetFileNameWithoutExtension(filePath);
                var extension = Path.GetExtension(filePath);
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var backupName = $"{filename}_backup_{timestamp}{extension}";
                var backupPath = Path.Combine(directory ?? string.Empty, backupName);

                File.Copy(filePath, backupPath, overwrite: false);

                AppendLine(diagnosticLog, $"[DEDUPER] Created backup: {backupPath}");
                return backupPath;
            }
            catch (Exception ex)
            {
                AppendLine(diagnosticLog, $"[DEDUPER] WARNING: Failed to create backup for {filePath}: {ex.Message}");
                return string.Empty;
            }
        }

        private static void AppendLine(StringBuilder sb, string message)
        {
            sb?.AppendLine(message);
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info(message);
            }
        }
    }
}

