using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Services
{
    /// <summary>
    /// Phase 1-2 discovery service: gathers cluster sleeves and enriches them with parameter snapshots.
    /// </summary>
    public class CombinedClusterDiscoveryService : ICombinedClusterDiscoveryService
    {
        private readonly ICombinedClusterRepository _repository;

        public CombinedClusterDiscoveryService(ICombinedClusterRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
        }

        public IReadOnlyList<ClusterSleeveInfo> Discover(Autodesk.Revit.UI.UIDocument uiDoc, string filterName, IReadOnlyCollection<string> categories, Autodesk.Revit.DB.BoundingBoxXYZ? sectionBox = null)
        {
            if (!OptimizationFlags.UseCombinedClustering || !OptimizationFlags.UseCombinedClusteringPhase1And2)
            {
                DebugLogger.Warning("[CombinedDiscovery] Flags disabled, returning empty");
                return Array.Empty<ClusterSleeveInfo>();
            }

            var zones = _repository.LoadClusteredZones(uiDoc.Document, filterName, categories);
            DebugLogger.Info($"[CombinedDiscovery] Loaded {zones?.Count ?? 0} zones from repository");

            if (zones == null || zones.Count == 0)
            {
                return Array.Empty<ClusterSleeveInfo>();
            }

            // SECTION BOX FILTERING: Simplified approach
            // 1. Filter zones by UI-selected categories first (zones already have MepElementCategory)
            // 2. Collect all sleeve family instances in the section box from Revit
            // 3. Match the filtered zones against sleeves in section box
            if (sectionBox != null)
            {
                // ✅ DIAGNOSTIC: Log section box bounds
                DebugLogger.Info($"[CombinedDiscovery] Section Box Bounds: Min=({sectionBox.Min.X:F3}, {sectionBox.Min.Y:F3}, {sectionBox.Min.Z:F3}), Max=({sectionBox.Max.X:F3}, {sectionBox.Max.Y:F3}, {sectionBox.Max.Z:F3})");
                DebugLogger.Info($"[CombinedDiscovery] UI-selected categories: [{string.Join(", ", categories)}]");
                
                // ✅ STEP 1: Filter zones by category (zones already have MepElementCategory set)
                var categorySet = new System.Collections.Generic.HashSet<string>(categories, System.StringComparer.OrdinalIgnoreCase);
                var categoryFilteredZones = zones.Where(z => 
                    !string.IsNullOrWhiteSpace(z.MepElementCategory) && 
                    categorySet.Contains(z.MepElementCategory))
                    .ToList();
                
                DebugLogger.Info($"[CombinedDiscovery] Category filter: {zones.Count} -> {categoryFilteredZones.Count} zones match UI-selected categories");
                
                try
                {
                    // ✅ STEP 2: Get all sleeves in section box from Revit
                    var sectionBoxBounds = JSE_RevitAddin_MEP_OPENINGS.Helpers.SectionBoxHelper.GetSectionBoxBounds(uiDoc.ActiveView as Autodesk.Revit.DB.View3D);
                    if (sectionBoxBounds != null)
                    {
                        var outline = new Autodesk.Revit.DB.Outline(sectionBoxBounds.Min, sectionBoxBounds.Max);
                        var bbFilter = new Autodesk.Revit.DB.BoundingBoxIntersectsFilter(outline);
                        
                        // Collect all sleeve family instances in section box
                        var sleevesInSectionBox = new Autodesk.Revit.DB.FilteredElementCollector(uiDoc.Document)
                            .OfClass(typeof(Autodesk.Revit.DB.FamilyInstance))
                            .WherePasses(bbFilter)
                            .Cast<Autodesk.Revit.DB.FamilyInstance>()
                            .Where(fi => fi.Symbol?.FamilyName != null && 
                                   (fi.Symbol.FamilyName.Contains("Sleeve") || 
                                    fi.Symbol.FamilyName.Contains("Opening")))
                            .ToList();
                        
                        var sleeveIdsInSectionBox = new System.Collections.Generic.HashSet<int>(
                            sleevesInSectionBox.Select(s => s.Id.IntegerValue));
                        
                        DebugLogger.Info($"[CombinedDiscovery] Found {sleeveIdsInSectionBox.Count} sleeve elements in section box: [{string.Join(", ", sleeveIdsInSectionBox.Take(10))}]");
                        
                        // ✅ STEP 3: Filter category-matched zones to only those with sleeves in section box
                        int beforeSectionBox = categoryFilteredZones.Count;
                        
                        // ✅ DIAGNOSTIC: Log each zone before filtering
                        DebugLogger.Info($"[CombinedDiscovery] Section box contains {sleeveIdsInSectionBox.Count} sleeve IDs: [{string.Join(", ", sleeveIdsInSectionBox.Take(20))}]");
                        foreach (var z in categoryFilteredZones)
                        {
                            bool individualMatch = z.SleeveInstanceId > 0 && sleeveIdsInSectionBox.Contains(z.SleeveInstanceId);
                            bool clusterMatch = z.ClusterSleeveInstanceId > 0 && sleeveIdsInSectionBox.Contains(z.ClusterSleeveInstanceId);
                            bool afterClusterMatch = z.AfterClusterSleevePlacedSleeveInstanceId > 0 && sleeveIdsInSectionBox.Contains(z.AfterClusterSleevePlacedSleeveInstanceId);
                            bool combinedMatch = z.CombinedClusterSleeveInstanceId > 0 && sleeveIdsInSectionBox.Contains(z.CombinedClusterSleeveInstanceId);
                            bool willPass = individualMatch || clusterMatch || afterClusterMatch || combinedMatch;
                            DebugLogger.Info($"[CombinedDiscovery]   Zone {z.Id} ({z.MepElementCategory}): SleeveId={z.SleeveInstanceId}, ClusterId={z.ClusterSleeveInstanceId}, AfterClusterId={z.AfterClusterSleevePlacedSleeveInstanceId}, CombinedId={z.CombinedClusterSleeveInstanceId}, IndividualMatch={individualMatch}, ClusterMatch={clusterMatch}, AfterClusterMatch={afterClusterMatch}, CombinedMatch={combinedMatch}, WillPass={willPass}");
                        }
                        
                        // ✅ FIX: Check all possible sleeve ID fields (Individual, Cluster, AfterCluster, Combined)
                        zones = categoryFilteredZones.Where(z => 
                            (z.SleeveInstanceId > 0 && sleeveIdsInSectionBox.Contains(z.SleeveInstanceId)) || 
                            (z.ClusterSleeveInstanceId > 0 && sleeveIdsInSectionBox.Contains(z.ClusterSleeveInstanceId)) ||
                            (z.AfterClusterSleevePlacedSleeveInstanceId > 0 && sleeveIdsInSectionBox.Contains(z.AfterClusterSleevePlacedSleeveInstanceId)) ||
                            (z.CombinedClusterSleeveInstanceId > 0 && sleeveIdsInSectionBox.Contains(z.CombinedClusterSleeveInstanceId)))
                            .ToList();
                        
                        DebugLogger.Info($"[CombinedDiscovery] Section box filter: {beforeSectionBox} -> {zones.Count} zones have sleeves in section box");
                        DebugLogger.Info($"[CombinedDiscovery] ✅ FINAL: {zones.Count} zones passed both category and section box filters");
                    }
                    else
                    {
                        DebugLogger.Warning($"[CombinedDiscovery] Could not get section box bounds, using category filter only");
                        zones = categoryFilteredZones;
                    }
                }
                catch (System.Exception ex)
                {
                    DebugLogger.Warning($"[CombinedDiscovery] Section box filtering failed: {ex.Message}, using category filter only");
                    zones = categoryFilteredZones;
                }
            }

            int added = 0;
            int fromClashZoneNull = 0;

            var sleeves = new ConcurrentBag<ClusterSleeveInfo>();
            System.Threading.Tasks.Parallel.ForEach(zones, BuildParallelOptions(), zone =>
            {
                var sleeve = ClusterSleeveInfo.FromClashZone(zone);
                if (sleeve != null)
                {
                    sleeves.Add(sleeve);
                    System.Threading.Interlocked.Increment(ref added);
                }
                else
                {
                    System.Threading.Interlocked.Increment(ref fromClashZoneNull);
                    DebugLogger.Info($"[CombinedDiscovery] Zone {zone.Id} REJECTED by FromClashZone: SleeveId={zone.SleeveInstanceId}, ClusterId={zone.ClusterSleeveInstanceId}, IsClusterResolved={zone.IsClusterResolved}");
                }
            });

            DebugLogger.Info($"[CombinedDiscovery] Results: {added} added, {fromClashZoneNull} rejected by FromClashZone");

            var materialized = sleeves.ToList();
            if (materialized.Count == 0)
            {
                return Array.Empty<ClusterSleeveInfo>();
            }

            var snapshotLookup = _repository.LoadSnapshotParameters(materialized.Select(s => s.SourceSleeveInstanceId));
            foreach (var sleeve in materialized)
            {
                if (sleeve.SourceSleeveInstanceId > 0 && snapshotLookup.TryGetValue(sleeve.SourceSleeveInstanceId, out var parameters))
                {
                    sleeve.ReplaceParameterSnapshot(parameters);
                }
            }

            return materialized
                .OrderBy(s => s.ClusterSleeveInstanceId)
                    .ThenBy(s => s.SourceSleeveInstanceId)
                    .ToList();
        }
             

        private static ParallelOptions BuildParallelOptions()
        {
            var max = OptimizationFlags.CombinedClusterMaxDegreeOfParallelism;
            if (max == 0)
            {
                max = Math.Max(1, Environment.ProcessorCount - 1);
            }

            return new ParallelOptions { MaxDegreeOfParallelism = max };
        }
    }
}
