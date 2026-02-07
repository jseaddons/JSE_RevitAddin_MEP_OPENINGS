using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;

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
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] [CombinedSleeve] Discovery: categories=[{string.Join(", ", categories)}], zones loaded={zones?.Count ?? 0}\n");

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
                
                // ✅ STEP 1: Filter zones by category (REMOVED: redundant and broken for Cable Tray vs CableTrays)
                // The repository LoadClusteredZones already filters and performs fuzzy matching.
                var categoryFilteredZones = zones; 
                
                DebugLogger.Info($"[CombinedDiscovery] Category filter: Using all {zones.Count} zones from repository (which already handled category mapping)");
                
                try
                {
                    // ✅ SECTION BOX FILTERING: Get cached section box bounds if available
                    BoundingBoxXYZ sectionBoxBounds = null;
                    
                    // ✅ OPTIMIZATION: Use cached section box if flag is enabled
                    if (OptimizationFlags.UseSectionBoxFilterForParameterTransfer)
                    {
                        // Try to get cached section box bounds from database
                        try
                        {
                            using (var dbContext = new SleeveDbContext(uiDoc.Document, msg => { }))
                            {
                                var sectionBoxService = new SectionBoxService();
                                sectionBoxBounds = sectionBoxService.GetSectionBoxBounds(dbContext.Connection);
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[CombinedDiscovery] Failed to get cached section box bounds: {ex.Message}");
                            }
                        }
                        
                        if (sectionBoxBounds != null && !DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CombinedDiscovery] Using cached section box bounds for filtering");
                        }
                    }
                    
                    // ✅ FALLBACK: If no cached section box, try live Revit API
                    if (sectionBoxBounds == null && uiDoc.ActiveView is Autodesk.Revit.DB.View3D view3D && view3D.IsSectionBoxActive)
                    {
                        sectionBoxBounds = JSE_RevitAddin_MEP_OPENINGS.Services.Helpers.SectionBoxHelper.GetSectionBoxBounds(view3D);
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[CombinedDiscovery] Using live Revit section box bounds (cache not available)");
                        }
                    }
                    
                    if (sectionBoxBounds != null)
                    {
                        // ✅ STEP 3: Filter zones by spatial coordinates (World Space)
                        // This allows discovery of UNPLACED zones (SleeveInstanceId = -1) 
                        // which would be rejected by FilteredElementCollector.
                        int beforeSpatialFilter = categoryFilteredZones.Count;
                        
                        // Add a small tolerance (1 inch) to the section box to be safe
                        double tolerance = 1.0 / 12.0; 
                        
                        zones = categoryFilteredZones.Where(z => {
                            // PRIORITY: Use calculated placement point if available, otherwise raw intersection point
                            double x = z.SleevePlacementPointX != 0 ? z.SleevePlacementPointX : z.IntersectionPointX;
                            double y = z.SleevePlacementPointY != 0 ? z.SleevePlacementPointY : z.IntersectionPointY;
                            double zCoord = z.SleevePlacementPointZ != 0 ? z.SleevePlacementPointZ : z.IntersectionPointZ;
                            
                            // Check if point is inside section box (with tolerance)
                            bool isInside = x >= sectionBoxBounds.Min.X - tolerance && x <= sectionBoxBounds.Max.X + tolerance &&
                                           y >= sectionBoxBounds.Min.Y - tolerance && y <= sectionBoxBounds.Max.Y + tolerance &&
                                           zCoord >= sectionBoxBounds.Min.Z - tolerance && zCoord <= sectionBoxBounds.Max.Z + tolerance;
                                           
                            if (!isInside && !DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CombinedDiscovery]   Zone {z.Id} ({z.MepElementCategory}) EXCLUDED: Point({x:F2}, {y:F2}, {zCoord:F2}) outside SectionBox");
                            }
                            
                            return isInside;
                        }).ToList();
                        
                        DebugLogger.Info($"[CombinedDiscovery] Spatial filter (Section Box): {beforeSpatialFilter} -> {zones.Count} zones are within the bounds");
                        DebugLogger.Info($"[CombinedDiscovery] ✅ FINAL: {zones.Count} zones passed both category and spatial filters");
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

            // ✅ CLUSTER COLLAPSE: For combined proximity we need one sleeve per cluster, not one per zone.
            // Keep all individual zones; for cluster-resolved zones keep one representative per ClusterSleeveInstanceId.
            var individualZones = zones.Where(z => !z.IsClusterResolved || z.ClusterSleeveInstanceId <= 0).ToList();
            var onePerCluster = zones
                .Where(z => z.IsClusterResolved && z.ClusterSleeveInstanceId > 0)
                .GroupBy(z => z.ClusterSleeveInstanceId)
                .Select(g => g.First())
                .ToList();
            zones = individualZones.Concat(onePerCluster).ToList();
            if (onePerCluster.Count > 0 && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[CombinedDiscovery] Collapsed to 1 zone per cluster: {onePerCluster.Count} cluster(s), {individualZones.Count} individual(s) -> {zones.Count} sleeves for proximity");
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
            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] [CombinedSleeve] Discovery result: {added} sleeves added, {fromClashZoneNull} rejected by FromClashZone (categories=[{string.Join(", ", categories)}])\n");

            var materialized = sleeves.ToList();
            if (materialized.Count == 0)
            {
                return Array.Empty<ClusterSleeveInfo>();
            }
            
            // ✅ NOTE: No deduplication needed for parameter aggregation
            // Each zone represents one MEP element (pipe, duct, etc.)
            // 2 pipes = 2 zones = 2 Size values (this is correct!)
            // Cluster-level data is NOT used for parameters - only individual zone data

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
