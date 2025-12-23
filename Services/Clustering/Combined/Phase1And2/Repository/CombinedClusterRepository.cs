using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository
{
    /// <summary>
    /// Repository wrapper that exposes cluster sleeves and parameter snapshots for combined clustering.
    /// </summary>
    public class CombinedClusterRepository : ICombinedClusterRepository
    {
        private readonly IClashZoneRepository _clashZoneRepository;

        public CombinedClusterRepository(IClashZoneRepository clashZoneRepository)
        {
            _clashZoneRepository = clashZoneRepository ?? throw new ArgumentNullException(nameof(clashZoneRepository));
        }

        public IReadOnlyList<ClashZone> LoadClusteredZones(Autodesk.Revit.DB.Document doc, string filterName, IReadOnlyCollection<string> categories)
        {
            if (categories == null || categories.Count == 0)
            {
                DebugLogger.Info($"[LoadClusteredZones] No categories provided, returning empty");
                return Array.Empty<ClashZone>();
            }

            DebugLogger.Info($"[LoadClusteredZones] Starting discovery for filter='{filterName}', categories={string.Join(", ", categories)}");

            // ✅ CRITICAL: Verify sleeves (individual, cluster, combined) still exist in Revit and reset flags for deleted ones
            // This prevents placing sleeves inside remaining combined sleeves after one is deleted
            // Uses existing verification logic that handles all sleeve types
            if (doc != null)
            {
                _clashZoneRepository.VerifyExistingSleevesAndResetFlags(doc, new List<string> { filterName }, categories.ToList());
            }

            var resolved = new List<ClashZone>();
            // ✅ AUTO-DISCOVERY FIX: Handle wildcard or missing filter name by scanning all filters
            bool scanAllFilters = string.IsNullOrWhiteSpace(filterName) || filterName.Trim() == "*";

            foreach (var category in categories)
            {
                List<ClashZone> zones;
                if (scanAllFilters)
                {
                    zones = _clashZoneRepository.GetClashZonesByCategory(category);
                    DebugLogger.Info($"[LoadClusteredZones] Category '{category}': loaded {zones?.Count ?? 0} zones from DB (all filters)");
                }
                else
                {
                    zones = _clashZoneRepository.GetClashZonesByFilter(filterName, category, unresolvedOnly: false, readyForPlacementOnly: false);
                    DebugLogger.Info($"[LoadClusteredZones] Category '{category}': loaded {zones?.Count ?? 0} zones from DB (filter={filterName})");
                }

                // Fuzzy Fallback for Cable Trays (Revit 'CableTrays' vs DB 'Cable Tray')
                if ((zones == null || zones.Count == 0))
                {
                    string altCategory = null;
                    if (string.Equals(category, "CableTrays", StringComparison.OrdinalIgnoreCase)) altCategory = "Cable Tray";
                    else if (string.Equals(category, "Cable Tray", StringComparison.OrdinalIgnoreCase)) altCategory = "CableTrays";
                    
                    if (altCategory != null)
                    {
                        List<ClashZone> altZones;
                        if (scanAllFilters)
                            altZones = _clashZoneRepository.GetClashZonesByCategory(altCategory);
                        else
                            altZones = _clashZoneRepository.GetClashZonesByFilter(filterName, altCategory, unresolvedOnly: false, readyForPlacementOnly: false);

                        if (altZones != null && altZones.Count > 0)
                        {
                            DebugLogger.Info($"[LoadClusteredZones] ⚠️ Category '{category}' yielded 0 results, but alternate '{altCategory}' found {altZones.Count}. Using alternate.");
                            zones = altZones;
                        }
                    }
                    
                    // DIAGNOSTIC: If still 0, log available categories to help user debug
                    if (zones == null || zones.Count == 0)
                    {
                        var available = _clashZoneRepository.GetDistinctCategories();
                        DebugLogger.Warning($"[LoadClusteredZones] ❌ No zones found for '{category}' (or alternate). Available categories in DB: [{string.Join(", ", available)}]");
                    }
                }

                if (zones == null || zones.Count == 0)
                {
                    continue;
                }

                // ✅ CRITICAL FIX: Exclude zones already in combined sleeves to prevent duplicate placement
                int beforeFilter = zones.Count;
                zones = zones.Where(z => !z.IsCombinedResolved).ToList();
                if (beforeFilter > zones.Count)
                {
                    DebugLogger.Info($"[LoadClusteredZones] Category '{category}': filtered out {beforeFilter - zones.Count} zones already in combined sleeves (IsCombinedResolved=1)");
                }

                // ✅ FIX: Include BOTH individual sleeves AND cluster sleeves for Combined Sleeve discovery
                // Individual sleeves: SleeveInstanceId > 0 (not yet clustered, but placed)
                // Cluster sleeves: IsClusterResolved && ClusterSleeveInstanceId > 0 (already clustered)
                // We need both types to find cross-category proximity groups
                // DEBUG: Log each zone before adding
                foreach (var z in zones)
                {
                    DebugLogger.Info($"[LoadClusteredZones]   -> Zone {z.Id}: SleeveInstanceId={z.SleeveInstanceId}, ClusterSleeveInstanceId={z.ClusterSleeveInstanceId}, IsClusterResolved={z.IsClusterResolved}, IsCombinedResolved={z.IsCombinedResolved}");
                }
                // Add all zones (no pre-filter). Revit-first section-box filtering will be applied later in discovery.
                resolved.AddRange(zones);

            }

            // ✅ CRITICAL FIX: Query ClusterSleeves table to find cluster sleeves for selected categories
            // This handles the case where ClashZone.ClusterSleeveInstanceId is -1 but cluster sleeves exist
            try
            {
                var allClusterSleeves = _clashZoneRepository.GetAllClusterSleeves();
                if (allClusterSleeves != null && allClusterSleeves.Count > 0)
                {
                    DebugLogger.Info($"[LoadClusteredZones] Found {allClusterSleeves.Count} cluster sleeves in ClusterSleeves table");
                    
                    // Filter cluster sleeves by category
                    var categorySet = new System.Collections.Generic.HashSet<string>(categories, System.StringComparer.OrdinalIgnoreCase);
                    var relevantClusterSleeves = allClusterSleeves
                        .Where(cs => !string.IsNullOrWhiteSpace(cs.Category) && categorySet.Contains(cs.Category))
                        .ToList();
                    
                    DebugLogger.Info($"[LoadClusteredZones] {relevantClusterSleeves.Count} cluster sleeves match selected categories");
                    
                    // For each cluster sleeve, update or create zones
                    foreach (var clusterSleeve in relevantClusterSleeves)
                    {
                        // Try to find existing zone for this cluster sleeve
                        var existingZone = resolved.FirstOrDefault(z => z.ClusterSleeveInstanceId == clusterSleeve.ClusterInstanceId);
                        
                        if (existingZone != null)
                        {
                            DebugLogger.Info($"[LoadClusteredZones] ✅ Found existing zone for ClusterInstanceId={clusterSleeve.ClusterInstanceId}");
                        }
                        else
                        {
                            // No zone found - create a synthetic zone for this cluster sleeve
                            var syntheticZone = new ClashZone
                            {
                                Id = System.Guid.NewGuid(),
                                ClusterSleeveInstanceId = clusterSleeve.ClusterInstanceId,
                                IsClusterResolved = true,
                                SleeveInstanceId = -1,
                                MepElementCategory = clusterSleeve.Category,
                                IsResolved = true
                            };
                            resolved.Add(syntheticZone);
                            DebugLogger.Info($"[LoadClusteredZones] ⚠️ Created synthetic zone for ClusterInstanceId={clusterSleeve.ClusterInstanceId}, Category={clusterSleeve.Category}");
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                DebugLogger.Warning($"[LoadClusteredZones] Failed to query ClusterSleeves table: {ex.Message}");
            }

            // ✅ CLUSTER CORNER DATA FIX: For cluster sleeves, fetch corner data from ClusterSleeves table
            var clusterInstanceIds = resolved
                .Where(z => z.IsClusterResolved && z.ClusterSleeveInstanceId > 0)
                .Select(z => z.ClusterSleeveInstanceId)
                .Distinct()
                .ToList();

            if (clusterInstanceIds.Count > 0)
            {
                DebugLogger.Info($"[LoadClusteredZones] Fetching corner data for {clusterInstanceIds.Count} cluster sleeves from ClusterSleeves table");
                var clusterSleeves = _clashZoneRepository.GetClusterSleevesByInstanceIds(clusterInstanceIds);
                var clusterCornerLookup = clusterSleeves.ToDictionary(cs => cs.ClusterInstanceId);

                // Merge cluster corner data into ClashZone objects
                foreach (var zone in resolved.Where(z => z.IsClusterResolved && z.ClusterSleeveInstanceId > 0))
                {
                    if (clusterCornerLookup.TryGetValue(zone.ClusterSleeveInstanceId, out var clusterSleeve))
                    {
                        // Override the ClashZone corner data with ClusterSleeve corner data
                        zone.SleeveCorner1X = clusterSleeve.Corner1X;
                        zone.SleeveCorner1Y = clusterSleeve.Corner1Y;
                        zone.SleeveCorner1Z = clusterSleeve.Corner1Z;
                        zone.SleeveCorner2X = clusterSleeve.Corner2X;
                        zone.SleeveCorner2Y = clusterSleeve.Corner2Y;
                        zone.SleeveCorner2Z = clusterSleeve.Corner2Z;
                        zone.SleeveCorner3X = clusterSleeve.Corner3X;
                        zone.SleeveCorner3Y = clusterSleeve.Corner3Y;
                        zone.SleeveCorner3Z = clusterSleeve.Corner3Z;
                        zone.SleeveCorner4X = clusterSleeve.Corner4X;
                        zone.SleeveCorner4Y = clusterSleeve.Corner4Y;
                        zone.SleeveCorner4Z = clusterSleeve.Corner4Z;

                        // ✅ FIX: Calculate BBox from Corners (ClusterSleeve model lacks BBox fields)
                        double c1x = clusterSleeve.Corner1X.GetValueOrDefault();
                        double c2x = clusterSleeve.Corner2X.GetValueOrDefault();
                        double c3x = clusterSleeve.Corner3X.GetValueOrDefault();
                        double c4x = clusterSleeve.Corner4X.GetValueOrDefault();

                        double c1y = clusterSleeve.Corner1Y.GetValueOrDefault();
                        double c2y = clusterSleeve.Corner2Y.GetValueOrDefault();
                        double c3y = clusterSleeve.Corner3Y.GetValueOrDefault();
                        double c4y = clusterSleeve.Corner4Y.GetValueOrDefault();

                        double c1z = clusterSleeve.Corner1Z.GetValueOrDefault();
                        double c2z = clusterSleeve.Corner2Z.GetValueOrDefault();
                        double c3z = clusterSleeve.Corner3Z.GetValueOrDefault();
                        double c4z = clusterSleeve.Corner4Z.GetValueOrDefault();

                        zone.ClusterSleeveBoundingBoxMinX = System.Math.Min(System.Math.Min(c1x, c2x), System.Math.Min(c3x, c4x));
                        zone.ClusterSleeveBoundingBoxMaxX = System.Math.Max(System.Math.Max(c1x, c2x), System.Math.Max(c3x, c4x));
                        
                        zone.ClusterSleeveBoundingBoxMinY = System.Math.Min(System.Math.Min(c1y, c2y), System.Math.Min(c3y, c4y));
                        zone.ClusterSleeveBoundingBoxMaxY = System.Math.Max(System.Math.Max(c1y, c2y), System.Math.Max(c3y, c4y));

                        zone.ClusterSleeveBoundingBoxMinZ = System.Math.Min(System.Math.Min(c1z, c2z), System.Math.Min(c3z, c4z));
                        zone.ClusterSleeveBoundingBoxMaxZ = System.Math.Max(System.Math.Max(c1z, c2z), System.Math.Max(c3z, c4z));

                        DebugLogger.Info($"[LoadClusteredZones]   ✅ Merged cluster corner & BBox data for ClusterInstanceId={zone.ClusterSleeveInstanceId}");
                    }
                    else
                    {
                        DebugLogger.Info($"[LoadClusteredZones]   ⚠️ No cluster corner data found for ClusterInstanceId={zone.ClusterSleeveInstanceId}");
                    }
                }
            }

            DebugLogger.Info($"[LoadClusteredZones] Total returned: {resolved.Count} zones");
            return resolved;
        }

        public Dictionary<int, Dictionary<string, string>> LoadSnapshotParameters(IEnumerable<int> sleeveInstanceIds)
        {
            var ids = sleeveInstanceIds?.Where(id => id > 0).Distinct().ToList();
            if (ids == null || ids.Count == 0)
            {
                return new Dictionary<int, Dictionary<string, string>>();
            }

            return _clashZoneRepository.GetSnapshotMepParametersForSleeveIds(ids);
        }

        public void SaveCombinedClusterMetadata(string filterName, IReadOnlyCollection<ClashZone> updatedZones)
        {
            if (updatedZones == null || updatedZones.Count == 0)
            {
                return;
            }

            var filter = filterName ?? string.Empty;
            foreach (var group in updatedZones.GroupBy(z => z.MepElementCategory ?? string.Empty))
            {
                _clashZoneRepository.InsertOrUpdateClashZones(group.ToList(), filter, group.Key);
            }
        }

        public List<ClashZone> FindIndividualSleevesNearBoundingBox(Autodesk.Revit.DB.BoundingBoxXYZ boundingBox, string hostType, double level, double toleranceFt)
        {
            // Implementation: Retrieve all candidate zones and filter in memory
            // Candidates: Placed (Resolved) but NOT Clustered (IsClusterResolved = false)
            // We search across all relevant categories.
            
            var categories = new[] { "Ducts", "Pipes", "Cable Trays", "Conduits", "Duct Accessories" };
            var allCandidates = new List<ClashZone>();

            foreach (var cat in categories)
            {
                var zones = _clashZoneRepository.GetClashZonesByCategory(cat);
                if (zones != null)
                {
                    // Filter for placed individual sleeves
                    allCandidates.AddRange(zones.Where(z => z.IsResolved && !z.IsClusterResolved));
                }
            }

            var result = new List<ClashZone>();
            foreach (var zone in allCandidates)
            {
                // Basic overlapping check
                // Zone placement point vs BoundingBox + tolerance
                // ClashZone stores "Active" coordinates if placed
                
                // If zone.SleevePlacementPointActiveX is not set/0, fallback to raw point? 
                // Let's assume SleevePlacementPointX/Y/Z are valid.
                
                // Simple point-in-box (expanded) check
                if (zone.SleevePlacementPointX >= boundingBox.Min.X - toleranceFt &&
                    zone.SleevePlacementPointX <= boundingBox.Max.X + toleranceFt &&
                    zone.SleevePlacementPointY >= boundingBox.Min.Y - toleranceFt &&
                    zone.SleevePlacementPointY <= boundingBox.Max.Y + toleranceFt &&
                    zone.SleevePlacementPointZ >= boundingBox.Min.Z - toleranceFt &&
                    zone.SleevePlacementPointZ <= boundingBox.Max.Z + toleranceFt)
                {
                    result.Add(zone);
                }
            }
            return result;
        }
    }
}
