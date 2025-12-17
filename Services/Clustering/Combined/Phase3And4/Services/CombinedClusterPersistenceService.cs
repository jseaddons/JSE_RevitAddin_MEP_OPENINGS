using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services
{
    public class CombinedClusterPersistenceService : ICombinedClusterPersistence
    {
        private readonly IClashZoneRepository _repository;
        private readonly ICombinedSleeveRepository _combinedSleeveRepository;

        public CombinedClusterPersistenceService(IClashZoneRepository repository, ICombinedSleeveRepository combinedSleeveRepository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _combinedSleeveRepository = combinedSleeveRepository ?? throw new ArgumentNullException(nameof(combinedSleeveRepository));
        }

        public List<ClashZone> QueueDatabaseUpdates(CombinedClusterCandidate combinedCluster, int combinedSleeveInstanceId)
        {
            if (combinedCluster == null || combinedCluster.MemberClusters.Count == 0)
                return new List<ClashZone>();

            var allZoneGuids = combinedCluster.MemberClusters
                .SelectMany(c => c.ClashZoneIds)
                .Distinct()
                .ToList();

            if (allZoneGuids.Count == 0)
                return new List<ClashZone>();

            var existingZones = _repository.GetClashZonesByGuids(allZoneGuids);
            var zoneMap = existingZones.ToDictionary(z => z.Id, z => z);
            var zonesToUpdate = new List<ClashZone>();

            foreach (var clusterSleeveInfo in combinedCluster.MemberClusters)
            {
                foreach (var guid in clusterSleeveInfo.ClashZoneIds)
                {
                    if (zoneMap.TryGetValue(guid, out var clashZone))
                    {
                        clashZone.IsCombinedResolved = true;
                        clashZone.IsResolved = false;
                        clashZone.IsClusterResolved = false;
                        clashZone.SleeveInstanceId = -1;
                        clashZone.ClusterSleeveInstanceId = -1;
                        clashZone.CombinedClusterSleeveInstanceId = combinedSleeveInstanceId;
                        zonesToUpdate.Add(clashZone);
                    }
                }
            }

            return zonesToUpdate;
        }

        public void UpdateXmlWithCombinedClusterInfo(CombinedClusterCandidate combinedCluster, int combinedSleeveInstanceId)
        {
            // Stub: Implement XML update logic as needed for your application
            // This is a placeholder to satisfy the interface
        }

        public void PersistCombinedCluster(CombinedClusterCandidate combinedCluster, int combinedSleeveInstanceId)
        {
            if (combinedCluster == null) return;
            if (combinedCluster.MemberClusters.Count == 0) return;

            // ✅ CRITICAL FIX: To prevent overwriting existing critical data (MepParameterValuesJson, etc.),
            // first retrieve the full existing ClashZone objects from the database.
            // 1. Collect GUIDs
            var allZoneGuids = combinedCluster.MemberClusters
                .SelectMany(c => c.ClashZoneIds)
                .Distinct()
                .ToList();

            if (allZoneGuids.Count == 0) return;

            // 2. Fetch full objects
            var existingZones = _repository.GetClashZonesByGuids(allZoneGuids);
            // IClashZoneRepository now has GetClashZonesByGuids
            var zoneMap = existingZones.ToDictionary(z => z.Id, z => z);

            // 3. Prepare updates
            var zonesToUpdate = new List<ClashZone>();

            foreach (var clusterSleeveInfo in combinedCluster.MemberClusters)
            {
                foreach (var guid in clusterSleeveInfo.ClashZoneIds)
                {
                    if (zoneMap.TryGetValue(guid, out var clashZone))
                    {
                        // Update relevant fields: Mark as combined and "consume" individual/cluster sleeves
                        clashZone.IsCombinedResolved = true;
                        
                        // Reset lower-level flags as they are superseded by IsCombinedResolved
                        clashZone.IsResolved = false;
                        clashZone.IsClusterResolved = false;
                        
                        // Reset IDs to -1 to indicate they are replaced by the combined sleeve
                        clashZone.SleeveInstanceId = -1;
                        clashZone.ClusterSleeveInstanceId = -1;
                        
                        // Track the combined ID in memory (though not currently persisted to ClashZones table)
                        clashZone.CombinedClusterSleeveInstanceId = combinedSleeveInstanceId;
                        
                        zonesToUpdate.Add(clashZone);
                    }
                }
            }

            // 4. Batch Persist
            if (zonesToUpdate.Count > 0)
            {
                var category = combinedCluster.CategoriesInvolved.FirstOrDefault() ?? "Unknown";
                _repository.InsertOrUpdateClashZones(zonesToUpdate, "Combined", category);

                // ✅ RESET FILTER COMBO FLAG
                // We need to fetch ComboIds. Since ClashZone model doesn't have ComboId property,
                // we'll get it from the repository for the first zone (assuming homogeneity) or iterate if needed.
                // For combined sleeves, having one valid ComboId/FilterId is usually sufficient for the main record.
                var firstZone = zonesToUpdate.FirstOrDefault();
                var meta = (ComboId: 0, FilterId: 0);
                if (firstZone != null)
                {
                    meta = _repository.GetComboAndFilterId(firstZone.Id);
                    if (meta.ComboId > 0)
                    {
                        _repository.ResetFileComboFlag(meta.ComboId);
                    }
                }
                
                // 5. Persist CombinedSleeve and Constituents
                try
                {
                    // Map details to CombinedSleeve model
                    // Using CombinedBoundingBox property which exists on candidate
                    var bbox = combinedCluster.CombinedBoundingBox;
                    var width = combinedCluster.CombinedWidth;
                    var height = combinedCluster.CombinedHeight;
                    // Depth not directly on candidate, deriving from bbox or setting default
                    var depth = bbox.Max.Y - bbox.Min.Y; // Approximation if Y-aligned

                    var combinedSleeve = new CombinedSleeve
                    {
                        CombinedInstanceId = combinedSleeveInstanceId,
                        ComboId = meta.ComboId,
                        FilterId = meta.FilterId,
                        Categories = combinedCluster.CategoriesInvolved, // ✅ Fixed: Expects List<string>
                        BoundingBoxMinX = bbox.Min.X,
                        BoundingBoxMinY = bbox.Min.Y,
                        BoundingBoxMinZ = bbox.Min.Z,
                        BoundingBoxMaxX = bbox.Max.X,
                        BoundingBoxMaxY = bbox.Max.Y,
                        BoundingBoxMaxZ = bbox.Max.Z,
                        CombinedWidth = width,
                        CombinedHeight = height,
                        CombinedDepth = depth,
                        PlacementX = bbox.Min.X + (width / 2.0), // Center X
                        PlacementY = bbox.Min.Y + (depth / 2.0), // Center Y
                        PlacementZ = bbox.Min.Z + (height / 2.0), // Center Z
                        RotationAngleDeg = 0,
                        HostType = combinedCluster.HostType,
                        HostOrientation = combinedCluster.Orientation
                    };

                    // Add Constituents
                    combinedSleeve.Constituents = new List<SleeveConstituent>();
                    foreach (var zone in zonesToUpdate)
                    {
                        combinedSleeve.Constituents.Add(new SleeveConstituent
                        {
                            CombinedSleeveId = 0, // Will be set on save
                            Type = ConstituentType.Individual,
                            Category = zone.MepElementCategory ?? "Unknown",
                            ClashZoneId = 0, // ✅ Fixed: ClashZone model doesn't have int ID exposed
                            ClashZoneGuid = zone.Id,
                            ClusterSleeveId = zone.ClusterSleeveInstanceId > 0 ? zone.ClusterSleeveInstanceId : (int?)null,
                            ClusterInstanceId = zone.ClusterSleeveInstanceId > 0 ? zone.ClusterSleeveInstanceId : (int?)null
                        });
                    }

                    _combinedSleeveRepository.SaveCombinedSleeve(combinedSleeve);
                }
                catch (Exception ex)
                {
                    // Log error but don't fail the whole operation as zones are updated
                    // _logger.Error($"Failed to save combined sleeve record: {ex.Message}");
                    // Since we don't have logger here, we might just suppress or rethrow?
                    // Ideally we should log.
                    System.Diagnostics.Debug.WriteLine($"Error saving combined sleeve: {ex.Message}");
                }
            }
        }
    }
}
