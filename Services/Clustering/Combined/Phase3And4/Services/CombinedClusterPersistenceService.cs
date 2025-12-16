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

        public CombinedClusterPersistenceService(IClashZoneRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
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
                        clashZone.IsClusterResolved = true;
                        clashZone.ClusterSleeveInstanceId = combinedSleeveInstanceId;
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
                        // Update relevant fields
                        clashZone.IsClusterResolved = true;
                        clashZone.ClusterSleeveInstanceId = combinedSleeveInstanceId;
                        zonesToUpdate.Add(clashZone);
                    }
                }
            }

            // 4. Batch Persist
            if (zonesToUpdate.Count > 0)
            {
                var category = combinedCluster.CategoriesInvolved.FirstOrDefault() ?? "Unknown";
                _repository.InsertOrUpdateClashZones(zonesToUpdate, "Combined", category);
            }
        }
    }
}
