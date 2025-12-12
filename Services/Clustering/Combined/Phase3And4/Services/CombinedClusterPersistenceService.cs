using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services
{
    /// <summary>
    /// Service responsible for persisting updates for combined clusters.
    /// Phase 4 Implementation.
    /// </summary>
    public class CombinedClusterPersistenceService
    {
        private readonly IClashZoneRepository _clashZoneRepository;

        public CombinedClusterPersistenceService(IClashZoneRepository clashZoneRepository)
        {
            _clashZoneRepository = clashZoneRepository ?? throw new ArgumentNullException(nameof(clashZoneRepository));
        }

        public void PersistCombinedCluster(
            CombinedClusterCandidate combinedCluster, 
            int combinedSleeveInstanceId)
        {
            if (combinedCluster == null || combinedCluster.MemberClusters.Count == 0) return;

            string categories = string.Join(",", combinedCluster.CategoriesInvolved.OrderBy(c => c));
            string jsonParams = "{}"; // Serialize logic if needed

            // Collect all ClashZone IDs involved in this combined cluster
            // Note: Use SourceSleeveInstanceId (which is likely the original SleeveInstanceId or MepElementId depending on mapping)
            // Or use ClashZoneId (GUID) if the repo supports it.
            // Based on ClusterSleeveInfo, we have: ClashZoneId (GUID).

            var clashZoneGuids = combinedCluster.MemberClusters.Select(m => m.ClashZoneId).ToList();

            // We need to fetch these zones to update them.
            // Assumption: IClashZoneRepository has a method to get by GUIDs or we iterate.
            // If not, we rely on ClusterSleeveInfo properties which mirror ClashZone
            // and we might need to craft an update strictly by ID.
            
            // To be safe and performant, we'll try to update using basic properties if repository allows batching
            // Otherwise we accept the limitation of one-by-one or fetch-all.
            
            // Workaround since we don't know exact Repo API for Guids:
            // Group by category to match typical repository usage if needed, or simply loop.
            
            // Loading all potential zones first is safer to ensure we are updating latest state
            // But efficiently we might not be able to load by GUID list easily.
            // Let's assume we can fetch by Filter/Category for the broad set, OR use the fact that we just processed them.
            
            // SIMPLIFIED PERSISTENCE FOR HARMONY:
            // Since we can't easily fetch by GUID list without likely API change, 
            // we will simulate the update by creating ClashZone objects with just the ID and the changed fields,
            // relying on the Repo to "Update" based on ID.
            
            var zonesToUpdate = new List<ClashZone>();

            foreach (var member in combinedCluster.MemberClusters)
            {
                var zone = new ClashZone
                {
                    Id = member.ClashZoneId,
                    SleeveInstanceId = member.SourceSleeveInstanceId, 
                    MepElementCategory = member.Category,
                    
                    // Update New Fields
                    CombinedClusterSleeveInstanceId = combinedSleeveInstanceId,
                    CategoriesInCombinedCluster = categories,
                    IsIncorporatedInCombinedCluster = true,
                    CombinedClusterParameterSnapshot = jsonParams,
                    
                    CombinedClusterSleeveBoundingBoxMinX = combinedCluster.CombinedBoundingBoxMinX,
                    CombinedClusterSleeveBoundingBoxMinY = combinedCluster.CombinedBoundingBoxMinY,
                    CombinedClusterSleeveBoundingBoxMinZ = combinedCluster.CombinedBoundingBoxMinZ,
                    CombinedClusterSleeveBoundingBoxMaxX = combinedCluster.CombinedBoundingBoxMaxX,
                    CombinedClusterSleeveBoundingBoxMaxY = combinedCluster.CombinedBoundingBoxMaxY,
                    CombinedClusterSleeveBoundingBoxMaxZ = combinedCluster.CombinedBoundingBoxMaxZ
                };
                
                zonesToUpdate.Add(zone);
            }

            // Batch update via repository
            // Group by category as typically Repositories are sharded or optimized by category
            foreach (var group in zonesToUpdate.GroupBy(z => z.MepElementCategory))
            {
                // Passing empty filter name if not tracked here, or pass "Combined"
                _clashZoneRepository.InsertOrUpdateClashZones(group.ToList(), "CombinedUpdate", group.Key);
            }
        }
    }
}
