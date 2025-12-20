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
             // 2025-12-18: REFACTORED
             // This method previously queued updates for the background worker.
             // However, with the new 'UpdateCombinedResolutionFlags' direct database method,
             // updates must be synchronous to prevent race conditions during subsequent placements.
             // Therefore, this method is now a no-op as the updates are handled in PersistCombinedCluster.
             // We return empty list to signify no zones need further memory-queue processing.
             return new List<ClashZone>();
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

            // 3. Update resolution flags efficiently using the new repository method
            // This prevents overwriting other properties of the ClashZone and focuses purely on flag management
            // "Update Existing, Do Not Create New"
            _repository.UpdateCombinedResolutionFlags(allZoneGuids, combinedSleeveInstanceId);

            // Re-fetch zones to continue with other logic if necessary, or just rely on GUIDs
            // For the following logic (ResetFileComboFlag and CombinedSleeve creation), we need at least one zone to get metadata
            var firstZoneGuid = allZoneGuids.FirstOrDefault();
            
            // 4. Batch Persist - NO LONGER NEEDED for flags as we used UpdateCombinedResolutionFlags
            // But we still need to persist the CombinedSleeve entity below/

            // ✅ RESET FILTER COMBO FLAG
            if (firstZoneGuid != Guid.Empty)
            {
                var meta = _repository.GetComboAndFilterId(firstZoneGuid);
                if (meta.ComboId > 0)
                {
                    _repository.ResetFileComboFlag(meta.ComboId);
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
                    
                    // We need basic info for constituents. Since we skipped full object fetch for optimization,
                    // we can either fetch them now or use available info. 
                    // Let's fetch them now as it is safer for constituent data accuracy.
                    var existingZones = _repository.GetClashZonesByGuids(allZoneGuids);
                    
                    foreach (var zone in existingZones)
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
                    System.Diagnostics.Debug.WriteLine($"Error saving combined sleeve: {ex.Message}");
                }
            }


        }
    }
}
