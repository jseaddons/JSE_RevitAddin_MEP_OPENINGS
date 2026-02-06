using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB; // ✅ Added for Element access
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
        private readonly Document _document; // ✅ Added Document

        public CombinedClusterPersistenceService(IClashZoneRepository repository, ICombinedSleeveRepository combinedSleeveRepository, Document document)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _combinedSleeveRepository = combinedSleeveRepository ?? throw new ArgumentNullException(nameof(combinedSleeveRepository));
            _document = document ?? throw new ArgumentNullException(nameof(document));
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

            // ✅ CRITICAL FIX: To prevent overwriting existing critical data and ensure ROBUST flag updates,
            // we split the update into two efficient database calls:
            // 1. Update flags for Individual Constituents (by GUID)
            // 2. Update flags for Cluster Constituents (by ClusterInstanceId - robust against missing JSON linkage)

            var individualZoneGuids = new List<Guid>();
            var clusterInstanceIds = new List<int>();
            var allZoneGuids = new List<Guid>(); // Still collect for potential downstream logic

            foreach (var cluster in combinedCluster.MemberClusters)
            {
                // Collect GUIDs if available (for "allZoneGuids" usage)
                if (cluster.ClashZoneIds != null)
                {
                    allZoneGuids.AddRange(cluster.ClashZoneIds);
                }

                // Determine type and collect ID for flag update
                if (cluster.ClusterSleeveInstanceId > 0)
                {
                    clusterInstanceIds.Add(cluster.ClusterSleeveInstanceId);
                }
                else
                {
                    // It's an individual sleeve (or synthetic cluster without ID, but unlikely here)
                    if (cluster.ClashZoneIds != null && cluster.ClashZoneIds.Count > 0)
                    {
                        individualZoneGuids.AddRange(cluster.ClashZoneIds);
                    }
                }
            }

            allZoneGuids = allZoneGuids.Distinct().ToList();
            individualZoneGuids = individualZoneGuids.Distinct().ToList();
            clusterInstanceIds = clusterInstanceIds.Distinct().ToList();

            if (individualZoneGuids.Count == 0 && clusterInstanceIds.Count == 0) return;

            // 1. Update Individuals
            if (individualZoneGuids.Count > 0)
            {
                _repository.UpdateCombinedResolutionFlags(individualZoneGuids, combinedSleeveInstanceId);
            }

            // 2. Update Clusters (ROBUST FIX)
            if (clusterInstanceIds.Count > 0)
            {
                _repository.UpdateCombinedResolutionFlagsByClusterIds(clusterInstanceIds, combinedSleeveInstanceId);
            }

            // Re-fetch zones to continue with other logic if necessary, or just rely on GUIDs
            // For the following logic (ResetFileComboFlag and CombinedSleeve creation), we need at least one zone to get metadata
            var firstZoneGuid = allZoneGuids.FirstOrDefault();
            
            // 4. Batch Persist - NO LONGER NEEDED for flags as we used UpdateCombinedResolutionFlags
            // But we still need to persist the CombinedSleeve entity below/
            
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
                    
                    // We need basic info for constituents. Since we skipped full object fetch for optimization,
                    // we can either fetch them now or use available info. 
                    // Let's fetch them now as it is safer for constituent data accuracy and Needed for Thickness calculation/
                    var existingZones = _repository.GetClashZonesByGuids(allZoneGuids);

                    // ✅ FIXED DEPTH CALCULATION: Use actual structural element thickness instead of bbox approximation
                    // This fixes "Extrusion is too thin" error for non-Y-aligned elements (e.g. cable trays, floors)
                    double maxThickness = 0.0;
                    try 
                    {
                        if (existingZones.Any())
                        {
                            maxThickness = existingZones.Max(z => {
                                try 
                                {
                                    // Prefer SleeveDepth if populated (from previous calculation)
                                    if (z.SleeveDepth > 0.01) return z.SleeveDepth;
                                    
                                    // Otherwise fallback to querying the element from Document
                                    var structId = new ElementId(z.StructuralElementIdValue);
                                    var structElem = _document.GetElement(structId);
                                    if (structElem != null)
                                    {
                                        return GetElementThickness(structElem);
                                    }
                                    return 0.0;
                                }
                                catch { return 0.0; }
                            });
                        }
                    } 
                    catch { }

                    // Apply min valid depth (e.g. 6 inches) if calculation fails or returns 0
                    // But if maxThickness is valid, use it. Bbox approximation is fallback-fallback.
                    var depth = maxThickness > 0.001 ? maxThickness : (bbox.Max.Y - bbox.Min.Y);
                    
                    if (depth < 0.001) depth = 0.5; // Final safety net: 6 inches

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

        /// <summary>
        /// Get element thickness for walls, floors, and structural framing.
        /// Adapted from ClashZoneService logic.
        /// </summary>
        private double GetElementThickness(Element element)
        {
            try
            {
                if (element is Wall wall)
                {
                    return wall.Width;
                }
                else if (element is Floor floor)
                {
                    return floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?.AsDouble() ?? 0.1;
                }
                else if (element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    // Structural framing: read TYPE parameter 'b'/'d'/'Width'/'Depth'
                    try
                    {
                        ElementId typeId = element.GetTypeId();
                        var typeElem = _document.GetElement(typeId);

                        if (typeElem == null) return 0.1;

                        Parameter p = null;
                        double val = 0.0;

                        // Try multiple parameter names
                        string[] possibleNames = { "b", "B", "Breadth", "Depth", "Width", "Height", "d", "D" };

                        foreach (var paramName in possibleNames)
                        {
                            p = typeElem.LookupParameter(paramName);
                            if (p != null && !p.IsReadOnly && p.StorageType == StorageType.Double)
                            {
                                val = p.AsDouble();
                                if (val > 0) return val;
                            }
                        }
                    }
                    catch { }
                }
                
                return 0.1; // Default fallback
            }
            catch
            {
                return 0.1;
            }
        }
    }
}
