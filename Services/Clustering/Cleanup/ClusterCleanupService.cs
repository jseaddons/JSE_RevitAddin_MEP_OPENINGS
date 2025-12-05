using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup
{
    /// <summary>
    /// Extracted cleanup logic from UniversalClusterService. Simplified: retains safety checks and protection set.
    /// </summary>
    public class ClusterCleanupService : IClusterCleanupService
    {
        public int CleanupSleevesWithinClusters(Document doc, List<FamilyInstance> placedClusters)
        {
            int deletedCount = 0;
            try
            {
                if (placedClusters == null || placedClusters.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: No placed clusters, skipping cleanup\n");
                    return 0;
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Starting cleanup with {placedClusters.Count} placed clusters\n");

                // Build protection set of cluster sleeve IDs
                var clusterSleeveIds = new HashSet<int>(placedClusters.Where(c => c != null).Select(c => c.Id.IntegerValue));
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Protection set contains {clusterSleeveIds.Count} cluster sleeve IDs: {string.Join(", ", clusterSleeveIds)}\n");

                // Collect all potential sleeve family instances
                var allSleeves = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(s =>
                    {
                        string fam = s.Symbol?.FamilyName ?? string.Empty;
                        bool keyword = fam.Contains("Sleeve", StringComparison.OrdinalIgnoreCase) || fam.Contains("Opening", StringComparison.OrdinalIgnoreCase);
                        bool known = fam.Contains("CircularOpening", StringComparison.OrdinalIgnoreCase) || fam.Contains("RectangularOpening", StringComparison.OrdinalIgnoreCase);
                        return (s.Category?.Name == "Generic Models" || s.Category?.Name == "Structural Connections") && (keyword || known);
                    })
                    .ToList();

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Found {allSleeves.Count} total sleeves in document\n");

                var individualSleeves = new List<FamilyInstance>();
                foreach (var s in allSleeves)
                {
                    int id = s.Id.IntegerValue;
                    if (clusterSleeveIds.Contains(id))
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Sleeve {id} is in protection set (cluster sleeve) - SKIPPING\n");
                        continue; // protected
                    }
                    var clusterParam = s.LookupParameter("Cluster Sleeve Instance ID");
                    var sleeveInstanceParam = s.LookupParameter("Sleeve Instance ID");
                    int clusterValue = clusterParam?.AsInteger() ?? -1;
                    int sleeveInstanceValue = sleeveInstanceParam?.AsInteger() ?? -999;
                    bool isClusterSleeve = (sleeveInstanceValue == -1) || (clusterValue > 0 && clusterValue == id);
                    if (isClusterSleeve)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Sleeve {id} is identified as cluster sleeve (ClusterParam={clusterValue}, SleeveInstanceId={sleeveInstanceValue}) - SKIPPING\n");
                        continue;
                    }
                    // ✅ DIAGNOSTIC: Log individual sleeve parameters for debugging
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Individual sleeve {id} added to check list - ClusterParam={clusterValue}, SleeveInstanceId={sleeveInstanceValue}\n");
                    individualSleeves.Add(s);
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Found {individualSleeves.Count} individual sleeves to check (IDs: {string.Join(", ", individualSleeves.Select(s => s.Id.IntegerValue))})\n");

                if (individualSleeves.Count == 0) return 0;

                // Build cluster sleeve bounding boxes (axis-aligned)
                // ✅ CRITICAL: Use active view for bounding box calculation (matches methodology requirement)
                // According to methodology: "Check for other individual sleeves that fall within the cluster bounding box"
                // The bounding box must be calculated from the actual placed cluster sleeve element
                var clusterBboxes = new List<BoundingBoxXYZ>();
                var clusterBboxMap = new Dictionary<int, BoundingBoxXYZ>(); // Map cluster ID to bbox for logging
                foreach (var cluster in placedClusters)
                {
                    if (cluster == null) continue;
                    int clusterId = cluster.Id.IntegerValue;
                    
                    // ✅ CRITICAL FIX: Refresh element and get fresh bounding box
                    // Sometimes bounding boxes are stale after placement, so we get the element fresh
                    var freshCluster = doc.GetElement(cluster.Id) as FamilyInstance;
                    if (freshCluster == null)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: Cluster {clusterId} not found in document - SKIPPING\n");
                        continue;
                    }
                    
                    // ✅ CRITICAL FIX: Calculate bounding box from placement point + parameters
                    // Don't rely on get_BoundingBox() which may be stale if parameters haven't been flushed/regenerated
                    // Get placement point from instance
                    XYZ placementPoint = null;
                    if (freshCluster.Location is LocationPoint locationPoint)
                    {
                        placementPoint = locationPoint.Point;
                    }
                    else if (freshCluster.Location is LocationCurve locationCurve)
                    {
                        var curve = locationCurve.Curve;
                        if (curve != null)
                        {
                            placementPoint = curve.GetEndPoint(0);
                        }
                    }
                    
                    // Get dimensions from parameters
                    var widthParam = freshCluster.LookupParameter("Width");
                    var heightParam = freshCluster.LookupParameter("Height");
                    var depthParam = freshCluster.LookupParameter("Depth");
                    double paramWidth = widthParam?.AsDouble() ?? 0.0;
                    double paramHeight = heightParam?.AsDouble() ?? 0.0;
                    double paramDepth = depthParam?.AsDouble() ?? 0.0;
                    
                    // ✅ CRITICAL: If parameters are valid and placement point exists, calculate bounding box
                    if (placementPoint != null && paramWidth > 0.001 && paramHeight > 0.001 && paramDepth > 0.001)
                    {
                        // Calculate bounding box from placement point + dimensions
                        double halfWidth = paramWidth / 2.0;
                        double halfHeight = paramHeight / 2.0;
                        double halfDepth = paramDepth / 2.0;
                        
                        var calculatedBbox = new BoundingBoxXYZ();
                        calculatedBbox.Min = new XYZ(
                            placementPoint.X - halfWidth,
                            placementPoint.Y - halfHeight,
                            placementPoint.Z - halfDepth
                        );
                        calculatedBbox.Max = new XYZ(
                            placementPoint.X + halfWidth,
                            placementPoint.Y + halfHeight,
                            placementPoint.Z + halfDepth
                        );
                        calculatedBbox.Enabled = true;
                        
                        clusterBboxes.Add(calculatedBbox);
                        clusterBboxMap[clusterId] = calculatedBbox;
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Cluster {clusterId} bbox (CALCULATED from params) - " +
                            $"PlacementPoint=({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3}), " +
                            $"Min=({calculatedBbox.Min.X:F3}, {calculatedBbox.Min.Y:F3}, {calculatedBbox.Min.Z:F3}), " +
                            $"Max=({calculatedBbox.Max.X:F3}, {calculatedBbox.Max.Y:F3}, {calculatedBbox.Max.Z:F3}), " +
                            $"BBoxSize=({paramWidth*304.8:F1}mm x {paramHeight*304.8:F1}mm x {paramDepth*304.8:F1}mm), " +
                            $"Params=(W={paramWidth*304.8:F1}mm, H={paramHeight*304.8:F1}mm, D={paramDepth*304.8:F1}mm)\n");
                    }
                    else
                    {
                        // Fallback to get_BoundingBox() if parameters are missing
                        var bbox = freshCluster.get_BoundingBox(null);
                        if (bbox != null && bbox.Enabled)
                        {
                            clusterBboxes.Add(bbox);
                            clusterBboxMap[clusterId] = bbox;
                            
                            double bboxWidth = (bbox.Max.X - bbox.Min.X) * 304.8;
                            double bboxHeight = (bbox.Max.Y - bbox.Min.Y) * 304.8;
                            double bboxDepth = (bbox.Max.Z - bbox.Min.Z) * 304.8;
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Cluster {clusterId} bbox (FALLBACK from get_BoundingBox) - " +
                                $"Min=({bbox.Min.X:F3}, {bbox.Min.Y:F3}, {bbox.Min.Z:F3}), " +
                                $"Max=({bbox.Max.X:F3}, {bbox.Max.Y:F3}, {bbox.Max.Z:F3}), " +
                                $"BBoxSize=({bboxWidth:F1}mm x {bboxHeight:F1}mm x {bboxDepth:F1}mm), " +
                                $"Params=(W={paramWidth*304.8:F1}mm, H={paramHeight*304.8:F1}mm, D={paramDepth*304.8:F1}mm)\n");
                        }
                        else
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: Cluster {clusterId} has no valid placement point or parameters - SKIPPING\n");
                        }
                    }
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Built {clusterBboxes.Count} cluster bounding boxes from {placedClusters.Count} placed clusters\n");

                if (clusterBboxes.Count == 0) return 0;

                var toDelete = new List<ElementId>();
                foreach (var individual in individualSleeves)
                {
                    int individualId = individual.Id.IntegerValue;
                    
                    // ✅ CRITICAL SAFETY CHECK: Double-check this is NOT a cluster sleeve
                    // This prevents any cluster sleeve from being marked for deletion, even if it somehow got into individualSleeves
                    if (clusterSleeveIds.Contains(individualId))
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ SAFETY CHECK: Sleeve {individualId} found in individualSleeves but is in protection set (cluster sleeve) - SKIPPING containment check\n");
                        continue;
                    }
                    
                    // ✅ IMPROVED LOGIC: Get individual sleeve bounding box for overlap check
                    // Check if individual sleeve's bounding box overlaps with cluster bounding box
                    // This is more robust than just checking placement point
                    var individualBbox = individual.get_BoundingBox(null);
                    if (individualBbox == null || !individualBbox.Enabled)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Individual sleeve {individualId} has no bounding box or is disabled - SKIPPING\n");
                        continue;
                    }
                    
                    // Get placement point for logging
                    XYZ sleevePlacementPoint = null;
                    if (individual.Location is LocationPoint locationPoint)
                    {
                        sleevePlacementPoint = locationPoint.Point;
                    }
                    else
                    {
                        sleevePlacementPoint = (individualBbox.Min + individualBbox.Max) / 2.0;
                    }
                    
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Checking individual sleeve {individualId} - " +
                        $"PlacementPoint=({sleevePlacementPoint.X:F3}, {sleevePlacementPoint.Y:F3}, {sleevePlacementPoint.Z:F3}), " +
                        $"BBox=({individualBbox.Min.X:F3},{individualBbox.Min.Y:F3},{individualBbox.Min.Z:F3}) to ({individualBbox.Max.X:F3},{individualBbox.Max.Y:F3},{individualBbox.Max.Z:F3})\n");
                    
                    bool matchedAnyCluster = false;
                    int clusterIndex = 0;
                    foreach (var kvp in clusterBboxMap)
                    {
                        int clusterId = kvp.Key;
                        var cbbox = kvp.Value;
                        clusterIndex++;
                        if (cbbox == null) continue;
                        
                        // ✅ IMPROVED LOGIC: Check bounding box overlap instead of just placement point
                        // Two bounding boxes overlap if they intersect on all three axes
                        // This catches edge cases where placement point is outside but bounding box overlaps
                        bool bboxOverlaps = individualBbox.Max.X >= cbbox.Min.X && individualBbox.Min.X <= cbbox.Max.X &&
                                           individualBbox.Max.Y >= cbbox.Min.Y && individualBbox.Min.Y <= cbbox.Max.Y &&
                                           individualBbox.Max.Z >= cbbox.Min.Z && individualBbox.Min.Z <= cbbox.Max.Z;
                        
                        // Also check if placement point is inside (for small sleeves where bbox might not overlap)
                        bool placementPointInside = sleevePlacementPoint.X >= cbbox.Min.X && sleevePlacementPoint.X <= cbbox.Max.X &&
                                                    sleevePlacementPoint.Y >= cbbox.Min.Y && sleevePlacementPoint.Y <= cbbox.Max.Y &&
                                                    sleevePlacementPoint.Z >= cbbox.Min.Z && sleevePlacementPoint.Z <= cbbox.Max.Z;
                        
                        // Delete if either bounding box overlaps OR placement point is inside
                        bool shouldDelete = bboxOverlaps || placementPointInside;
                        
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP:   vs Cluster #{clusterIndex} (ID={clusterId}) bbox - " +
                            $"Cluster=({cbbox.Min.X:F3},{cbbox.Min.Y:F3},{cbbox.Min.Z:F3}) to ({cbbox.Max.X:F3},{cbbox.Max.Y:F3},{cbbox.Max.Z:F3}), " +
                            $"IndividualBBox=({individualBbox.Min.X:F3},{individualBbox.Min.Y:F3},{individualBbox.Min.Z:F3}) to ({individualBbox.Max.X:F3},{individualBbox.Max.Y:F3},{individualBbox.Max.Z:F3}), " +
                            $"PlacementPoint=({sleevePlacementPoint.X:F3},{sleevePlacementPoint.Y:F3},{sleevePlacementPoint.Z:F3}), " +
                            $"bboxOverlaps={bboxOverlaps}, placementPointInside={placementPointInside}, shouldDelete={shouldDelete}\n");
                        
                        // ✅ CRITICAL: Delete if bounding box overlaps OR placement point is inside cluster bbox
                        if (shouldDelete)
                        {
                            // ✅ CRITICAL SAFETY CHECK: Triple-check this is NOT a cluster sleeve before marking for deletion
                            if (clusterSleeveIds.Contains(individualId))
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL SAFETY: Sleeve {individualId} matched containment check but is a cluster sleeve - NOT MARKING FOR DELETION\n");
                                matchedAnyCluster = true;
                                break;
                            }
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Individual sleeve {individualId} placement point is inside cluster #{clusterIndex} (ID={clusterId}) bbox - MARKING FOR DELETION\n");
                            toDelete.Add(individual.Id);
                            matchedAnyCluster = true;
                            break;
                        }
                    }
                    
                    if (!matchedAnyCluster)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Individual sleeve {individualId} placement point is NOT inside any cluster bbox - KEEPING\n");
                    }
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Found {toDelete.Count} individual sleeves to delete\n");

                if (toDelete.Count == 0) return 0;

                // ✅ CRITICAL PROTECTION: Final check before deletion - ensure NO cluster sleeve IDs are in toDelete list
                // This is a safety net in case filtering logic above failed
                var protectedIds = new HashSet<int>(clusterSleeveIds);
                var filteredToDelete = toDelete.Where(id => !protectedIds.Contains(id.IntegerValue)).ToList();
                int removedClusterSleeves = toDelete.Count - filteredToDelete.Count;
                
                if (removedClusterSleeves > 0)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL PROTECTION: Removed {removedClusterSleeves} cluster sleeve ID(s) from deletion list before batch delete!\n");
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Protection set: {string.Join(", ", protectedIds)}\n");
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Original toDelete IDs: {string.Join(", ", toDelete.Select(id => id.IntegerValue))}\n");
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Filtered toDelete IDs: {string.Join(", ", filteredToDelete.Select(id => id.IntegerValue))}\n");
                }
                
                if (filteredToDelete.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: After protection filter, no sleeves to delete\n");
                    return 0;
                }

                // ✅ CRITICAL: Check if we're already in a transaction
                // If doc.IsModifiable is true, we're in a transaction - don't create a new one
                bool alreadyInTransaction = doc.IsModifiable;
                
                if (alreadyInTransaction)
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Already in transaction, batch deleting {filteredToDelete.Count} sleeves (protected {removedClusterSleeves} cluster sleeves)\n");
                    
                    // ✅ PERFORMANCE: Use batch delete instead of deleting one by one
                    try
                    {
                        var deletedIds = doc.Delete(filteredToDelete);
                        deletedCount = deletedIds.Count;
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Batch deleted {deletedCount} individual sleeves (IDs: {string.Join(", ", deletedIds.Select(id => id.IntegerValue))})\n");
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Batch delete failed: {ex.Message}, falling back to individual deletes\n");
                        
                        // Fallback to individual deletes if batch fails
                        foreach (var id in filteredToDelete)
                        {
                            try
                            {
                                // ✅ FINAL PROTECTION: Double-check before individual delete
                                if (protectedIds.Contains(id.IntegerValue))
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL PROTECTION: Skipping cluster sleeve {id.IntegerValue} in fallback delete!\n");
                                    continue;
                                }
                                
                                doc.Delete(id);
                                deletedCount++;
                            }
                            catch (Exception delEx)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Failed deleting sleeve {id.IntegerValue}: {delEx.Message}\n");
                                DebugLogger.Warning($"[ClusterCleanupService] Failed deleting sleeve {id.IntegerValue}: {delEx.Message}");
                            }
                        }
                    }
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Not in transaction, creating new transaction for batch delete of {filteredToDelete.Count} sleeves (protected {removedClusterSleeves} cluster sleeves)\n");
                    
                    using (var tx = new Transaction(doc, "Delete Individual Sleeves Within Clusters"))
                    {
                        tx.Start();
                        
                        // ✅ PERFORMANCE: Use batch delete instead of deleting one by one
                        try
                        {
                            var deletedIds = doc.Delete(filteredToDelete);
                            deletedCount = deletedIds.Count;
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Batch deleted {deletedCount} individual sleeves (IDs: {string.Join(", ", deletedIds.Select(id => id.IntegerValue))})\n");
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Batch delete failed: {ex.Message}, falling back to individual deletes\n");
                            
                            // Fallback to individual deletes if batch fails
                            foreach (var id in filteredToDelete)
                            {
                                try
                                {
                                    // ✅ FINAL PROTECTION: Double-check before individual delete
                                    if (protectedIds.Contains(id.IntegerValue))
                                    {
                                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                            $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL PROTECTION: Skipping cluster sleeve {id.IntegerValue} in fallback delete!\n");
                                        continue;
                                    }
                                    
                                    doc.Delete(id);
                                    deletedCount++;
                                }
                                catch (Exception delEx)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                        $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ Failed deleting sleeve {id.IntegerValue}: {delEx.Message}\n");
                                    DebugLogger.Warning($"[ClusterCleanupService] Failed deleting sleeve {id.IntegerValue}: {delEx.Message}");
                                }
                            }
                        }
                        
                        tx.Commit();
                    }
                }
                
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ✅ Completed cleanup: Deleted {deletedCount} individual sleeves\n");
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: ❌ ERROR during cleanup: {ex.Message}\n{ex.StackTrace}\n");
                DebugLogger.Error($"[ClusterCleanupService] Error during cleanup: {ex.Message}");
            }
            return deletedCount;
        }

        public void ResetClusterFlagsForDeletedSleeves(Document doc, string? xmlFilePath = null)
        {
            try
            {
                int resetCount = 0;
                var resetClashZones = new List<ClashZone>();
                var oldClusterInstanceIds = new Dictionary<Guid, int>(); // For potential future diagnostics
                var categories = new[] { "Ducts", "Pipes", "Cable Trays", "Duct Accessories", "Cable Tray Fittings" };

                using var context = new SleeveDbContext(doc, _ => { });
                var repository = new ClashZoneRepository(context, _ => { });

                foreach (var category in categories)
                {
                    var dbZones = repository.GetClashZonesByCategory(category)
                        ?.Where(z => z != null && z.IsClusterResolved && z.ClusterSleeveInstanceId > 0)
                        .ToList();
                    if (dbZones == null || dbZones.Count == 0) continue;
                    foreach (var clashZone in dbZones)
                    {
                        var clusterSleeveId = new ElementId(clashZone.ClusterSleeveInstanceId);
                        var clusterSleeve = doc.GetElement(clusterSleeveId);
                        if (clusterSleeve == null)
                        {
                            oldClusterInstanceIds[clashZone.Id] = clashZone.ClusterSleeveInstanceId;
                            clashZone.IsClusterResolved = false;
                            clashZone.ClusterSleeveInstanceId = -1;
                            resetClashZones.Add(clashZone);
                            resetCount++;
                        }
                    }
                }

                if (resetClashZones.Count > 0)
                {
                    // Direct DB update (repository has no bulk UpdateClashZones API)
                    using (var tx = context.Connection.BeginTransaction())
                    {
                        foreach (var z in resetClashZones)
                        {
                            try
                            {
                                using var cmd = context.Connection.CreateCommand();
                                cmd.Transaction = tx;
                                cmd.CommandText = @"UPDATE ClashZones
                                    SET IsClusterResolvedFlag = 0,
                                        ClusterInstanceId = -1,
                                        UpdatedAt = CURRENT_TIMESTAMP
                                    WHERE UPPER(ClashZoneGuid) = UPPER(@Guid)
                                      AND ClashZoneGuid != '' AND ClashZoneGuid IS NOT NULL";
                                cmd.Parameters.AddWithValue("@Guid", z.Id.ToString());
                                cmd.ExecuteNonQuery();
                            }
                            catch (Exception updEx)
                            {
                                DebugLogger.Warning($"[ClusterCleanupService] Failed DB flag reset for zone {z.Id}: {updEx.Message}");
                            }
                        }
                        tx.Commit();
                    }
                }
                DebugLogger.Info($"[ClusterCleanupService] Reset {resetCount} cluster flags for deleted sleeves.");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ClusterCleanupService] Error resetting cluster flags: {ex.Message}");
            }
        }
    }
}
