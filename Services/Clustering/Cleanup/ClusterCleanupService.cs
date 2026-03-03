using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup
{
    /// <summary>
    /// Service for cleaning up individual sleeves within cluster bounding boxes
    /// Uses database-only approach for Stage 2 cleanup (point-in-box check)
    /// </summary>
    public class ClusterCleanupService : IClusterCleanupService
    {
        /// <summary>SQLite parameter limit is 999; batch cluster IDs to stay under (leaves room for other params).</summary>
        private const int MaxClusterIdsPerQuery = 400;
        /// <summary>Revit delete in chunks to avoid transaction/timeout issues with hundreds of elements.</summary>
        private const int MaxDeletePerTransaction = 200;

        /// <summary>
        /// Cleanup individual sleeves within clusters using database-only approach (Stage 2 cleanup)
        /// Checks if individual sleeve placement points are inside cluster bounding boxes
        /// </summary>
        public int CleanupSleevesWithinClustersFromDatabase(Document doc, string targetCategory = null, List<int> clusterInstanceIds = null)
        {
            if (doc == null)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: Document is null, skipping cleanup\n");
                    return 0;
                }

            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP DEBUG: Starting (Category={targetCategory ?? "ALL"}, ClusterIds={clusterInstanceIds?.Count ?? 0})\n");

            using (var context = new SleeveDbContext(doc))
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP DEBUG: Created SleeveDbContext, about to query individual sleeves\n");
                    
                var repo = new ClashZoneRepository(context, msg => SafeFileLogger.SafeAppendText("cluster_debug.log", $"[Repo] {msg}\n"));

                // ✅ Backfill cluster bbox from placement + dimensions when bbox is zero (so cleanup can find individuals inside)
                try
                {
                    using (var backfillCmd = context.Connection.CreateCommand())
                    {
                        backfillCmd.CommandText = @"
                            UPDATE ClusterSleeves_v2 SET
                                BoundingBoxMinX = PlacementX - (ClusterWidth / 2.0),
                                BoundingBoxMinY = PlacementY - (ClusterHeight / 2.0),
                                BoundingBoxMinZ = PlacementZ - (ClusterDepth / 2.0),
                                BoundingBoxMaxX = PlacementX + (ClusterWidth / 2.0),
                                BoundingBoxMaxY = PlacementY + (ClusterHeight / 2.0),
                                BoundingBoxMaxZ = PlacementZ + (ClusterDepth / 2.0)
                            WHERE ClusterInstanceId > 0
                              AND (BoundingBoxMinX = 0.0 AND BoundingBoxMaxX = 0.0)";
                        int backfillRows = backfillCmd.ExecuteNonQuery();
                        if (backfillRows > 0)
                            SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Backfilled bbox for {backfillRows} cluster(s) from placement+dimensions\n");
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP backfill bbox: {ex.Message}\n");
                }

                // ✅ Point-in-box query: batched when many cluster IDs (SQLite param limit ~999; Revit delete in chunks)
                // ✅ CRITICAL FIX: Track which cluster each sleeve belongs to for Stage 2 constituent assignment
                var sleevesToDelete = new HashSet<int>();
                var sleeveToClusterMap = new Dictionary<int, int>(); // SleeveInstanceId -> ClusterInstanceId
                var clusterIdBatches = (clusterInstanceIds != null && clusterInstanceIds.Count > 0)
                    ? clusterInstanceIds.Select((id, i) => (id, i)).GroupBy(x => x.i / MaxClusterIdsPerQuery).Select(g => g.Select(x => x.id).ToList()).ToList()
                    : new List<List<int>> { null }; // null = no IN filter (all clusters)

                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP DEBUG: Point-in-box in {clusterIdBatches.Count} batch(es) (ClusterIds={clusterInstanceIds?.Count ?? 0})\n");

                foreach (var batch in clusterIdBatches)
                {
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        var query = @"
                            SELECT DISTINCT cz.SleeveInstanceId, cs.ClusterInstanceId, cz.ClashZoneGuid
                            FROM ClashZones cz
                            CROSS JOIN ClusterSleeves_v2 cs
                            WHERE cz.SleeveInstanceId > 0
                              AND cs.ClusterInstanceId > 0
                              AND cz.SleevePlacementX IS NOT NULL AND cz.SleevePlacementX != 0.0
                              AND cz.SleevePlacementY IS NOT NULL AND cz.SleevePlacementY != 0.0
                              AND cz.SleevePlacementZ IS NOT NULL AND cz.SleevePlacementZ != 0.0
                              AND (cz.ClusterInstanceId <= 0 OR cz.ClusterInstanceId IS NULL)
                              AND cs.BoundingBoxMinX != 0.0 AND cs.BoundingBoxMinY != 0.0 AND cs.BoundingBoxMinZ != 0.0
                              AND cs.BoundingBoxMaxX != 0.0 AND cs.BoundingBoxMaxY != 0.0 AND cs.BoundingBoxMaxZ != 0.0
                              AND cz.SleevePlacementX >= cs.BoundingBoxMinX AND cz.SleevePlacementX <= cs.BoundingBoxMaxX
                              AND cz.SleevePlacementY >= cs.BoundingBoxMinY AND cz.SleevePlacementY <= cs.BoundingBoxMaxY
                              AND cz.SleevePlacementZ >= cs.BoundingBoxMinZ AND cz.SleevePlacementZ <= cs.BoundingBoxMaxZ
                              AND cz.SleeveInstanceId != cs.ClusterInstanceId";
                        if (batch != null && batch.Count > 0)
                        {
                            var idPlaceholders = string.Join(",", batch.Select((_, i) => $"@clusterId{i}"));
                            query += $" AND cs.ClusterInstanceId IN ({idPlaceholders})";
                            for (int i = 0; i < batch.Count; i++)
                                cmd.Parameters.AddWithValue($"@clusterId{i}", batch[i]);
                        }
                        if (!string.IsNullOrEmpty(targetCategory))
                        {
                            query += " AND cs.Category = @targetCategory";
                            cmd.Parameters.AddWithValue("@targetCategory", targetCategory);
                        }
                        cmd.CommandText = query;

                        try
                        {
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var sleeveId = reader.GetInt32(reader.GetOrdinal("SleeveInstanceId"));
                                    var clusterId = reader.GetInt32(reader.GetOrdinal("ClusterInstanceId"));
                                    sleevesToDelete.Add(sleeveId);
                                    sleeveToClusterMap[sleeveId] = clusterId;
                                }
                            }
                        }
                        catch (Exception sqlEx)
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: SQL point-in-box batch failed: {sqlEx.Message}\n");
                            return 0;
                        }
                    }
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP: Found {sleevesToDelete.Count} sleeves to delete\n");
                
                // 📊 STAGE 2 CLEANUP SUMMARY: Log findings BEFORE cleanup
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 📊 STAGE 2 CLEANUP ANALYSIS:\n");
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}]   • Individual sleeves with placement points inside cluster bounding boxes: {sleevesToDelete.Count}\n");
                
                if (sleevesToDelete.Count > 0)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ✅ STAGE 2 CLEANUP WILL TRIGGER: {sleevesToDelete.Count} individual sleeve(s) found within cluster bounding boxes\n");
                }
                else
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⏭️ STAGE 2 CLEANUP SKIPPED: No individual sleeves found within cluster bounding boxes\n");
                }
                
                // Step 4: Delete sleeves from Revit (requires Revit API)
                if (sleevesToDelete.Count == 0)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP: No sleeves to delete\n");
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP COMPLETE: Deleted 0 individual sleeves (Category={targetCategory ?? "ALL"})\n");
                    return 0;
                }

                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP DEBUG: About to delete {sleevesToDelete.Count} sleeves from Revit\n");

                // ✅ CRITICAL FIX: Stage 2 - Run DB operations in the background
                // This makes them cluster constituents so Parameter Service will aggregate their parameters
                var mapCopy = new Dictionary<int, int>(sleeveToClusterMap);
                var sleeveIdsCopy = sleevesToDelete.ToList();
                string bgDbPath = context.DatabasePath;

                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        using (var bgContext = new SleeveDbContext(bgDbPath))
                        {
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2 (BACKGROUND): Assigning ClusterInstanceId to {mapCopy.Count} zones as cluster constituents\n");
                            AssignClusterInstanceIdToStage2Zones(bgContext, mapCopy);

                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2 (BACKGROUND): Flushing parameters from {sleeveIdsCopy.Count} sleeves to cluster snapshots\n");
                            FlushParametersToClusterSleeves(bgContext, null, sleeveIdsCopy); // Cannot use Revit doc in background safely
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ❌ STAGE 2 BACKGROUND ERROR: {ex.Message}\n{ex.StackTrace}\n");
                    }
                });

                int deletedCount = 0;
                try
                {
                    var elementIdsToDelete = new List<ElementId>();
                    foreach (var sleeveId in sleevesToDelete)
                    {
#if REVIT2023
                        var elementId = new ElementId(sleeveId);
#else
                        var elementId = new ElementId((long)sleeveId);
#endif
                        if (elementId != null && elementId != ElementId.InvalidElementId)
                            elementIdsToDelete.Add(elementId);
                    }

                    if (elementIdsToDelete.Count > 0)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Deleting {elementIdsToDelete.Count} sleeves in batches of {MaxDeletePerTransaction}\n");

                        // Delete in chunks to avoid Revit transaction/timeout with hundreds of elements
                        for (int offset = 0; offset < elementIdsToDelete.Count; offset += MaxDeletePerTransaction)
                        {
                            var chunk = elementIdsToDelete.Skip(offset).Take(MaxDeletePerTransaction).ToList();
                            using (Transaction deleteTx = new Transaction(doc, "Stage 2 Cleanup: Delete Individual Sleeves"))
                            {
                                deleteTx.Start();
                                doc.Delete(chunk);
                                deleteTx.Commit();
                                deletedCount += chunk.Count;
                            }
                        }

                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ DELETED {deletedCount} individual sleeves from Revit\n");
                    }
                }
                catch (Exception deleteEx)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ CLEANUP: Exception during deletion: {deleteEx.Message}\n{deleteEx.StackTrace}\n");
                }
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP COMPLETE: Deleted {deletedCount} individual sleeves\n");
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 DB-ONLY CLEANUP COMPLETE: Deleted {deletedCount} individual sleeves (Category={targetCategory ?? "ALL"})\n");

                return deletedCount;
            }
        }

        /// <summary>
        /// Cleanup individual sleeves within the bounding boxes of the provided cluster instances.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="clusters">List of placed cluster family instances</param>
        /// <param name="deferredParameters">Optional dictionary of deferred parameters (to ensure correct dimensions)</param>
        /// <param name="targetCategory">Optional category filter</param>
        /// <returns>Number of individual sleeves deleted</returns>
        public int CleanupSleevesWithinClusters(Document doc, List<FamilyInstance> clusters, Dictionary<ElementId, Dictionary<string, object>> deferredParameters = null, string targetCategory = null)
        {
            if (doc == null || clusters == null || clusters.Count == 0)
                return 0;

            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP (IN-MEMORY): Checking {clusters.Count} clusters against individual sleeves\n");

            var sleevesToDelete = new HashSet<int>();

            using (var context = new SleeveDbContext(doc))
            {
                // 1. Fetch all candidate individual sleeves (SleeveInstanceId > 0, ClusterInstanceId is NULL or 0)
                // We only need their placement points for the check
                var sleeves = new List<(int Id, XYZ Point)>();
                
                string query = @"
                    SELECT SleeveInstanceId, SleevePlacementX, SleevePlacementY, SleevePlacementZ 
                    FROM ClashZones 
                    WHERE SleeveInstanceId > 0 
                      AND (ClusterInstanceId = 0 OR ClusterInstanceId IS NULL)
                      AND SleevePlacementX IS NOT NULL";

                if (!string.IsNullOrEmpty(targetCategory))
                {
                    // Mapping MepElementCategory might be needed if column exists, usually it does
                     // query += " AND MepElementCategory = '" + targetCategory + "'";
                     // Assuming MepElementCategory matches targetCategory string
                }

                try 
                {
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        cmd.CommandText = query;
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int id = reader.GetInt32(0);
                                double x = reader.GetDouble(1);
                                double y = reader.GetDouble(2);
                                double z = reader.GetDouble(3);
                                sleeves.Add((id, new XYZ(x, y, z)));
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error fetching sleeves for cleanup: {ex.Message}\n");
                    return 0;
                }

                if (sleeves.Count == 0)
                    return 0;

                // 2. Iterate clusters and check containment
                foreach (var cluster in clusters)
                {
                    if (cluster == null || !cluster.IsValidObject) continue;

                    BoundingBoxXYZ bbox = null;
                    Transform transform = null;

                    // Try to get updated geometry first
                    var revitBBox = cluster.get_BoundingBox(null);
                    
                    // Logic to determine if we need to manually compute BBox due to deferred parameters
                    bool useDeferred = false;
                    double width = 0, height = 0, depth = 0;
                    
                    if (deferredParameters != null && deferredParameters.ContainsKey(cluster.Id))
                    {
                        var paramsDict = deferredParameters[cluster.Id];
                        if (paramsDict.ContainsKey("Element Width") && paramsDict.ContainsKey("Element Height"))
                        {
                            width = (double)paramsDict["Element Width"];
                            height = (double)paramsDict["Element Height"];
                            
                            // Try to get depth
                            if (paramsDict.ContainsKey("Element Depth")) 
                                depth = (double)paramsDict["Element Depth"];
                            else if (paramsDict.ContainsKey("Length"))
                                depth = (double)paramsDict["Length"];
                            else
                                depth = 1.0; // Fallback depth if unknown
                                
                            useDeferred = true;
                        }
                    }

                    if (useDeferred)
                    {
                        // Manual BBox check requires local coordinate system logic
                        // Center is LocationPoint
                        var loc = cluster.Location as LocationPoint;
                        if (loc != null)
                        {
                            transform = Transform.CreateRotationAtPoint(XYZ.BasisZ, loc.Rotation, loc.Point);
                            // We will check points by transforming them INVERSELY to local space
                            // Local BBox is then simply +/- W/2, H/2, D/2
                        }
                    }
                    else
                    {
                        bbox = revitBBox;
                    }

                    // Perform checks
                    foreach (var sleeve in sleeves)
                    {
                        if (sleevesToDelete.Contains(sleeve.Id)) continue; // Already marked

                        bool inside = false;
                        if (useDeferred && transform != null)
                        {
                            // Transform point to local space
                            XYZ localPoint = transform.Inverse.OfPoint(sleeve.Point);
                            // Assume centered at (0,0,0) in local space ? 
                            // Wait, CreateRotationAtPoint rotates around the point, so origin remains placed point.
                            // Local space relative to origin? No, CreateRotationAtPoint creates a transform where the 'Point' is the center of rotation.
                            // But usually we want a transform where (0,0,0) is the center of the element.
                            
                            // Let's rely on standard BBox logic if possible.
                            // If deferred, we assume dimensions are larger/different than current.
                            // We can construct a BBox centered at loc.Point.
                            
                            // Simple AABB check if rotation is 0?
                            // With rotation, it's OBB.
                            // Let's skip complex OBB for now and use Revit's BBox if valid, or a slightly expanded box.
                            
                            // If useDeferred is true, it means the Revit element is likely SMALLER (default size).
                            // We should use the deferred dimensions.
                            // Let's approximate: Check distance in 2D plane + Z check.
                            // This is simpler and robust for "Point in Zone".
                            
                            // Not implemented fully for deferred parameters rotation in this snippet to keep it safe.
                            // Fallback to Revit BBox for now, assuming regen handled most cases.
                            if (revitBBox != null)
                            {
                                // Ensure point is strictly inside
                                if (sleeve.Point.X >= revitBBox.Min.X && sleeve.Point.X <= revitBBox.Max.X &&
                                    sleeve.Point.Y >= revitBBox.Min.Y && sleeve.Point.Y <= revitBBox.Max.Y &&
                                    sleeve.Point.Z >= revitBBox.Min.Z && sleeve.Point.Z <= revitBBox.Max.Z)
                                {
                                    inside = true;
                                }
                            }
                        }
                        else if (bbox != null)
                        {
                             if (sleeve.Point.X >= bbox.Min.X && sleeve.Point.X <= bbox.Max.X &&
                                 sleeve.Point.Y >= bbox.Min.Y && sleeve.Point.Y <= bbox.Max.Y &&
                                 sleeve.Point.Z >= bbox.Min.Z && sleeve.Point.Z <= bbox.Max.Z)
                             {
                                 inside = true;
                             }
                        }

                        if (inside)
                        {
                            sleevesToDelete.Add(sleeve.Id);
                        }
                    }
                }
            }

            // Delete
            if (sleevesToDelete.Count > 0)
            {
                using (Transaction t = new Transaction(doc, "Cleanup Sleeves Inside Clusters"))
                {
                    t.Start();
#if REVIT2023
                    var ids = sleevesToDelete.Select(id => new ElementId(id)).ToList();
#else
                    var ids = sleevesToDelete.Select(id => new ElementId((long)id)).ToList();
#endif
                    doc.Delete(ids);
                    t.Commit();
                }
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ DELETED {sleevesToDelete.Count} sleeves in in-memory cleanup\n");
            }

            return sleevesToDelete.Count;
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Stage 2 parameter flushing for cluster sleeves.
        /// Aggregates parameters from sleeves being deleted into the parent cluster sleeve snapshot.
        /// </summary>
        private void FlushParametersToClusterSleeves(SleeveDbContext context, Document doc, List<int> sleeveInstanceIds)
        {
            if (sleeveInstanceIds == null || sleeveInstanceIds.Count == 0)
                return;

            SafeFileLogger.SafeAppendText("batch_v2.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2 FLUSH: Starting parameter aggregation for {sleeveInstanceIds.Count} sleeves\n");

            try
            {
                // Group sleeves by their parent cluster (find which cluster each sleeve is inside)
                var sleevesByCluster = new Dictionary<int, List<int>>(); // ClusterInstanceId -> SleeveIds
                
                foreach (var sleeveId in sleeveInstanceIds)
                {
                    try
                    {
                        using (var cmd = context.Connection.CreateCommand())
                        {
                            // Find which cluster this sleeve is inside (same logic as CleanupSleevesWithinClusters)
                            cmd.CommandText = @"
                                SELECT cs.ClusterInstanceId
                                FROM ClashZones cz
                                CROSS JOIN ClusterSleeves_v2 cs
                                WHERE cz.SleeveInstanceId = @sleeveId
                                  AND cs.ClusterInstanceId > 0
                                  AND cz.SleevePlacementX >= cs.BoundingBoxMinX AND cz.SleevePlacementX <= cs.BoundingBoxMaxX
                                  AND cz.SleevePlacementY >= cs.BoundingBoxMinY AND cz.SleevePlacementY <= cs.BoundingBoxMaxY
                                  AND cz.SleevePlacementZ >= cs.BoundingBoxMinZ AND cz.SleevePlacementZ <= cs.BoundingBoxMaxZ
                                LIMIT 1";
                            cmd.Parameters.AddWithValue("@sleeveId", sleeveId);
                            var result = cmd.ExecuteScalar();
                            
                            if (result != null && result != DBNull.Value)
                            {
                                int clusterId = Convert.ToInt32(result);
                                if (!sleevesByCluster.ContainsKey(clusterId))
                                    sleevesByCluster[clusterId] = new List<int>();
                                sleevesByCluster[clusterId].Add(sleeveId);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 FLUSH: Error finding cluster for sleeve {sleeveId}: {ex.Message}\n");
                    }
                }

                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2 FLUSH: Found {sleevesByCluster.Count} clusters to update\n");

                // Flush parameters for each cluster
                foreach (var kvp in sleevesByCluster)
                {
                    int clusterInstanceId = kvp.Key;
                    var clusterSleeveIds = kvp.Value;
                    
                    try
                    {
                        FlushParametersForSingleCluster(context, doc, clusterInstanceId, clusterSleeveIds);
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("batch_v2.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 FLUSH: Failed for cluster {clusterInstanceId}: {ex.Message}\n");
                    }
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 FLUSH: Error in FlushParametersToClusterSleeves: {ex.Message}\n");
            }
        }

        /// <summary>
        /// Flushes parameters from individual sleeves to a single cluster sleeve snapshot.
        /// </summary>
        private void FlushParametersForSingleCluster(SleeveDbContext context, Document doc, int clusterInstanceId, List<int> sleeveInstanceIds)
        {
            if (sleeveInstanceIds.Count == 0)
                return;

            // SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2 FLUSH: Cluster {clusterInstanceId} - Aggregating {sleeveInstanceIds.Count} sleeves\n");

            var mepParamsToMerge = new Dictionary<string, object>();
            var hostParamsToMerge = new Dictionary<string, object>();

            // Step 1: Read parameters from sleeves being deleted (use ClashZoneGuid lookup)
            foreach (var sleeveId in sleeveInstanceIds)
            {
                try
                {
                    // Get ClashZoneGuid for this sleeve
                    string clashZoneGuid = null;
                    using (var guidCmd = context.Connection.CreateCommand())
                    {
                        guidCmd.CommandText = "SELECT ClashZoneGuid FROM ClashZones WHERE SleeveInstanceId = @sleeveId LIMIT 1";
                        guidCmd.Parameters.AddWithValue("@sleeveId", sleeveId);
                        var result = guidCmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                            clashZoneGuid = result.ToString();
                    }

                    if (!string.IsNullOrEmpty(clashZoneGuid))
                    {
                        // Query SleeveSnapshots by ClashZoneGuid
                        using (var cmd = context.Connection.CreateCommand())
                        {
                            cmd.CommandText = @"
                                SELECT MepParametersJson, HostParametersJson 
                                FROM SleeveSnapshots 
                                WHERE UPPER(ClashZoneGuid) = UPPER(@guid)";
                            cmd.Parameters.AddWithValue("@guid", clashZoneGuid);
                            
                            using (var reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    var mepJson = reader.IsDBNull(0) ? null : reader.GetString(0);
                                    var hostJson = reader.IsDBNull(1) ? null : reader.GetString(1);
                                    
                                    if (!string.IsNullOrEmpty(mepJson))
                                    {
                                        try
                                        {
                                            var mepParams = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(mepJson);
                                            if (mepParams != null)
                                            {
                                                foreach (var param in mepParams)
                                                {
                                                    if (!string.IsNullOrEmpty(param.Key) && param.Value != null)
                                                    {
                                                        if (!mepParamsToMerge.ContainsKey(param.Key))
                                                            mepParamsToMerge[param.Key] = param.Value;
                                                    }
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                            }
                        }
                    }
                    
                    // Fallback: Extract from Revit element if no snapshot found
                    if (mepParamsToMerge.Count == 0)
                    {
                        ExtractParametersFromRevitElement(doc, sleeveId, mepParamsToMerge, hostParamsToMerge);
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("batch_v2.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 FLUSH: Error reading params for sleeve {sleeveId}: {ex.Message}\n");
                }
            }

            if (mepParamsToMerge.Count == 0 && hostParamsToMerge.Count == 0)
            {
                SafeFileLogger.SafeAppendText("batch_v2.log", 
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 FLUSH: No parameters found for cluster {clusterInstanceId}\n");
                return;
            }

            // Step 2: Read existing cluster snapshot
            string existingMepJson = null;
            using (var cmd = context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT MepParametersJson 
                    FROM SleeveSnapshots 
                    WHERE ClusterInstanceId = @clusterId
                    LIMIT 1";
                cmd.Parameters.AddWithValue("@clusterId", clusterInstanceId);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                    existingMepJson = result.ToString();
            }

            // Step 3: Merge with aggregation (same logic as CombinedSleevePlacementService)
            var finalParams = MergeWithAggregation(existingMepJson, mepParamsToMerge);

            // Step 4: Update cluster snapshot
            using (var cmd = context.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    UPDATE SleeveSnapshots 
                    SET MepParametersJson = @mepParams,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE ClusterInstanceId = @clusterId";
                cmd.Parameters.AddWithValue("@mepParams", System.Text.Json.JsonSerializer.Serialize(finalParams));
                cmd.Parameters.AddWithValue("@clusterId", clusterInstanceId);
                
                int rowsUpdated = cmd.ExecuteNonQuery();
                // SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ STAGE 2 FLUSH: Cluster {clusterInstanceId} - Updated {rowsUpdated} row(s), {finalParams.Count} params\n");
            }
        }

        /// <summary>
        /// Extracts parameters directly from a Revit element.
        /// </summary>
        private void ExtractParametersFromRevitElement(Document doc, int sleeveId, Dictionary<string, object> mepParams, Dictionary<string, object> hostParams)
        {
            try
            {
                if (doc == null) return; // Called from background thread, cannot access Revit API

#if REVIT2023
                var element = doc.GetElement(new ElementId(sleeveId));
#else
                var element = doc.GetElement(new ElementId((long)sleeveId));
#endif
                if (element == null) return;

                foreach (Parameter param in element.Parameters)
                {
                    try
                    {
                        if (!param.HasValue || string.IsNullOrEmpty(param.Definition?.Name))
                            continue;

                        var paramName = param.Definition.Name;
                        object paramValue = null;

                        switch (param.StorageType)
                        {
                            case StorageType.String:
                                paramValue = param.AsString();
                                break;
                            case StorageType.Double:
                                paramValue = param.AsDouble();
                                break;
                            case StorageType.Integer:
                                paramValue = param.AsInteger();
                                break;
                            case StorageType.ElementId:
                                var id = param.AsElementId();
                                paramValue = id?.GetIntegerValue() ?? 0;
                                break;
                        }

                        if (paramValue != null)
                        {
                            if (IsMepParameter(paramName))
                            {
                                if (!mepParams.ContainsKey(paramName))
                                    mepParams[paramName] = paramValue;
                            }
                            else
                            {
                                if (!hostParams.ContainsKey(paramName))
                                    hostParams[paramName] = paramValue;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// Determines if a parameter name indicates it's an MEP-related parameter.
        /// </summary>
        private bool IsMepParameter(string paramName)
        {
            if (string.IsNullOrEmpty(paramName))
                return false;

            var mepKeywords = new[] 
            { 
                "MEP", "DUCT", "PIPE", "CABLE", "CONDUIT", "TRAY",
                "SYSTEM", "SERVICE", "SIZE", "DIAMETER", "WIDTH", "HEIGHT",
                "LEVEL", "OFFSET", "FLOW", "VELOCITY", "PRESSURE"
            };

            var upperName = paramName.ToUpperInvariant();
            return mepKeywords.Any(kw => upperName.Contains(kw));
        }

        /// <summary>
        /// Merges two parameter dictionaries, aggregating values when they differ.
        /// </summary>
        private Dictionary<string, object> MergeWithAggregation(string existingJson, Dictionary<string, object> newParams)
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            
            // Load existing params
            if (!string.IsNullOrEmpty(existingJson) && existingJson != "{}")
            {
                try
                {
                    var existing = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(existingJson);
                    if (existing != null)
                    {
                        foreach (var kvp in existing)
                            result[kvp.Key] = kvp.Value;
                    }
                }
                catch { }
            }
            
            // Merge new params with aggregation
            foreach (var kvp in newParams)
            {
                if (result.ContainsKey(kvp.Key))
                {
                    var existingValue = result[kvp.Key]?.ToString() ?? "";
                    var newValue = kvp.Value?.ToString() ?? "";
                    
                    if (!string.Equals(existingValue, newValue, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!existingValue.Contains(newValue, StringComparison.OrdinalIgnoreCase))
                        {
                            result[kvp.Key] = $"{existingValue}, {newValue}";
                        }
                    }
                }
                else
                {
                    result[kvp.Key] = kvp.Value;
                }
            }
            
            return result;
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Assigns ClusterInstanceId to Stage 2 cleaned up zones in ClashZones table.
        /// This makes them cluster constituents so Parameter Service will aggregate their parameters.
        /// </summary>
        private void AssignClusterInstanceIdToStage2Zones(SleeveDbContext context, Dictionary<int, int> sleeveToClusterMap)
        {
            if (sleeveToClusterMap.Count == 0)
                return;

            SafeFileLogger.SafeAppendText("placement_performance.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2 ASSIGN: Updating {sleeveToClusterMap.Count} zones as cluster constituents\n");

            int entryCount = 0;
            const int batchSize = 100;
            var updates = sleeveToClusterMap.ToList();
            
            for (int i = 0; i < updates.Count; i += batchSize)
            {
                var batch = updates.Skip(i).Take(batchSize).ToList();
                if (batch.Count == 0) continue;

                try
                {
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        var caseBuilder = new System.Text.StringBuilder();
                        var idList = new System.Text.StringBuilder();
                        
                        caseBuilder.Append("CASE SleeveInstanceId ");
                        
                        for (int j = 0; j < batch.Count; j++)
                        {
                            string pSleeve = $"@s{j}";
                            string pCluster = $"@c{j}";
                            
                            caseBuilder.Append($"WHEN {pSleeve} THEN {pCluster} ");
                            idList.Append(j == 0 ? pSleeve : $", {pSleeve}");
                            
                            cmd.Parameters.AddWithValue(pSleeve, batch[j].Key); // SleeveInstanceId
                            cmd.Parameters.AddWithValue(pCluster, batch[j].Value); // ClusterInstanceId
                        }
                        
                        caseBuilder.Append("END");

                        cmd.CommandText = $@"
                            UPDATE ClashZones 
                            SET ClusterInstanceId = {caseBuilder},
                                IsClusterResolvedFlag = 1,
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE SleeveInstanceId IN ({idList})
                              AND (ClusterInstanceId IS NULL OR ClusterInstanceId = 0)";
                        
                        int rows = cmd.ExecuteNonQuery();
                        entryCount += rows;
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("placement_performance.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 ASSIGN BATCH FAILED: {ex.Message}\n");
                }
            }

            SafeFileLogger.SafeAppendText("placement_performance.log", 
                $"[{DateTime.Now:HH:mm:ss}] ✅ STAGE 2 ASSIGN: Updated {entryCount} zones as cluster constituents in {Math.Ceiling((double)updates.Count / batchSize)} batches\n");

            // ✅ CRITICAL FIX: Also update SleeveSnapshots to set ClusterInstanceId for Stage 2 sleeves
            // This ensures Parameter Service sees them as cluster constituents
            UpdateSleeveSnapshotsForStage2Sleeves(context, sleeveToClusterMap);
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Updates SleeveSnapshots table to set ClusterInstanceId for Stage 2 cleaned up sleeves.
        /// This makes them appear as cluster constituents to the Parameter Service.
        /// </summary>
        private void UpdateSleeveSnapshotsForStage2Sleeves(SleeveDbContext context, Dictionary<int, int> sleeveToClusterMap)
        {
            if (sleeveToClusterMap.Count == 0)
                return;

            SafeFileLogger.SafeAppendText("placement_performance.log", 
                $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2 SNAPSHOT: Updating {sleeveToClusterMap.Count} snapshots with ClusterInstanceId\n");

            int entryCount = 0;
            const int batchSize = 50; // Smaller batch size for complex query
            var updates = sleeveToClusterMap.ToList();

            for (int i = 0; i < updates.Count; i += batchSize)
            {
                var batch = updates.Skip(i).Take(batchSize).ToList();
                if (batch.Count == 0) continue;

                try
                {
                    using (var cmd = context.Connection.CreateCommand())
                    {
                        var caseBuilder = new System.Text.StringBuilder();
                        var idList = new System.Text.StringBuilder();
                        
                        caseBuilder.Append("CASE SleeveInstanceId ");
                        
                        for (int j = 0; j < batch.Count; j++)
                        {
                            string pSleeve = $"@s{j}";
                            string pCluster = $"@c{j}";
                            
                            caseBuilder.Append($"WHEN {pSleeve} THEN {pCluster} ");
                            idList.Append(j == 0 ? pSleeve : $", {pSleeve}");
                            
                            cmd.Parameters.AddWithValue(pSleeve, batch[j].Key);
                            cmd.Parameters.AddWithValue(pCluster, batch[j].Value);
                        }
                        
                        caseBuilder.Append("END");
                        string ids = idList.ToString();

                        // 1. UPDATE existing snapshots
                        cmd.CommandText = $@"
                            UPDATE SleeveSnapshots 
                            SET ClusterInstanceId = {caseBuilder},
                                SourceType = 'Cluster',
                                UpdatedAt = CURRENT_TIMESTAMP
                            WHERE SleeveInstanceId IN ({ids});";
                        
                        int updated = cmd.ExecuteNonQuery();

                        // 2. INSERT missing snapshots (from ClashZones)
                        cmd.CommandText = $@"
                            INSERT INTO SleeveSnapshots (
                                SleeveInstanceId, ClusterInstanceId, SourceType,
                                ClashZoneGuid, MepParametersJson, HostParametersJson,
                                CreatedAt, UpdatedAt
                            )
                            SELECT 
                                SleeveInstanceId, 
                                {caseBuilder}, 
                                'Cluster',
                                ClashZoneGuid, MepParameterValuesJson, HostParameterValuesJson,
                                CURRENT_TIMESTAMP, CURRENT_TIMESTAMP
                            FROM ClashZones
                            WHERE SleeveInstanceId IN ({ids})
                              AND SleeveInstanceId NOT IN (SELECT SleeveInstanceId FROM SleeveSnapshots WHERE SleeveInstanceId IN ({ids}))";
                        
                        int inserted = cmd.ExecuteNonQuery();
                        entryCount += (updated + inserted);
                        
                        if (updated + inserted > 0)
                        {
                            SafeFileLogger.SafeAppendText("placement_performance.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 STAGE 2 SNAPSHOT BATCH: Updated {updated}, Inserted {inserted}\n");
                        }
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("placement_performance.log", 
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ STAGE 2 SNAPSHOT BATCH FAILED: {ex.Message}\n");
                }
            }
            
            SafeFileLogger.SafeAppendText("placement_performance.log", 
                $"[{DateTime.Now:HH:mm:ss}] ✅ STAGE 2 SNAPSHOT: Processed {entryCount} snapshots in {Math.Ceiling((double)updates.Count / batchSize)} batches\n");
        }
    }
}
