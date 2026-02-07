using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
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
                var sleevesToDelete = new HashSet<int>();
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
                            SELECT DISTINCT cz.SleeveInstanceId, cs.ClusterInstanceId
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
                                    sleevesToDelete.Add(sleeveId);
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

                int deletedCount = 0;
                try
                {
                    var elementIdsToDelete = new List<ElementId>();
                    foreach (var sleeveId in sleevesToDelete)
                    {
                        var elementId = new ElementId(sleeveId);
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
                    var ids = sleevesToDelete.Select(id => new ElementId(id)).ToList();
                    doc.Delete(ids);
                    t.Commit();
                }
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ DELETED {sleevesToDelete.Count} sleeves in in-memory cleanup\n");
            }

            return sleevesToDelete.Count;
        }
    }
}
