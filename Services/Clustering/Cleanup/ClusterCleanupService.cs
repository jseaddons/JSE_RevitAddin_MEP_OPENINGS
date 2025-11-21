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
                    individualSleeves.Add(s);
                }

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Found {individualSleeves.Count} individual sleeves to check\n");

                if (individualSleeves.Count == 0) return 0;

                // Build cluster sleeve bounding boxes (axis-aligned)
                var clusterBboxes = placedClusters
                    .Select(c => c?.get_BoundingBox(null))
                    .Where(b => b != null && b.Enabled)
                    .ToList();

                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                    $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Built {clusterBboxes.Count} cluster bounding boxes\n");

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
                    
                    var ibbox = individual.get_BoundingBox(null);
                    if (ibbox == null || !ibbox.Enabled) continue;
                    
                    foreach (var cbbox in clusterBboxes)
                    {
                        if (cbbox == null) continue;
                        
                        // ✅ EDGE CASE: Use more lenient containment check (allows partial overlap)
                        // Original: strict containment (all corners inside)
                        // New: center point inside OR significant overlap
                        XYZ individualCenter = (ibbox.Min + ibbox.Max) / 2.0;
                        bool centerInside = individualCenter.X >= cbbox.Min.X && individualCenter.X <= cbbox.Max.X &&
                                           individualCenter.Y >= cbbox.Min.Y && individualCenter.Y <= cbbox.Max.Y &&
                                           individualCenter.Z >= cbbox.Min.Z && individualCenter.Z <= cbbox.Max.Z;
                        
                        // Also check if there's significant overlap (at least 50% of individual sleeve volume)
                        bool hasSignificantOverlap = ibbox.Min.X < cbbox.Max.X && ibbox.Max.X > cbbox.Min.X &&
                                                    ibbox.Min.Y < cbbox.Max.Y && ibbox.Max.Y > cbbox.Min.Y &&
                                                    ibbox.Min.Z < cbbox.Max.Z && ibbox.Max.Z > cbbox.Min.Z;
                        
                        // Calculate overlap volume
                        double overlapVolume = 0.0;
                        if (hasSignificantOverlap)
                        {
                            double overlapX = Math.Max(0, Math.Min(ibbox.Max.X, cbbox.Max.X) - Math.Max(ibbox.Min.X, cbbox.Min.X));
                            double overlapY = Math.Max(0, Math.Min(ibbox.Max.Y, cbbox.Max.Y) - Math.Max(ibbox.Min.Y, cbbox.Min.Y));
                            double overlapZ = Math.Max(0, Math.Min(ibbox.Max.Z, cbbox.Max.Z) - Math.Max(ibbox.Min.Z, cbbox.Min.Z));
                            overlapVolume = overlapX * overlapY * overlapZ;
                            
                            double individualVolume = (ibbox.Max.X - ibbox.Min.X) * (ibbox.Max.Y - ibbox.Min.Y) * (ibbox.Max.Z - ibbox.Min.Z);
                            double overlapRatio = individualVolume > 0 ? overlapVolume / individualVolume : 0.0;
                            hasSignificantOverlap = overlapRatio >= 0.5; // At least 50% overlap
                        }
                        
                        if (centerInside || hasSignificantOverlap)
                        {
                            // ✅ CRITICAL SAFETY CHECK: Triple-check this is NOT a cluster sleeve before marking for deletion
                            if (clusterSleeveIds.Contains(individualId))
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️⚠️⚠️ CRITICAL SAFETY: Sleeve {individualId} matched containment check but is a cluster sleeve - NOT MARKING FOR DELETION\n");
                                break;
                            }
                            
                            SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 CLEANUP: Individual sleeve {individualId} is within cluster bbox (centerInside={centerInside}, hasOverlap={hasSignificantOverlap}) - MARKING FOR DELETION\n");
                            toDelete.Add(individual.Id);
                            break;
                        }
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
