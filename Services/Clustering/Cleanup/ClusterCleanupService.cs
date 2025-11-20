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
                if (placedClusters == null || placedClusters.Count == 0) return 0;

                // Build protection set of cluster sleeve IDs
                var clusterSleeveIds = new HashSet<int>(placedClusters.Where(c => c != null).Select(c => c.Id.IntegerValue));

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

                var individualSleeves = new List<FamilyInstance>();
                foreach (var s in allSleeves)
                {
                    int id = s.Id.IntegerValue;
                    if (clusterSleeveIds.Contains(id)) continue; // protected
                    var clusterParam = s.LookupParameter("Cluster Sleeve Instance ID");
                    var sleeveInstanceParam = s.LookupParameter("Sleeve Instance ID");
                    int clusterValue = clusterParam?.AsInteger() ?? -1;
                    int sleeveInstanceValue = sleeveInstanceParam?.AsInteger() ?? -999;
                    bool isClusterSleeve = (sleeveInstanceValue == -1) || (clusterValue > 0 && clusterValue == id);
                    if (isClusterSleeve) continue;
                    individualSleeves.Add(s);
                }

                if (individualSleeves.Count == 0) return 0;

                // Build cluster sleeve bounding boxes (axis-aligned)
                var clusterBboxes = placedClusters
                    .Select(c => c?.get_BoundingBox(null))
                    .Where(b => b != null && b.Enabled)
                    .ToList();

                if (clusterBboxes.Count == 0) return 0;

                var toDelete = new List<ElementId>();
                foreach (var individual in individualSleeves)
                {
                    var ibbox = individual.get_BoundingBox(null);
                    if (ibbox == null || !ibbox.Enabled) continue;
                    foreach (var cbbox in clusterBboxes)
                    {
                        if (cbbox == null) continue;
                        bool inside = ibbox.Min.X >= cbbox.Min.X && ibbox.Max.X <= cbbox.Max.X &&
                                      ibbox.Min.Y >= cbbox.Min.Y && ibbox.Max.Y <= cbbox.Max.Y &&
                                      ibbox.Min.Z >= cbbox.Min.Z && ibbox.Max.Z <= cbbox.Max.Z;
                        if (inside)
                        {
                            toDelete.Add(individual.Id);
                            break;
                        }
                    }
                }

                if (toDelete.Count == 0) return 0;

                using (var tx = new Transaction(doc, "Delete Individual Sleeves Within Clusters"))
                {
                    tx.Start();
                    foreach (var id in toDelete)
                    {
                        try
                        {
                            doc.Delete(id);
                            deletedCount++;
                        }
                        catch (Exception ex)
                        {
                            DebugLogger.Warning($"[ClusterCleanupService] Failed deleting sleeve {id.IntegerValue}: {ex.Message}");
                        }
                    }
                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
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
