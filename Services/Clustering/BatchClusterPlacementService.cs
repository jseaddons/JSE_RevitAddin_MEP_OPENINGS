using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Cleanup;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering
{
    /// <summary>
    /// Service for Phase 2: Sequential Placement from Database - V2 Architecture
    /// Responsibilities:
    /// 1. Query 'Pending' clusters from ClusterSleeves_v2
    /// 2. Place Family Instances (Sequential, Single Thread)
    /// 3. Update DB (Status='Placed', ClusterInstanceId)
    /// </summary>
    public class BatchClusterPlacementService
    {
        private readonly string _databasePath;
        private readonly IClashZoneRepository _repository;
        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor _performanceMonitor;
        private readonly JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveParameterService _parameterService;
        private readonly IClusterCleanupService _cleanupService;

        public BatchClusterPlacementService(
            string databasePath, 
            IClashZoneRepository repository,
            JSE_RevitAddin_MEP_OPENINGS.Services.Placement.SleeveParameterService parameterService,
            IClusterCleanupService cleanupService = null,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IPerformanceMonitor performanceMonitor = null)
        {
            _databasePath = databasePath;
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            _parameterService = parameterService ?? throw new ArgumentNullException(nameof(parameterService));
            _cleanupService = cleanupService ?? new ClusterCleanupService(); // Default if not injected
            _performanceMonitor = performanceMonitor;
        }

        public (int placed, int failed) PlaceFromDatabase(Document doc, string batchId, bool useSingleTransaction = true)
        {
            using (var tracker = _performanceMonitor?.TrackOperation("Step 8: BULK PLACEMENT - CLUSTERS"))
            {
                int placedCount = 0;
                int failedCount = 0;
                var placedInstances = new List<FamilyInstance>(); // Track placed instances for 2nd cleanup

                // 1. Query Pending Clusters
                List<BatchClusterData> pendingClusters;
                using (_performanceMonitor?.TrackOperation("Query Pending Clusters"))
                {
                    pendingClusters = GetPendingClusters(batchId);
                }
                
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🏗️ STARTING PLACEMENT V2: Batch {batchId}, Pending={pendingClusters.Count}, Mode={(useSingleTransaction ? "BULK" : "SEQUENTIAL")}\n");

                if (pendingClusters.Count == 0) return (0, 0);

                // ✅ OPTIMIZATION: Pre-fetch all required Family Symbols
                var symbolCache = new Dictionary<string, FamilySymbol>();
                using (_performanceMonitor?.TrackOperation("Cache Family Symbols"))
                {
                    var uniqueFamilyNames = pendingClusters.Select(c => c.FamilyName).Distinct().ToList();
                    var allSymbols = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .ToList();

                    foreach (var name in uniqueFamilyNames)
                    {
                        var symbol = allSymbols.FirstOrDefault(s => s.Name == name || s.Family.Name == name);
                        if (symbol != null) 
                        {
                            symbolCache[name] = symbol;
                            // Ensure active
                            if (!symbol.IsActive) symbol.Activate();
                        }
                    }
                }

                // ✅ OPTIMIZATION: Pre-fetch all constituent ClashZones to avoid per-cluster DB queries
                var zoneCache = new Dictionary<Guid, ClashZone>();
                using (_performanceMonitor?.TrackOperation("Pre-fetch Constituent Zones"))
                {
                    var allConstituentGuids = pendingClusters
                        .SelectMany(c => (c.ConstituentZoneGuids ?? "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                        .Select(s => Guid.TryParse(s.Trim(), out var g) ? g : Guid.Empty)
                        .Where(g => g != Guid.Empty)
                        .Distinct()
                        .ToList();

                    if (allConstituentGuids.Count > 0)
                    {
                        var fetchedZones = _repository.GetClashZonesByGuids(allConstituentGuids);
                        foreach (var z in fetchedZones) zoneCache[z.Id] = z;
                    }
                }

                // Check for existing transaction (Revit specific)
                bool isNestedTransaction = doc.IsModifiable;

                // ✅ REF: Aggregated Logging Context
                using (var placementTracker = _performanceMonitor?.TrackOperation("Place Cluster Instances"))
                {
                    // 🚀 NEW: BULK PLACEMENT PATH (NewFamilyInstances2)
                    if (OptimizationFlags.UseBulkClusterSleevePlacement && (useSingleTransaction || isNestedTransaction))
                    {
                         try
                         {
                             // We must ensure we are in a transaction
                             if (isNestedTransaction)
                             {
                                 var result = PlaceBulkClusters(doc, pendingClusters, symbolCache, zoneCache, placementTracker, placedInstances);
                                 placedCount = result.placed;
                                 failedCount = result.failed;
                             }
                             else
                             {
                                 using (Transaction t = new Transaction(doc, "Place Batch Clusters V2 (Bulk Optimized)"))
                                 {
                                     t.Start();
                                     var result = PlaceBulkClusters(doc, pendingClusters, symbolCache, zoneCache, placementTracker, placedInstances);
                                     placedCount = result.placed;
                                     failedCount = result.failed;
                                     t.Commit();
                                 }
                             }
                         }
                         catch(Exception ex)
                         {
                             SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ BULK PLACEMENT ERROR: {ex.Message} - Falling back to sequential\n");
                         }
                    }
                    else
                    {
                        // LEGACY: SEQUENTIAL OR INDIVIDUAL TRANSACTION PATH
                        if (useSingleTransaction || isNestedTransaction)
                        {
                            if (isNestedTransaction)
                            {
                                // MODE A-1: BULK (Existing Transaction)
                                foreach (var cluster in pendingClusters)
                                {
                                    if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances)) placedCount++;
                                    else failedCount++;
                                }
                            }
                            else
                            {
                                // MODE A-2: BULK (New Transaction)
                                using (Transaction t = new Transaction(doc, "Place Batch Clusters V2 (Bulk)"))
                                {
                                    t.Start();
                                    foreach (var cluster in pendingClusters)
                                    {
                                        if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances)) placedCount++;
                                        else failedCount++;
                                    }
                                    t.Commit();
                                }
                            }
                        }
                        else
                        {
                            // MODE B: SEQUENTIAL (Transaction per Cluster)
                            foreach (var cluster in pendingClusters)
                            {
                                using (Transaction t = new Transaction(doc, $"Place Cluster {cluster.ClusterGUID}"))
                                {
                                    try
                                    {
                                        t.Start();
                                        if (PlaceSingleCluster(doc, cluster, symbolCache, zoneCache, placementTracker, placedInstances))
                                        {
                                            placedCount++;
                                            t.Commit();
                                        }
                                        else
                                        {
                                            failedCount++;
                                            t.RollBack();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ TRANS ERROR: {ex.Message}\n");
                                        failedCount++;
                                    }
                                }
                            }
                        }
                    }
                    placementTracker?.SetItemCount(placedCount);
                }

                // ✅ CRITICAL FIX: Flush batched parameters
                if (OptimizationFlags.UseBatchedParameterWrites && placedCount > 0)
                {
                    using (var flushTracker = _performanceMonitor?.TrackOperation("Set Cluster Parameters (Flush)"))
                    {
                        var clusterFlushTimer = System.Diagnostics.Stopwatch.StartNew();
                        int flushedCount = 0;
                        
                        try
                        {
                            flushedCount = _parameterService.FlushDeferredParameters(clearList: true);
                            doc.Regenerate();
                            clusterFlushTimer.Stop();

                            SafeFileLogger.SafeAppendText("performance.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [CLUSTER-FLUSH] Clusters={placedCount}, Parameters={flushedCount}, Time={clusterFlushTimer.ElapsedMilliseconds}ms\n");
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ FLUSH FAILED: {ex.Message}\n");
                        }
                        
                        flushTracker?.SetItemCount(flushedCount);
                    }
                }
                
                // ✅ 2ND CLEANUP PASS: Delete individual sleeves within placed cluster bounding boxes
                if (_cleanupService != null && placedInstances.Count > 0)
                {
                   using (var cleanupTracker = _performanceMonitor?.TrackOperation("Cleanup Individual Sleeves (2nd Pass)"))
                   {
                        try
                        {
                            string targetCategory = null;
                            var batchParts = batchId.Split('_');
                            if (batchParts.Length >= 3)
                            {
                                targetCategory = batchParts[2];
                            }
                            
                            int deletedCount = _cleanupService.CleanupSleevesWithinClusters(
                                doc, 
                                placedInstances, 
                                deferredParameters: null, 
                                targetCategory: targetCategory);
                            
                            SafeFileLogger.SafeAppendText("batch_v2.log", 
                                $"[{DateTime.Now:HH:mm:ss}] 🧹 2ND CLEANUP: Deleted {deletedCount} individual sleeves overlapping with {placedInstances.Count} cluster sleeves (Category={targetCategory ?? "ALL"})\n");
                            
                            cleanupTracker?.SetItemCount(deletedCount);
                        }
                        catch (Exception ex)
                        {
                            SafeFileLogger.SafeAppendText("placement_errors.log", 
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ 2ND CLEANUP FAILED: {ex.Message}\n");
                        }
                   }
                }
                tracker?.SetItemCount(placedCount);
                SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] ✅ BATCH COMPLETED: Placed={placedCount}, Failed={failedCount}\n");
                return (placedCount, failedCount);
            }
        }

        private (int placed, int failed) PlaceBulkClusters(
            Document doc,
            List<BatchClusterData> clusters,
            Dictionary<string, FamilySymbol> symbolCache,
            Dictionary<Guid, ClashZone> zoneCache,
            JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker tracker,
            List<FamilyInstance> placedInstances)
        {
            var creationDataList = new List<Autodesk.Revit.Creation.FamilyInstanceCreationData>();
            var validClusters = new List<BatchClusterData>();

            foreach(var cluster in clusters)
            {
                 FamilySymbol symbol = null;
                 if (symbolCache != null && symbolCache.TryGetValue(cluster.FamilyName, out var cachedSymbol)) symbol = cachedSymbol;
                 else
                 {
                     symbol = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(x => x.Name == cluster.FamilyName || x.Family.Name == cluster.FamilyName);
                 }
                 
                 if (symbol == null) 
                 {
                    UpdateStatus(cluster.ClusterGUID, "Failed", "Missing Symbol");
                    continue;
                 }
                 if (!symbol.IsActive) symbol.Activate();

                 Element host = null;
                 Level level = null;
                 if (cluster.HostElementId > 0)
                 {
                    try { host = doc.GetElement(new ElementId((int)cluster.HostElementId)); } catch { }
                    if (host != null) level = doc.GetElement(host.LevelId) as Level;
                 }
                 if (level == null) level = new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstOrDefault() as Level;

                 XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                 var structuralType = Autodesk.Revit.DB.Structure.StructuralType.NonStructural;

                 Autodesk.Revit.Creation.FamilyInstanceCreationData data;
                 if (host != null)
                    data = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, host, level, structuralType);
                 else if (level != null)
                    data = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, level, structuralType);
                 else
                    data = new Autodesk.Revit.Creation.FamilyInstanceCreationData(location, symbol, structuralType);
                 
                 creationDataList.Add(data);
                 validClusters.Add(cluster);
            }

            if (creationDataList.Count == 0) return (0, clusters.Count);

            ICollection<ElementId> createdIds;
            try 
            {
                createdIds = doc.Create.NewFamilyInstances2(creationDataList);
            }
            catch(Exception ex)
            {
                SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ NewFamilyInstances2 FAILED: {ex.Message}\n");
                return (0, clusters.Count);
            }

            int placedCount = 0;
            int i = 0;

            // Collections for bulk DB updates
            var statusUpdates = new List<(string Guid, string Status, string Msg, int InstanceId)>();
            var calculatedColumnUpdates = new List<(List<Guid> Guids, BatchClusterData Cluster, int InstanceId)>();
            var legacySleeveInserts = new List<(BatchClusterData Cluster, int InstanceId)>();
            var oldClusterIdsToDelete = new List<int>();
            var allToDeleteRevitIds = new List<ElementId>();
            var allRepositoryFlagUpdates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsClusteredFlag, bool MarkedForClusterProcess, int AfterClusterSleeveId)>();

            foreach(var id in createdIds)
            {
                var cluster = validClusters[i];
                i++;
                
                try 
                {
                    var instance = doc.GetElement(id) as FamilyInstance;
                    if (instance != null)
                    {
                        placedInstances.Add(instance);
                        int clusterInstanceId = instance.Id.IntegerValue;

                        statusUpdates.Add((cluster.ClusterGUID, "Placed", null, clusterInstanceId));

                        var zoneGuids = new List<Guid>();
                        if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                        {
                            foreach (var guidStr in cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                if (Guid.TryParse(guidStr.Trim(), out Guid parsedGuid)) zoneGuids.Add(parsedGuid);
                            }
                        }

                        if (zoneGuids.Any())
                        {
                            calculatedColumnUpdates.Add((zoneGuids, cluster, clusterInstanceId));
                            legacySleeveInserts.Add((cluster, clusterInstanceId));

                            var zones = _repository.GetClashZonesByGuids(zoneGuids);
                            var oldIds = zones
                                .Where(z => z.ClusterInstanceId > 0 && z.ClusterInstanceId != clusterInstanceId)
                                .Select(z => z.ClusterInstanceId)
                                .Distinct();
                            oldClusterIdsToDelete.AddRange(oldIds);

                            var revitIdsToDelete = zones
                                .Where(z => z.SleeveInstanceId > 0)
                                .Select(z => new ElementId(z.SleeveInstanceId));
                            allToDeleteRevitIds.AddRange(revitIdsToDelete);

                            foreach (var z in zones)
                            {
                                allRepositoryFlagUpdates.Add((z.Id, true, true, false, -1, clusterInstanceId, true, true, clusterInstanceId));
                            }
                        }

                        bool isCircular = cluster.FamilyName.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0 || cluster.FamilyName.IndexOf("Circular", StringComparison.OrdinalIgnoreCase) >= 0;
                        
                        ClashZone templateZone = null;
                        if (zoneGuids.Any())
                        {
                            if (zoneCache != null && zoneCache.TryGetValue(zoneGuids[0], out var cached)) templateZone = cached;
                            else templateZone = _repository.GetClashZonesByGuids(new List<Guid> { zoneGuids[0] }).FirstOrDefault();
                        }
                        if (templateZone == null) templateZone = new ClashZone();
                        
                        templateZone.CalculatedSleeveWidth = cluster.ClusterWidth;
                        templateZone.CalculatedSleeveHeight = cluster.ClusterHeight;
                        templateZone.CalculatedSleeveDepth = cluster.ClusterDepth;
                        templateZone.CalculatedRotation = cluster.RotationAngleRad;

                        _parameterService.SetSleeveParametersImmediate(
                             instance, 
                             cluster.ClusterWidth, 
                             cluster.ClusterHeight, 
                             isCircular ? cluster.ClusterWidth : 0, 
                             isCircular, 
                             templateZone,
                             isCircular ? (double?)null : cluster.ClusterDepth);
                                        
                        _parameterService.SetClusterSleeveInstanceId(instance, clusterInstanceId);

                        // ✅ CRITICAL FIX: Restore Missing Rotation Logic in Bulk Path
                        if (Math.Abs(cluster.RotationAngleRad) > 1e-6)
                        {
                            try 
                            {
                                XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                                Line axis = Line.CreateBound(location, location + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(doc, instance.Id, axis, cluster.RotationAngleRad);
                            }
                            catch (Exception rotEx)
                            {
                                SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Rotation failed for cluster {cluster.ClusterGUID}: {rotEx.Message}\n");
                            }
                        }

                        placedCount++;
                    }
                    else
                    {
                        statusUpdates.Add((cluster.ClusterGUID, "Failed", "Instance null after creation", -1));
                    }
                }
                catch(Exception ex)
                {
                     statusUpdates.Add((cluster.ClusterGUID, "Failed", $"Post-processing error: {ex.Message}", -1));
                     SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ Cluster {cluster.ClusterGUID} Post-processing failed: {ex.Message}\n");
                }
            }

            using (var dbTracker = _performanceMonitor?.TrackOperation("Post-Placement DB Updates (Bulk)"))
            {
                if (allToDeleteRevitIds.Any())
                {
                    try
                    {
                        var distinctDeleteIds = allToDeleteRevitIds.Distinct().ToList();
                        doc.Delete(distinctDeleteIds);
                        SafeFileLogger.SafeAppendText("batch_v2.log", $"[{DateTime.Now:HH:mm:ss}] 🗑️ BULK DELETED {distinctDeleteIds.Count} individual sleeves\n");
                    }
                    catch (Exception delEx)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Bulk Delete failed: {delEx.Message}\n");
                    }
                }

                if (statusUpdates.Any()) UpdateStatusesBulk(statusUpdates);
                if (calculatedColumnUpdates.Any()) UpdateClashZonesCalculatedColumnsBulk(calculatedColumnUpdates);
                if (allRepositoryFlagUpdates.Any()) _repository.BatchUpdateFlags(allRepositoryFlagUpdates);
                if (oldClusterIdsToDelete.Any()) DeleteOldClusterSleeves(oldClusterIdsToDelete.Distinct().ToList());
                if (legacySleeveInserts.Any()) SaveToClusterSleevesLegacyBulk(doc, legacySleeveInserts, zoneCache);
            }

            return (placedCount, clusters.Count - placedCount);
        }

        private void PostProcessPlacedCluster(Document doc, FamilyInstance instance, BatchClusterData cluster, Dictionary<Guid, ClashZone> zoneCache)
        {
             int clusterInstanceId = instance.Id.IntegerValue;

             ClashZone templateZone = null;
             if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
             {
                 var firstGuidStr = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Trim();
                 if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid g))
                 {
                     if (zoneCache != null && zoneCache.TryGetValue(g, out var cachedZone)) templateZone = cachedZone;
                     else templateZone = _repository.GetClashZonesByGuids(new List<Guid> { g }).FirstOrDefault();
                 }
             }

             if (templateZone == null) templateZone = new ClashZone();
                      
             templateZone.CalculatedSleeveWidth = cluster.ClusterWidth;
             templateZone.CalculatedSleeveHeight = cluster.ClusterHeight;
             templateZone.CalculatedSleeveDepth = cluster.ClusterDepth;
             templateZone.CalculatedRotation = cluster.RotationAngleRad;
                    
             PerformSwapDeletion(doc, cluster, clusterInstanceId, zoneCache);
             UpdateStatus(cluster.ClusterGUID, "Placed", null, clusterInstanceId);
             SaveToClusterSleevesLegacy(doc, cluster, clusterInstanceId, zoneCache);
             
             bool isCircular = cluster.FamilyName.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0 || cluster.FamilyName.IndexOf("Circular", StringComparison.OrdinalIgnoreCase) >= 0;
             _parameterService.SetSleeveParameters(
                  instance, 
                  cluster.ClusterWidth, 
                  cluster.ClusterHeight, 
                  isCircular ? cluster.ClusterWidth : 0, 
                  isCircular, 
                  templateZone,
                  isCircular ? (double?)null : cluster.ClusterDepth);
                            
             _parameterService.SetClusterSleeveInstanceId(instance, clusterInstanceId);
        }

        private List<BatchClusterData> GetPendingClusters(string batchId)
        {
            var list = new List<BatchClusterData>();
            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT ClusterGUID, PlacementX, PlacementY, PlacementZ, 
                                ClusterWidth, ClusterHeight, ClusterDepth, RotationAngleRad,
                                HostElementId, FamilyName, ConstituentZoneGuids
                        FROM ClusterSleeves_v2
                        WHERE (ClusterBatchId = @batch OR @batch IS NULL) AND Status = 'Pending'";
                    if (string.IsNullOrEmpty(batchId))
                        cmd.Parameters.AddWithValue("@batch", DBNull.Value);
                    else
                        cmd.Parameters.AddWithValue("@batch", batchId);
                    
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            list.Add(new BatchClusterData
                            {
                                ClusterGUID = reader["ClusterGUID"].ToString(),
                                PlacementX = Convert.ToDouble(reader["PlacementX"]),
                                PlacementY = Convert.ToDouble(reader["PlacementY"]),
                                PlacementZ = Convert.ToDouble(reader["PlacementZ"]),
                                ClusterWidth = Convert.ToDouble(reader["ClusterWidth"]),
                                ClusterHeight = Convert.ToDouble(reader["ClusterHeight"]),
                                ClusterDepth = Convert.ToDouble(reader["ClusterDepth"]),
                                RotationAngleRad = Convert.ToDouble(reader["RotationAngleRad"]),
                                HostElementId = Convert.ToInt64(reader["HostElementId"]),
                                FamilyName = reader["FamilyName"].ToString(),
                                ConstituentZoneGuids = reader["ConstituentZoneGuids"].ToString()
                            });
                        }
                    }
                }
            }
            return list;
        }

        private void UpdateStatus(string guid, string status, string msg = null, int instanceId = -1)
        {
             using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        UPDATE ClusterSleeves_v2 
                        SET Status = @status, 
                            ValidationMessage = @msg, 
                            ClusterInstanceId = @id,
                            PlacedAt = CURRENT_TIMESTAMP
                        WHERE ClusterGUID = @guid";
                    cmd.Parameters.AddWithValue("@status", status);
                    cmd.Parameters.AddWithValue("@msg", msg ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@id", instanceId);
                    cmd.Parameters.AddWithValue("@guid", guid);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void PerformSwapDeletion(Document doc, BatchClusterData cluster, int clusterElementId, Dictionary<Guid, ClashZone> zoneCache = null)
        {
            if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids)) return;

            var guids = new List<Guid>();
            var guidStrings = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var guidStr in guidStrings)
            {
                var trimmed = guidStr.Trim();
                if (Guid.TryParse(trimmed, out Guid parsedGuid)) guids.Add(parsedGuid);
            }

            if (!guids.Any()) return;

            var zones = new List<ClashZone>();
            var missingGuids = new List<Guid>();
            
            if (zoneCache != null)
            {
                foreach (var g in guids)
                {
                    if (zoneCache.TryGetValue(g, out var z)) zones.Add(z);
                    else missingGuids.Add(g);
                }
            }
            else missingGuids = guids;
            
            if (missingGuids.Any())
            {
                var fetched = _repository.GetClashZonesByGuids(missingGuids);
                zones.AddRange(fetched);
                if (zoneCache != null) foreach (var f in fetched) zoneCache[f.Id] = f;
            }

            var oldClusterInstanceIds = zones
                .Where(z => z.ClusterInstanceId > 0 && z.ClusterInstanceId != clusterElementId)
                .Select(z => z.ClusterInstanceId)
                .Distinct()
                .ToList();

            if (oldClusterInstanceIds.Any()) DeleteOldClusterSleeves(oldClusterInstanceIds);

            var toDeleteIds = new List<ElementId>();
            var updates = new List<(Guid ClashZoneId, bool IsResolved, bool IsClusterResolved, bool IsCombinedResolved, int SleeveInstanceId, int ClusterInstanceId, bool IsClusteredFlag, bool MarkedForClusterProcess, int AfterClusterSleeveId)>();

            foreach (var z in zones)
            {
                if (z.SleeveInstanceId > 0) toDeleteIds.Add(new ElementId(z.SleeveInstanceId));
                updates.Add((z.Id, true, true, false, -1, clusterElementId, true, true, clusterElementId));
            }

            if (toDeleteIds.Any())
            {
                try { doc.Delete(toDeleteIds); }
                catch (Exception delEx) { SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Bulk Delete failed: {delEx.Message}\n"); }
            }

            _repository.BatchUpdateFlags(updates);
            UpdateClashZonesCalculatedColumns(guids, cluster, clusterElementId);
        }

        private void UpdateClashZonesCalculatedColumns(List<Guid> zoneGuids, BatchClusterData cluster, int clusterInstanceId)
        {
            if (zoneGuids == null || !zoneGuids.Any()) return;

            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = transaction;
                            var guidParams = new List<string>();
                            for (int i = 0; i < zoneGuids.Count; i++)
                            {
                                guidParams.Add($"@Guid{i}");
                                cmd.Parameters.AddWithValue($"@Guid{i}", zoneGuids[i].ToString());
                            }
                            
                            cmd.CommandText = $@"
                                UPDATE ClashZones 
                                SET 
                                    CalculatedSleeveWidth = @Width,
                                    CalculatedSleeveHeight = @Height,
                                    CalculatedSleeveDepth = @Depth,
                                    CalculatedRotation = @Rotation,
                                    CalculatedFamilyName = @FamilyName,
                                    PlacedAt = CURRENT_TIMESTAMP,
                                    PlacementStatus = 'Placed',
                                    ClusterInstanceId = @ClusterInstanceId,
                                    IsClusterResolvedFlag = 1,
                                    SleeveState = 2,
                                    UpdatedAt = CURRENT_TIMESTAMP
                                WHERE ClashZoneGuid IN ({string.Join(", ", guidParams)})";
                            
                            cmd.Parameters.AddWithValue("@Width", cluster.ClusterWidth);
                            cmd.Parameters.AddWithValue("@Height", cluster.ClusterHeight);
                            cmd.Parameters.AddWithValue("@Depth", cluster.ClusterDepth);
                            cmd.Parameters.AddWithValue("@Rotation", cluster.RotationAngleRad);
                            cmd.Parameters.AddWithValue("@FamilyName", cluster.FamilyName ?? "");
                            cmd.Parameters.AddWithValue("@ClusterInstanceId", clusterInstanceId);
                            cmd.ExecuteNonQuery();
                        }
                        transaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction.Rollback();
                        SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ UpdateClashZonesCalculatedColumns failed: {ex.Message}\n");
                        throw;
                    }
                }
            }
        }

        private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, int clusterInstanceId, Dictionary<Guid, ClashZone> zoneCache = null)
        {
            try
            {
                int comboId = -1;
                int filterId = -1;
                string category = "";
                string hostType = "";
                string hostOrientation = "";

                if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                {
                    var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
                    if (!string.IsNullOrEmpty(firstGuidStr) && Guid.TryParse(firstGuidStr, out Guid firstGuid))
                    {
                        ClashZone firstZone = null;
                        if (zoneCache != null && zoneCache.TryGetValue(firstGuid, out firstZone)) { }
                        else firstZone = _repository.GetClashZonesByGuids(new List<Guid> { firstGuid }).FirstOrDefault();

                        if (firstZone != null)
                        {
                            comboId = firstZone.ComboId;
                            category = firstZone.MepElementCategory ?? "";
                            hostType = firstZone.StructuralElementType ?? "";
                            hostOrientation = firstZone.HostOrientation ?? "";

                            var zoneGuids = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(g => Guid.Parse(g.Trim())).ToList();
                            double halfWidth = cluster.ClusterWidth / 2.0;
                            double halfHeight = cluster.ClusterHeight / 2.0;
                            double halfDepth = cluster.ClusterDepth / 2.0;

                            using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(doc))
                            {
                                if (comboId > 0)
                                {
                                    using (var cmd = dbContext.Connection.CreateCommand())
                                    {
                                        cmd.CommandText = "SELECT FilterId FROM FileCombos WHERE ComboId = @ComboId LIMIT 1";
                                        cmd.Parameters.AddWithValue("@ComboId", comboId);
                                        var res = cmd.ExecuteScalar();
                                        if (res != null && res != DBNull.Value) filterId = Convert.ToInt32(res);
                                    }
                                }

                                if (comboId <= 0 || filterId <= 0) throw new InvalidOperationException("Missing IDs");

                                var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveRepository(dbContext);
                                repo.SaveClusterSleeve(
                                    clusterInstanceId: clusterInstanceId,
                                    comboId: comboId,
                                    filterId: filterId,
                                    category: category,
                                    boundingBoxMinX: cluster.PlacementX - halfWidth,
                                    boundingBoxMinY: cluster.PlacementY - halfHeight,
                                    boundingBoxMinZ: cluster.PlacementZ - halfDepth,
                                    boundingBoxMaxX: cluster.PlacementX + halfWidth,
                                    boundingBoxMaxY: cluster.PlacementY + halfHeight,
                                    boundingBoxMaxZ: cluster.PlacementZ + halfDepth,
                                    clusterWidth: cluster.ClusterWidth,
                                    clusterHeight: cluster.ClusterHeight,
                                    clusterDepth: cluster.ClusterDepth,
                                    rotationAngleDeg: cluster.RotationAngleRad * (180.0 / Math.PI),
                                    isRotated: Math.Abs(cluster.RotationAngleRad) > 1e-6,
                                    placementX: cluster.PlacementX,
                                    placementY: cluster.PlacementY,
                                    placementZ: cluster.PlacementZ,
                                    hostType: hostType,
                                    hostOrientation: hostOrientation,
                                    clashZoneIds: zoneGuids,
                                    sleeveFamilyName: cluster.FamilyName,
                                    corner1X: cluster.PlacementX - halfWidth * Math.Cos(cluster.RotationAngleRad) + halfHeight * Math.Sin(cluster.RotationAngleRad),
                                    corner1Y: cluster.PlacementY - halfWidth * Math.Sin(cluster.RotationAngleRad) - halfHeight * Math.Cos(cluster.RotationAngleRad),
                                    corner1Z: cluster.PlacementZ,
                                    corner2X: cluster.PlacementX + halfWidth * Math.Cos(cluster.RotationAngleRad) + halfHeight * Math.Sin(cluster.RotationAngleRad),
                                    corner2Y: cluster.PlacementY + halfWidth * Math.Sin(cluster.RotationAngleRad) - halfHeight * Math.Cos(cluster.RotationAngleRad),
                                    corner2Z: cluster.PlacementZ,
                                    corner3X: cluster.PlacementX + halfWidth * Math.Cos(cluster.RotationAngleRad) - halfHeight * Math.Sin(cluster.RotationAngleRad),
                                    corner3Y: cluster.PlacementY + halfWidth * Math.Sin(cluster.RotationAngleRad) + halfHeight * Math.Cos(cluster.RotationAngleRad),
                                    corner3Z: cluster.PlacementZ,
                                    corner4X: cluster.PlacementX - halfWidth * Math.Cos(cluster.RotationAngleRad) - halfHeight * Math.Sin(cluster.RotationAngleRad),
                                    corner4Y: cluster.PlacementY - halfWidth * Math.Sin(cluster.RotationAngleRad) + halfHeight * Math.Cos(cluster.RotationAngleRad),
                                    corner4Z: cluster.PlacementZ
                                );
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { SafeFileLogger.SafeAppendText("placement_errors.log", $"[{DateTime.Now:HH:mm:ss}] ❌ SaveToClusterSleevesLegacy FAILED: {ex.Message}\n"); throw; }
        }

        private bool PlaceSingleCluster(Document doc, BatchClusterData cluster, Dictionary<string, FamilySymbol> symbolCache = null, Dictionary<Guid, ClashZone> zoneCache = null, JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.IOperationTracker parentTracker = null, List<FamilyInstance> placedInstances = null)
        {
            try
            {
                FamilySymbol symbol = null;
                if (symbolCache != null && symbolCache.TryGetValue(cluster.FamilyName, out var cachedSymbol)) symbol = cachedSymbol;
                else symbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().FirstOrDefault(x => x.Name == cluster.FamilyName || x.Family.Name == cluster.FamilyName);

                if (symbol == null) { UpdateStatus(cluster.ClusterGUID, "Failed", "Missing Symbol"); return false; }
                if (!symbol.IsActive) symbol.Activate();

                Element host = null;
                Level level = null;
                if (cluster.HostElementId > 0) { try { host = doc.GetElement(new ElementId((int)cluster.HostElementId)); } catch { } if (host != null) level = doc.GetElement(host.LevelId) as Level; }
                if (level == null) level = new FilteredElementCollector(doc).OfClass(typeof(Level)).FirstOrDefault() as Level;

                XYZ location = new XYZ(cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ);
                FamilyInstance instance = null;
                if (host != null) instance = doc.Create.NewFamilyInstance(location, symbol, host, level, StructuralType.NonStructural);
                else instance = doc.Create.NewFamilyInstance(location, symbol, StructuralType.NonStructural);

                if (instance != null)
                {
                    int clusterInstanceId = instance.Id.IntegerValue;
                    ClashZone templateZone = null;
                    if (!string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
                    {
                        var firstGuidStr = cluster.ConstituentZoneGuids.Split(',').FirstOrDefault()?.Trim();
                        if (Guid.TryParse(firstGuidStr, out Guid g))
                        {
                            if (zoneCache != null && zoneCache.TryGetValue(g, out var cachedZone)) templateZone = cachedZone;
                            else templateZone = _repository.GetClashZonesByGuids(new List<Guid> { g }).FirstOrDefault();
                        }
                    }
                    if (templateZone == null) templateZone = new ClashZone();
                    templateZone.CalculatedSleeveWidth = cluster.ClusterWidth;
                    templateZone.CalculatedSleeveHeight = cluster.ClusterHeight;
                    templateZone.CalculatedSleeveDepth = cluster.ClusterDepth;
                    templateZone.CalculatedRotation = cluster.RotationAngleRad;
                    
                    PerformSwapDeletion(doc, cluster, clusterInstanceId, zoneCache);
                    UpdateStatus(cluster.ClusterGUID, "Placed", null, clusterInstanceId);
                    SaveToClusterSleevesLegacy(doc, cluster, clusterInstanceId, zoneCache);
                    
                    bool isCircular = cluster.FamilyName.IndexOf("Round", StringComparison.OrdinalIgnoreCase) >= 0 || cluster.FamilyName.IndexOf("Circular", StringComparison.OrdinalIgnoreCase) >= 0;
                    
                    // ✅ DIAGNOSTIC: Log cluster parameter write
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_params.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🔧 CLUSTER PARAMS: Instance={clusterInstanceId}, " +
                            $"W={cluster.ClusterWidth*304.8:F1}mm, H={cluster.ClusterHeight*304.8:F1}mm, D={cluster.ClusterDepth*304.8:F1}mm, " +
                            $"IsCircular={isCircular}, Family={cluster.FamilyName}\n");
                    }
                    
                    // ✅ CRITICAL FIX: Force immediate write for cluster parameters (batching causes parameters to never be written)
                    _parameterService.SetSleeveParametersImmediate(instance, cluster.ClusterWidth, cluster.ClusterHeight, isCircular ? cluster.ClusterWidth : 0, isCircular, templateZone, isCircular ? (double?)null : cluster.ClusterDepth);
                    _parameterService.SetClusterSleeveInstanceId(instance, clusterInstanceId);
                    
                    // ✅ DIAGNOSTIC: Confirm parameter write completed
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_params.log", 
                            $"[{DateTime.Now:HH:mm:ss}] ✅ PARAMS SET: Instance={clusterInstanceId}\n");
                    }

                    if (Math.Abs(cluster.RotationAngleRad) > 1e-6)
                    {
                        Line axis = Line.CreateBound(location, location + XYZ.BasisZ);
                        ElementTransformUtils.RotateElement(doc, instance.Id, axis, cluster.RotationAngleRad);
                    }
                    placedInstances?.Add(instance);
                    return true;
                }
                UpdateStatus(cluster.ClusterGUID, "Failed", "Creation returned null");
                return false;
            }
            catch (Exception ex) { UpdateStatus(cluster.ClusterGUID, "Failed", ex.Message); return false; }
        }

        private void DeleteOldClusterSleeves(List<int> clusterInstanceIds)
        {
            if (clusterInstanceIds == null || !clusterInstanceIds.Any()) return;
            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    try {
                        using (var cmd = conn.CreateCommand()) {
                            cmd.Transaction = transaction;
                            var ids = string.Join(",", clusterInstanceIds);
                            cmd.CommandText = $"DELETE FROM ClusterSleeves WHERE ClusterInstanceId IN ({ids})";
                            cmd.ExecuteNonQuery();
                            
                            // ✅ V2: Also delete from ClusterSleeves_v2
                            cmd.CommandText = $"DELETE FROM ClusterSleeves_v2 WHERE ClusterInstanceId IN ({ids})";
                            cmd.ExecuteNonQuery();
                        }
                        transaction.Commit();
                    } catch { transaction.Rollback(); }
                }
            }
        }

        private void UpdateStatusesBulk(List<(string Guid, string Status, string Msg, int InstanceId)> updates)
        {
            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var trans = conn.BeginTransaction())
                {
                    try {
                        using (var cmd = conn.CreateCommand()) {
                            cmd.Transaction = trans;
                            cmd.CommandText = "UPDATE ClusterSleeves_v2 SET Status = @status, ValidationMessage = @msg, ClusterInstanceId = @id, PlacedAt = CURRENT_TIMESTAMP WHERE ClusterGUID = @guid";
                            var pStatus = cmd.Parameters.Add("@status", System.Data.DbType.String);
                            var pMsg = cmd.Parameters.Add("@msg", System.Data.DbType.String);
                            var pId = cmd.Parameters.Add("@id", System.Data.DbType.Int32);
                            var pGuid = cmd.Parameters.Add("@guid", System.Data.DbType.String);
                            foreach (var u in updates) { 
                                pStatus.Value = u.Status; pMsg.Value = u.Msg ?? (object)DBNull.Value; pId.Value = u.InstanceId; pGuid.Value = u.Guid; 
                                cmd.ExecuteNonQuery(); 
                            }
                        }
                        trans.Commit();
                    } catch { trans.Rollback(); }
                }
            }
        }

        private void UpdateClashZonesCalculatedColumnsBulk(List<(List<Guid> Guids, BatchClusterData Cluster, int InstanceId)> updates)
        {
            using (var conn = new SQLiteConnection($"Data Source={_databasePath};Version=3;"))
            {
                conn.Open();
                using (var trans = conn.BeginTransaction())
                {
                    try {
                        using (var cmd = conn.CreateCommand()) {
                            cmd.Transaction = trans;
                            foreach (var u in updates) {
                                cmd.Parameters.Clear();
                                var guidParams = new List<string>();
                                for (int i = 0; i < u.Guids.Count; i++) { guidParams.Add($"@Guid{i}"); cmd.Parameters.AddWithValue($"@Guid{i}", u.Guids[i].ToString()); }
                                cmd.CommandText = $@"UPDATE ClashZones SET CalculatedSleeveWidth = @Width, CalculatedSleeveHeight = @Height, CalculatedSleeveDepth = @Depth, CalculatedRotation = @Rotation, CalculatedFamilyName = @FamilyName, PlacedAt = CURRENT_TIMESTAMP, PlacementStatus = 'Placed', ClusterInstanceId = @ClusterInstanceId, IsClusterResolvedFlag = 1, SleeveState = 2, UpdatedAt = CURRENT_TIMESTAMP WHERE ClashZoneGuid IN ({string.Join(", ", guidParams)})";
                                cmd.Parameters.AddWithValue("@Width", u.Cluster.ClusterWidth); cmd.Parameters.AddWithValue("@Height", u.Cluster.ClusterHeight); cmd.Parameters.AddWithValue("@Depth", u.Cluster.ClusterDepth); cmd.Parameters.AddWithValue("@Rotation", u.Cluster.RotationAngleRad); cmd.Parameters.AddWithValue("@FamilyName", u.Cluster.FamilyName ?? ""); cmd.Parameters.AddWithValue("@ClusterInstanceId", u.InstanceId);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        trans.Commit();
                    } catch { trans.Rollback(); }
                }
            }
        }

        private void SaveToClusterSleevesLegacyBulk(Document doc, List<(BatchClusterData Cluster, int InstanceId)> updates, Dictionary<Guid, ClashZone> zoneCache)
        {
            try {
                using (var dbContext = new JSE_RevitAddin_MEP_OPENINGS.Data.SleeveDbContext(doc)) {
                    var repo = new JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveRepository(dbContext);
                    using (var trans = dbContext.Connection.BeginTransaction()) {
                        foreach (var u in updates) {
                            var cluster = u.Cluster;
                            var clusterInstanceId = u.InstanceId;
                            int comboId = -1; int filterId = -1; string category = ""; string hostType = ""; string hostOrientation = "";
                            var zoneGuidsStrings = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            if (zoneGuidsStrings.Any()) {
                                if (Guid.TryParse(zoneGuidsStrings[0].Trim(), out Guid firstGuid)) {
                                    ClashZone firstZone = null;
                                    if (zoneCache != null && zoneCache.TryGetValue(firstGuid, out firstZone)) { }
                                    else firstZone = _repository.GetClashZonesByGuids(new List<Guid> { firstGuid }).FirstOrDefault();
                                    if (firstZone != null) {
                                        comboId = firstZone.ComboId; category = firstZone.MepElementCategory ?? ""; hostType = firstZone.StructuralElementType ?? ""; hostOrientation = firstZone.HostOrientation ?? "";
                                        if (comboId > 0) {
                                            using (var cmd = dbContext.Connection.CreateCommand()) { cmd.Transaction = trans; cmd.CommandText = "SELECT FilterId FROM FileCombos WHERE ComboId = @ComboId LIMIT 1"; cmd.Parameters.AddWithValue("@ComboId", comboId); var res = cmd.ExecuteScalar(); if (res != null && res != DBNull.Value) filterId = Convert.ToInt32(res); }
                                        }
                                    }
                                }
                            }
                            if (comboId <= 0 || filterId <= 0) continue;
                            var guids = zoneGuidsStrings.Select(g => Guid.Parse(g.Trim())).ToList();
                            double halfWidth = cluster.ClusterWidth / 2.0; double halfHeight = cluster.ClusterHeight / 2.0; double halfDepth = cluster.ClusterDepth / 2.0;
                            repo.SaveClusterSleeve(clusterInstanceId, comboId, filterId, category, cluster.PlacementX - halfWidth, cluster.PlacementY - halfHeight, cluster.PlacementZ - halfDepth, cluster.PlacementX + halfWidth, cluster.PlacementY + halfHeight, cluster.PlacementZ + halfDepth, cluster.ClusterWidth, cluster.ClusterHeight, cluster.ClusterDepth, cluster.RotationAngleRad * (180.0 / Math.PI), Math.Abs(cluster.RotationAngleRad) > 1e-6, cluster.PlacementX, cluster.PlacementY, cluster.PlacementZ, hostType, hostOrientation, guids, cluster.FamilyName, cluster.PlacementX - halfWidth * Math.Cos(cluster.RotationAngleRad) + halfHeight * Math.Sin(cluster.RotationAngleRad), cluster.PlacementY - halfWidth * Math.Sin(cluster.RotationAngleRad) - halfHeight * Math.Cos(cluster.RotationAngleRad), cluster.PlacementZ, cluster.PlacementX + halfWidth * Math.Cos(cluster.RotationAngleRad) + halfHeight * Math.Sin(cluster.RotationAngleRad), cluster.PlacementY + halfWidth * Math.Sin(cluster.RotationAngleRad) - halfHeight * Math.Cos(cluster.RotationAngleRad), cluster.PlacementZ, cluster.PlacementX + halfWidth * Math.Cos(cluster.RotationAngleRad) - halfHeight * Math.Sin(cluster.RotationAngleRad), cluster.PlacementY + halfWidth * Math.Sin(cluster.RotationAngleRad) + halfHeight * Math.Cos(cluster.RotationAngleRad), cluster.PlacementZ, cluster.PlacementX - halfWidth * Math.Cos(cluster.RotationAngleRad) - halfHeight * Math.Sin(cluster.RotationAngleRad), cluster.PlacementY - halfWidth * Math.Sin(cluster.RotationAngleRad) + halfHeight * Math.Cos(cluster.RotationAngleRad), cluster.PlacementZ);
                        }
                        trans.Commit();
                    }
                }
            } catch { }
        }
    }

    public class BatchClusterData
    {
        public string ClusterGUID { get; set; }
        public double PlacementX { get; set; }
        public double PlacementY { get; set; }
        public double PlacementZ { get; set; }
        public double ClusterWidth { get; set; }
        public double ClusterHeight { get; set; }
        public double ClusterDepth { get; set; }
        public double RotationAngleRad { get; set; }
        public long HostElementId { get; set; }
        public string FamilyName { get; set; }
        public string ConstituentZoneGuids { get; set; }
    }
}
