using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

using System.IO;
namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    // Partial extension to ParameterTransferService for Batch Operations
    public partial class ParameterTransferService
    {
        /// <summary>
        /// Optimized Batch Transfer Implementation (Read-Calculate-Write)
        /// Uses Parallel processing for calculations and Single Transaction for writes.
        /// </summary>
        public ParameterTransferResult ExecuteBatchTransferInTransaction(
            Document doc,
            List<ElementId> openingIds,
            ParameterTransferConfiguration config)
        {
            var result = new ParameterTransferResult();
            var successfullyTransferredSleeveIds = new HashSet<int>();
            
            // LOGGING SETUP
            string debugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");
            var stopwatchTotal = System.Diagnostics.Stopwatch.StartNew();
            var stopwatchRead = new System.Diagnostics.Stopwatch();
            var stopwatchLoadSnapshot = new System.Diagnostics.Stopwatch();
            var stopwatchCalc = new System.Diagnostics.Stopwatch();
            var stopwatchWrite = new System.Diagnostics.Stopwatch();

            if (!DeploymentConfiguration.DeploymentMode)
            {
                File.AppendAllText(debugLogPath, $"\n===== BATCH PARAM_TRANSFER SESSION STARTED {DateTime.Now} =====\n");
            }

            // 1. VALIDATION & SETUP
            if (openingIds == null || openingIds.Count == 0)
            {
                result.Success = false;
                result.Message = "No sleeves found.";
                return result;
            }

            // Load Snapshots
            stopwatchLoadSnapshot.Start();
            SleeveSnapshotIndex snapshotIndex = null;
            try
            {
                using (var dbContext = new SleeveDbContext(doc, msg => { if(!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg); }))
                {
                    var repo = new SleeveSnapshotRepository(dbContext, msg => { if(!DeploymentConfiguration.DeploymentMode) DebugLogger.Info(msg); });
                    snapshotIndex = repo.LoadSnapshotIndex();
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Failed to load snapshots: {ex.Message}";
                return result;
            }
            stopwatchLoadSnapshot.Stop();

            if (snapshotIndex.IsEmpty)
            {
                 // Handle empty DB logic
                 bool checkDbHasSleeves = false;
                 try 
                 {
                     using (var dbContext = new SleeveDbContext(doc, msg => { }))
                     {
                         using (var cmd = dbContext.Connection.CreateCommand())
                         {
                             cmd.CommandText = "SELECT COUNT(*) FROM ClashZones WHERE SleeveInstanceId > 0";
                             checkDbHasSleeves = Convert.ToInt32(cmd.ExecuteScalar()) > 0;
                         }
                     }
                 } catch {}

                 if (checkDbHasSleeves)
                 {
                     result.Message = "No snapshots loaded, but sleeves exist in DB. Run Refresh.";
                     if (!DeploymentConfiguration.DeploymentMode) DebugLogger.Warning(result.Message);
                 }
                 else
                 {
                     result.Message = "No snapshots and no sleeves in DB.";
                 }
                 result.Success = false;
                 return result;
            }

            // 2. READ PHASE (Main Thread)
            stopwatchRead.Start();
            var identities = new List<SleeveIdentity>(openingIds.Count);
            foreach (var id in openingIds)
            {
                var el = doc.GetElement(id);
                if (el == null) continue;
                
                var identity = new SleeveIdentity
                {
                    OpeningId = id,
                    SleeveInstanceId = GetIntegerParameter(el, "Sleeve Instance ID"),
                    ClusterInstanceId = GetIntegerParameter(el, "Cluster Sleeve Instance ID"),
                    CombinedInstanceId = id.IntegerValue
                };
                identities.Add(identity);
            }
            stopwatchRead.Stop();

            // 3. CALCULATE PHASE (Parallel)
            stopwatchCalc.Start();
            var updateActions = new System.Collections.Concurrent.ConcurrentBag<ParameterUpdateAction>();
            
            System.Threading.Tasks.Parallel.ForEach(identities, identity =>
            {
                SleeveSnapshotView snapshot = null;

                // Priority 1: Combined Sleeve
                if (snapshotIndex.TryGetByCombined(identity.CombinedInstanceId, out var constituents))
                {
                    var aggregatedParams = AggregateCombinedParameters(constituents, snapshotIndex, useHost: false);
                    var aggregatedHostParams = AggregateCombinedParameters(constituents, snapshotIndex, useHost: true);
                    snapshot = new SleeveSnapshotView 
                    { 
                        SleeveInstanceId = identity.SleeveInstanceId,
                        MepParameters = aggregatedParams,
                        HostParameters = aggregatedHostParams
                    };
                }
                // Priority 2: Individual Sleeve
                else if (identity.SleeveInstanceId > 0)
                {
                     if (snapshotIndex.TryGetBySleeve(identity.SleeveInstanceId, out var s))
                         snapshot = s;
                     else if (snapshotIndex.SleeveIdToClashZoneGuid.TryGetValue(identity.SleeveInstanceId, out var guid))
                     {
                         if (snapshotIndex.TryGetByClashZoneGuid(guid, out var sGuid)) snapshot = sGuid;
                     }
                }
                // Priority 3: Cluster Sleeve
                else if (identity.ClusterInstanceId > 0)
                {
                    if (snapshotIndex.TryGetByCluster(identity.ClusterInstanceId, out var sC))
                        snapshot = sC;
                }

                if (snapshot != null)
                {
                    foreach (var mapping in config.Mappings)
                    {
                        if (!mapping.IsEnabled) continue;

                        string resolvedValue = null;
                        
                        // Handle Level separately if needed
                        if (mapping.TransferType == TransferType.LevelToOpening)
                        {
                            continue; 
                        }

                        var sourceParams = (mapping.TransferType == TransferType.HostToOpening) ? snapshot.HostParameters : snapshot.MepParameters;
                        
                        if (sourceParams != null)
                        {
                            if (sourceParams.TryGetValue(mapping.SourceParameter, out var val))
                                resolvedValue = val;
                            else if (mapping.SourceParameter.Equals("MEP Size", StringComparison.OrdinalIgnoreCase) || mapping.SourceParameter.Equals("Size", StringComparison.OrdinalIgnoreCase))
                            {
                                 if (sourceParams.TryGetValue("Size", out var v1)) resolvedValue = v1;
                                 else if (sourceParams.TryGetValue("MEP Size", out var v2)) resolvedValue = v2;
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(resolvedValue))
                        {
                            updateActions.Add(new ParameterUpdateAction
                            {
                                OpeningId = identity.OpeningId,
                                TargetParameter = mapping.TargetParameter,
                                Value = resolvedValue
                            });
                        }
                    }
                }
            });
            stopwatchCalc.Stop();

            // 4. WRITE PHASE (Main Thread)
            stopwatchWrite.Start();
            int setSuccessCount = 0;
            int setFailCount = 0;
            
            foreach (var action in updateActions)
            {
                var el = doc.GetElement(action.OpeningId);
                if (el == null) continue;
                
                var param = el.LookupParameter(action.TargetParameter);
                if (SetParameterValueSafely(param, action.Value))
                {
                    setSuccessCount++;
                    successfullyTransferredSleeveIds.Add(action.OpeningId.IntegerValue);
                }
                else
                {
                    setFailCount++;
                }
            }
            
            // Handle Model Names (Sequential)
            if (config.TransferModelNames)
            {
                string modelName = doc.Title;
                foreach(var id in openingIds)
                {
                     var el = doc.GetElement(id);
                     if (el == null) continue;
                     var param = el.LookupParameter(config.ModelNameParameter);
                     if (SetParameterValueSafely(param, modelName)) setSuccessCount++;
                }
            }
            stopwatchWrite.Stop();
            stopwatchTotal.Stop();

            // PERFORMANCE LOGGING
            if (!DeploymentConfiguration.DeploymentMode && openingIds.Count > 0)
            {
                 var msg = $"PERFORMANCE [PARAM_TRANSFER]: Total={stopwatchTotal.ElapsedMilliseconds}ms | " +
                          $"SnapLoad={stopwatchLoadSnapshot.ElapsedMilliseconds}ms | " +
                          $"Read={stopwatchRead.ElapsedMilliseconds}ms | " +
                          $"Calc={stopwatchCalc.ElapsedMilliseconds}ms | " +
                          $"Write={stopwatchWrite.ElapsedMilliseconds}ms ({(setSuccessCount > 0 ? (stopwatchWrite.ElapsedMilliseconds / setSuccessCount) : 0)} ms/param)\n";
                 File.AppendAllText(debugLogPath, msg);
            }

            result.Success = true;
            result.TransferredCount = successfullyTransferredSleeveIds.Count;
            result.FailedCount = setFailCount;
            result.Message = $"Batch Transfer: {result.TransferredCount} sleeves updated ({setSuccessCount} params set).";
            
            return result;
        }
    }
}
