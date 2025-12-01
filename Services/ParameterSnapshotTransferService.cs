using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Implements batched (or immediate) transfer of snapshot MEP parameters to sleeve elements.
    /// Lightweight and focused: no zone loading, only direct snapshot parameter retrieval.
    /// </summary>
    public class ParameterSnapshotTransferService : IParameterSnapshotTransferService
    {
        public int TransferSnapshotParameters(Document doc, IEnumerable<ElementId> sleeveElementIds, IClashZoneRepository repository, IParameterBatchingService batchingService)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (sleeveElementIds == null || repository == null || batchingService == null) return 0;

            var ids = sleeveElementIds
                .Where(e => e != null && e != ElementId.InvalidElementId)
                .Select(e => e.IntegerValue)
                .Distinct()
                .ToList();
            if (ids.Count == 0) return 0;

            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[BATCH-PARAMS] 🚀 Snapshot parameter transfer start for {ids.Count} sleeves (batching={(batchingService.IsBatchingEnabled ? "on" : "off")})");
            }

            var snapshotParams = repository.GetSnapshotMepParametersForSleeveIds(ids);
            int affected = 0;

            if (batchingService.IsBatchingEnabled)
            {
                foreach (var kvp in snapshotParams)
                {
                    var elementId = new ElementId(kvp.Key);
                    foreach (var param in kvp.Value)
                    {
                        batchingService.DeferParameter(elementId, param.Key, param.Value);
                        affected++;
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[BATCH-PARAMS] 📥 Deferred {affected} snapshot parameters for {snapshotParams.Count} sleeves");
                }
            }
            else
            {
                int written = 0;
                using (var tx = new Transaction(doc, "Immediate Snapshot Parameter Transfer"))
                {
                    tx.Start();
                    foreach (var kvp in snapshotParams)
                    {
                        var element = doc.GetElement(new ElementId(kvp.Key));
                        if (element == null) continue;
                        foreach (var param in kvp.Value)
                        {
                            try
                            {
                                var p = element.LookupParameter(param.Key);
                                if (p == null || p.IsReadOnly) continue;
                                if (p.StorageType == StorageType.String)
                                {
                                    if (p.Set(param.Value)) written++; else continue;
                                }
                                else if (p.StorageType == StorageType.Double && double.TryParse(param.Value, out var d))
                                {
                                    if (p.Set(d)) written++; else continue;
                                }
                                else if (p.StorageType == StorageType.Integer && int.TryParse(param.Value, out var iv))
                                {
                                    if (p.Set(iv)) written++; else continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[BATCH-PARAMS] ❌ Failed immediate set '{param.Key}' on {kvp.Key}: {ex.Message}");
                                }
                            }
                        }
                    }
                    tx.Commit();
                }
                affected = written;
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[BATCH-PARAMS] ✅ Immediate snapshot parameter transfer wrote {written} parameters for {snapshotParams.Count} sleeves");
                }
            }

            return affected;
        }
    }
}
