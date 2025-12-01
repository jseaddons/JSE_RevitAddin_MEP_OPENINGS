using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Service for transferring snapshot-stored MEP parameters from the database to live Revit sleeve elements.
    /// Uses batching when enabled for performance; falls back to immediate writes otherwise.
    /// </summary>
    public interface IParameterSnapshotTransferService
    {
        /// <summary>
        /// Transfer snapshot parameters for given sleeve element ids.
        /// When batching enabled, defers writes into IParameterBatchingService; caller responsible for flush.
        /// When batching disabled, writes parameters immediately inside a transaction.
        /// </summary>
        /// <param name="doc">Active Revit Document</param>
        /// <param name="sleeveElementIds">Collection of sleeve element ids</param>
        /// <param name="repository">Clash zone repository (for snapshot access)</param>
        /// <param name="batchingService">Batching service</param>
        /// <returns>Count of parameters queued or written</returns>
        int TransferSnapshotParameters(Document doc, IEnumerable<ElementId> sleeveElementIds, IClashZoneRepository repository, IParameterBatchingService batchingService);
    }
}
