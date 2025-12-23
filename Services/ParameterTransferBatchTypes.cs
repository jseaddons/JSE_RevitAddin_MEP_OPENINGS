using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Represents the identity data extracted from a Revit elements for parameter transfer.
    /// Used to decouple Revit API reads from data processing.
    /// </summary>
    public class SleeveIdentity
    {
        public ElementId OpeningId { get; set; }
        public int SleeveInstanceId { get; set; }
        public int ClusterInstanceId { get; set; }
        public int CombinedInstanceId { get; set; } // Assuming this might be the same as SleeveInstanceId in current Revit schema, or a separate param
        
        // Helper to determine lookup strategy
        public bool IsCluster => ClusterInstanceId > 0;
        public bool IsIndividual => SleeveInstanceId > 0 && ClusterInstanceId <= 0;
    }

    /// <summary>
    /// Represents a calculated parameter update ready to be applied.
    /// </summary>
    public class ParameterUpdateAction
    {
        public ElementId OpeningId { get; set; }
        public string TargetParameter { get; set; }
        public string Value { get; set; }
    }
}
