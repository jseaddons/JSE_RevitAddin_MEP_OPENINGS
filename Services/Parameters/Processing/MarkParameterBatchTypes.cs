using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Processing
{
    /// <summary>
    /// Represents the identity data for a sleeve during mark calculation.
    /// Decouples Revit Element from logic.
    /// </summary>
    public class MarkSleeveIdentity
    {
        public ElementId SleeveId { get; set; }
        public int SleeveInstanceId { get; set; } // The ID as integer
        public string FamilyName { get; set; }
        public long MepElementId { get; set; }
        public string ExistingMark { get; set; }
        public string LevelName { get; set; } // For sorting/grouping
        public XYZ LocationPoint { get; set; } // For sorting
        
        // Parameters used for prefix resolution
        public string SystemType { get; set; }
        public string ServiceType { get; set; }
        public string MepCategory { get; set; } // From parameter if available, else from DB
        
        // ✅ NEW: Flag for combined sleeves (gets MEP prefix)
        public bool IsCombinedSleeve { get; set; } = false;

        // For Calculation Result
        public string CalculatedMark { get; set; }
    }
}
