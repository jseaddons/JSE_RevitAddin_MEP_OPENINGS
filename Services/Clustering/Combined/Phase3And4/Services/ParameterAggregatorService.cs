using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces;

using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Services
{
    /// <summary>
    /// Service responsible for preparing parameter sets for combined sleve family instances.
    /// Phase 3 Implementation.
    /// </summary>
    public class ParameterAggregatorService : IParameterAggregatorService
    {
        /// <summary>
        /// Creates the parameter dictionary to be applied to the new combined sleeve family instance.
        /// </summary>
        public Dictionary<string, object> CreateCombinedSleeveParameterSet(
            CombinedClusterCandidate combinedCluster, 
            BoundingBoxXYZ combinedBoundingBox)
        {
            var paramsToSet = new Dictionary<string, object>();
            
            // Standard parameters
            // Assuming BoundingBox is in internal units (feet)
            paramsToSet["Width"] = combinedBoundingBox.Max.X - combinedBoundingBox.Min.X;
            paramsToSet["Height"] = combinedBoundingBox.Max.Z - combinedBoundingBox.Min.Z;
            paramsToSet["Depth"] = combinedBoundingBox.Max.Y - combinedBoundingBox.Min.Y; 
            
            // Text Parameters
            string cats = string.Join(",", combinedCluster.CategoriesInvolved.OrderBy(c => c));
            int count = combinedCluster.MemberClusters.Count;

            paramsToSet["Comments"] = $"Combined Cluster: {cats} ({count} items)";

            // If damper involved, treat as STANDARD (non-MEP) route: skip MEP-specific fields
            bool isDamper = combinedCluster.CategoriesInvolved.Any(c =>
                c?.IndexOf("damper", StringComparison.OrdinalIgnoreCase) >= 0 ||
                c?.IndexOf("duct accessory", StringComparison.OrdinalIgnoreCase) >= 0 ||
                c?.IndexOf("duct accessories", StringComparison.OrdinalIgnoreCase) >= 0);

            if (!isDamper)
            {
                paramsToSet["MEP_Category"] = "Multi-Service"; 
                // Discipline prefix fixed to MEP for combined sleeves (non-damper)
                paramsToSet["MEP_System_Abbreviation"] = "MEP";
            }
            
            // Serialize aggregated parameters (optional, if JSON is available)
            // var snapshot = combinedCluster.ParameterSnapshots.FlattenDistinct();
            // paramsToSet["CombinedSleeveData"] = ...; 

            return paramsToSet;
        }
    }
}
