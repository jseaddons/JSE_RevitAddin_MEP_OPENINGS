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
            paramsToSet["Width"] = combinedBoundingBox.Max.X - combinedBoundingBox.Min.X;
            paramsToSet["Height"] = combinedBoundingBox.Max.Z - combinedBoundingBox.Min.Z;
            paramsToSet["Depth"] = combinedBoundingBox.Max.Y - combinedBoundingBox.Min.Y;

            // Text Parameters
            string cats = string.Join(",", combinedCluster.CategoriesInvolved.OrderBy(c => c));
            int count = combinedCluster.MemberClusters.Count;
            paramsToSet["Comments"] = $"Combined Cluster: {cats} ({count} items)";

            // Aggregate system type, system name, system abbreviation, and size for all sleeves
            var allParameterSnapshots = new List<Dictionary<string, string>>();
            // Add parameter snapshots from cluster sleeves
            foreach (var cluster in combinedCluster.MemberClusters)
            {
                if (cluster.ParameterSnapshot != null && cluster.ParameterSnapshot.Count > 0)
                    allParameterSnapshots.Add(cluster.ParameterSnapshot);
            }
            // Add parameter snapshots from incorporated individual sleeves (if any)
            if (combinedCluster.IncorporatedIndividualSleeves != null)
            {
                foreach (var cz in combinedCluster.IncorporatedIndividualSleeves)
                {
                    if (cz.MepParameterValues != null && cz.MepParameterValues.Count > 0)
                    {
                        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var kv in cz.MepParameterValues)
                        {
                            dict[kv.Key] = kv.Value;
                        }
                        allParameterSnapshots.Add(dict);
                    }
                }
            }

            // Helper to aggregate unique values for a given key
            string AggregateValues(string key)
            {
                return string.Join(",", allParameterSnapshots
                    .Where(snap => snap.ContainsKey(key) && !string.IsNullOrWhiteSpace(snap[key]))
                    .Select(snap => snap[key])
                    .Distinct()
                );
            }

            // System Type, System Name, System Abbreviation, Size
            paramsToSet["MEP_System_Type"] = AggregateValues("System Type");
            paramsToSet["MEP_System_Name"] = AggregateValues("System Name");
            paramsToSet["MEP_System_Abbreviation"] = AggregateValues("System Abbreviation");
            paramsToSet["MEP_Size"] = AggregateValues("Size");

            // If damper involved, treat as STANDARD (non-MEP) route: skip MEP-specific fields
            bool isDamper = combinedCluster.CategoriesInvolved.Any(c =>
                c?.IndexOf("damper", StringComparison.OrdinalIgnoreCase) >= 0 ||
                c?.IndexOf("duct accessory", StringComparison.OrdinalIgnoreCase) >= 0 ||
                c?.IndexOf("duct accessories", StringComparison.OrdinalIgnoreCase) >= 0);

            if (!isDamper)
            {
                paramsToSet["MEP_Category"] = "Multi-Service";
                // Discipline prefix fixed to MEP for combined sleeves (non-damper)
                if (!paramsToSet.ContainsKey("MEP_System_Abbreviation") || string.IsNullOrWhiteSpace(paramsToSet["MEP_System_Abbreviation"]?.ToString()))
                    paramsToSet["MEP_System_Abbreviation"] = "MEP";
            }

            // Optionally: Serialize all aggregated parameters as JSON for diagnostics
            // var snapshot = combinedCluster.AggregatedSnapshot?.FlattenDistinct();
            // paramsToSet["CombinedSleeveData"] = ...;

            return paramsToSet;
        }
    }
}
