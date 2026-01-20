using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase3And4.Interfaces
{
    public interface IParameterAggregatorService
    {
        Dictionary<string, object> CreateCombinedSleeveParameterSet(
            CombinedClusterCandidate combinedCluster, 
            BoundingBoxXYZ combinedBoundingBox);
    }
}
