using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository
{
    /// <summary>
    /// Data access abstraction for combined clustering discovery (Phase 1-2).
    /// </summary>
    public interface ICombinedClusterRepository
    {
        IReadOnlyList<ClashZone> LoadClusteredZones(string filterName, IReadOnlyCollection<string> categories);

        Dictionary<int, Dictionary<string, string>> LoadSnapshotParameters(IEnumerable<int> sleeveInstanceIds);

        void SaveCombinedClusterMetadata(string filterName, IReadOnlyCollection<ClashZone> updatedZones);

        List<ClashZone> FindIndividualSleevesNearBoundingBox(Autodesk.Revit.DB.BoundingBoxXYZ boundingBox, string hostType, double level, double toleranceFt);
    }
}
