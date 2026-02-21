using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Repository
{
    /// <summary>
    /// Data access abstraction for combined clustering discovery (Phase 1-2).
    /// </summary>
    public interface ICombinedClusterRepository
    {
        IReadOnlyList<ClashZone> LoadClusteredZones(Autodesk.Revit.DB.Document doc, string filterName, IReadOnlyCollection<string> categories);

        Dictionary<long, Dictionary<string, string>> LoadSnapshotParameters(IEnumerable<long> sleeveInstanceIds);

        void SaveCombinedClusterMetadata(string filterName, IReadOnlyCollection<ClashZone> updatedZones);

        List<ClashZone> FindIndividualSleevesNearBoundingBox(Autodesk.Revit.DB.BoundingBoxXYZ boundingBox, string hostType, double level, double toleranceFt);
    }
}
