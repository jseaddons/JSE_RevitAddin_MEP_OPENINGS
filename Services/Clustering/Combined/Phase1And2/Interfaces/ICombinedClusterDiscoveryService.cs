using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Combined.Phase1And2.Interfaces
{
    /// <summary>
    /// Discovers cluster sleeves eligible for combined clustering (Phase 1-2, DB/CPU only).
    /// </summary>
    public interface ICombinedClusterDiscoveryService
    {
        IReadOnlyList<ClusterSleeveInfo> Discover(Autodesk.Revit.UI.UIDocument uiDoc, string filterName, IReadOnlyCollection<string> categories, Autodesk.Revit.DB.BoundingBoxXYZ? sectionBox = null);
    }
}
