using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Interfaces
{
    /// <summary>
    /// ✅ SOLID SRP: Service responsible ONLY for calculating cluster placement points
    /// Single Responsibility: Calculate optimal placement point for cluster sleeve
    /// </summary>
    public interface IClusterPlacementPointService
    {
        /// <summary>
        /// Calculate cluster placement point from intersection points (centroid)
        /// </summary>
        /// <param name="cluster">List of sleeve data</param>
        /// <param name="xmlFilePath">Optional XML file path for data access</param>
        /// <returns>Placement point XYZ (centroid of intersection points)</returns>
        XYZ CalculatePlacementPoint(List<dynamic> cluster, string? xmlFilePath = null);
    }
}

