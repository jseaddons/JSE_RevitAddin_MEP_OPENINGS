using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement
{
    /// <summary>
    /// Interface for calculating the authoritative placement coordinate for clusters.
    /// Handles category-specific logic (e.g., Dampers vs Pipes) and host-specific logic (Walls vs Floors).
    public interface IClusterPlacementCalculationService
    {
        /// <summary>
        /// Calculates the placement point based on intersection points, element categories, and calculated bounding box.
        /// </summary>
        /// <param name="cluster">List of sleeve data.</param>
        /// <param name="bboxWidth">Calculated cluster width.</param>
        /// <param name="bboxHeight">Calculated cluster height.</param>
        /// <param name="bboxDepth">Calculated cluster depth.</param>
        /// <param name="rotatedMin">Optional rotated min point (for geometric center fallback).</param>
        /// <param name="rotatedMax">Optional rotated max point (for geometric center fallback).</param>
        /// <param name="xmlFilePath">Optional XML file path.</param>
        /// <returns>Authoritative XYZ coordinate for placement.</returns>
        XYZ CalculatePlacementPoint(
            List<dynamic> cluster, 
            double bboxWidth, 
            double bboxHeight, 
            double bboxDepth,
            XYZ? rotatedMin = null,
            XYZ? rotatedMax = null,
            XYZ? explicitCenter = null,
            string? xmlFilePath = null);
    }
}
