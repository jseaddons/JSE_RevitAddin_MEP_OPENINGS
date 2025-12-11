using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement
{
    /// <summary>
    /// Phase 5: Interface for cluster sleeve placement and parameter management.
    /// Handles creation, sizing, metadata, and family loading for cluster sleeves.
    /// </summary>
    public interface IClusterPlacementService
    {
        /// <summary>
        /// Place a cluster sleeve in the document.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="cluster">List of sleeves in the cluster (dynamic type)</param>
        /// <param name="groupKey">Group key (host type, system type, orientation)</param>
        /// <param name="targetCategory">Target MEP category</param>
        /// <param name="placementPoint">Placement point in world coordinates</param>
        /// <param name="width">Cluster width in internal units</param>
        /// <param name="height">Cluster height in internal units</param>
        /// <param name="depth">Cluster depth in internal units</param>
        /// <param name="rotationAngle">Rotation angle in radians</param>
        /// <param name="xmlFilePath">Optional XML file path for data lookup</param>
        /// <param name="placedClusterSleeve">Output: The placed cluster sleeve instance</param>
        /// <param name="capturedClusterSleeveId">Output: The captured cluster sleeve ID</param>
        /// <returns>True if placement succeeded, false otherwise</returns>
        /// <summary>
        /// Place a cluster sleeve in the document.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="cluster">List of sleeves in the cluster (dynamic type)</param>
        /// <param name="groupKey">Group key (host type, system type, orientation)</param>
        /// <param name="targetCategory">Target MEP category</param>
        /// <param name="placementPoint">Placement point in world coordinates</param>
        /// <param name="width">Cluster width in internal units</param>
        /// <param name="height">Cluster height in internal units</param>
        /// <param name="depth">Cluster depth in internal units</param>
        /// <param name="rotationAngle">Rotation angle in radians</param>
        /// <param name="xmlFilePath">Optional XML file path for data lookup</param>
        /// <param name="placedClusterSleeve">Output: The placed cluster sleeve instance</param>
        /// <param name="capturedClusterSleeveId">Output: The captured cluster sleeve ID</param>
        /// <param name="deferredParameters">Optional dictionary for batch parameter updates</param>
        /// <returns>True if placement succeeded, false otherwise</returns>
        bool PlaceClusterSleeve(
            Document doc,
            List<dynamic> cluster,
            SleeveGroupKey groupKey,
            string targetCategory,
            XYZ placementPoint,
            double width,
            double height,
            double depth,
            double rotationAngle,
            string? xmlFilePath,
            out FamilyInstance? placedClusterSleeve,
            out int? capturedClusterSleeveId,
            out XYZ? actualPlacementPoint,
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null);

        /// <summary>
        /// Set size parameters (Width, Height, Depth) on a cluster sleeve.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="clusterSleeve">The cluster sleeve instance</param>
        /// <param name="cluster">List of sleeves in the cluster (for dimension validation)</param>
        /// <param name="groupKey">Group key (host type, system type, orientation)</param>
        /// <param name="width">Cluster width in internal units</param>
        /// <param name="height">Cluster height in internal units</param>
        /// <param name="depth">Cluster depth in internal units</param>
        /// <param name="shouldSwapDimensions">Whether to swap dimensions for wall/framing hosts</param>
        /// <param name="deferredParameters">Optional dictionary for batch parameter updates</param>
        void SetSizeParameters(
            Document doc,
            FamilyInstance clusterSleeve,
            List<dynamic> cluster,
            SleeveGroupKey groupKey,
            double width,
            double height,
            double depth,
            bool shouldSwapDimensions = false,
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null);

        /// <summary>
        /// Set metadata parameters on a cluster sleeve (MEP_Category, Filter Name, Sleeve Instance ID, etc.).
        /// </summary>
        /// <param name="clusterSleeve">The cluster sleeve instance</param>
        /// <param name="category">MEP category name</param>
        /// <param name="filterName">Optional filter name (if null, will be looked up by category)</param>
        /// <param name="deferredParameters">Optional dictionary for batch parameter updates</param>
        void SetMetadata(
            FamilyInstance clusterSleeve,
            string category,
            string? filterName = null,
            Dictionary<ElementId, Dictionary<string, object>>? deferredParameters = null);

        /// <summary>
        /// Get reference level for cluster sleeve placement.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="sleeve">Reference sleeve from cluster (dynamic type)</param>
        /// <returns>Reference level, or null if not found</returns>
        Level? GetReferenceLevel(Document doc, dynamic sleeve);

        /// <summary>
        /// Load a universal family into the document if not already present.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="familyName">Name of the family to load (e.g., "RectangularOpeningOnWall")</param>
        /// <returns>True if family is available (loaded or already present), false otherwise</returns>
        bool LoadFamily(Document doc, string familyName);

        /// <summary>
        /// Get or cache MEP element by ElementId.
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="elementId">MEP element ID</param>
        /// <returns>MEP element, or null if not found</returns>
        Element? GetMepElement(Document doc, ElementId elementId);

        /// <summary>
        /// Get or cache bounding box for a sleeve instance.
        /// </summary>
        /// <param name="sleeve">Sleeve instance</param>
        /// <returns>Bounding box, or null if not available</returns>
        BoundingBoxXYZ? GetBoundingBox(FamilyInstance sleeve);

        /// <summary>
        /// Get or cache parameter for a sleeve instance.
        /// </summary>
        /// <param name="sleeve">Sleeve instance</param>
        /// <param name="parameterName">Parameter name</param>
        /// <returns>Parameter, or null if not found</returns>
        Parameter? GetParameter(FamilyInstance sleeve, string parameterName);

        /// <summary>
        /// Clear all caches (useful for memory management or testing).
        /// </summary>
        void ClearCaches();

        /// <summary>
        /// Get the MEP element cache (for pre-population).
        /// </summary>
        Dictionary<ElementId, Element> MepElementCache { get; }

        /// <summary>
        /// Get the bounding box cache (for pre-population).
        /// </summary>
        Dictionary<FamilyInstance, BoundingBoxXYZ> BoundingBoxCache { get; }

        /// <summary>
        /// Get the parameter cache (for pre-population).
        /// </summary>
        Dictionary<FamilyInstance, Dictionary<string, Parameter>> ParameterCache { get; }
    }
}

