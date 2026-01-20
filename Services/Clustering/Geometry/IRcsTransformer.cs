using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry
{
    /// <summary>
    /// ✅ SOLID DIP: Interface for RCS transformation operations
    /// Allows dependency injection instead of static class dependency
    /// </summary>
    public interface IRcsTransformer
    {
        /// <summary>
        /// Transform WCS bounding box to wall-aligned RCS
        /// </summary>
        BoundingBoxXYZ? TransformToRcs(BoundingBoxXYZ? wcsBbox, XYZ? wallDirection);
        
        /// <summary>
        /// Transform RCS point back to WCS
        /// </summary>
        XYZ? TransformToWcs(XYZ? rcsPoint, XYZ? wallDirection, XYZ? wallOrigin);
    }
}

