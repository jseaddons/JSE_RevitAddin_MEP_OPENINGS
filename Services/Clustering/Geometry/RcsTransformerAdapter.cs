using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry
{
    /// <summary>
    /// ✅ SOLID DIP: Adapter that wraps static WallRcsTransformer to implement IRcsTransformer
    /// This allows dependency injection while maintaining backward compatibility
    /// </summary>
    public class RcsTransformerAdapter : IRcsTransformer
    {
        public BoundingBoxXYZ? TransformToRcs(BoundingBoxXYZ? wcsBbox, XYZ? wallDirection)
        {
            return WallRcsTransformer.TransformToRcs(wcsBbox, wallDirection);
        }
        
        public XYZ? TransformToWcs(XYZ? rcsPoint, XYZ? wallDirection, XYZ? wallOrigin)
        {
            return WallRcsTransformer.TransformToWcs(rcsPoint, wallDirection, wallOrigin);
        }
    }
}

