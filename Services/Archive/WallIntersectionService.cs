using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Thin wrapper for wall intersection logic used by cable tray services.
    /// Delegates to the existing EfficientIntersectionService to avoid duplicating implementation.
    /// </summary>
    public static class WallIntersectionService
    {
        public static List<(ReferenceWithContext hit, XYZ direction, XYZ rayOrigin)> FindWallIntersections(
            MEPCurve mepElement,
            Line mepLine,
            View3D view3D,
            List<XYZ> testPoints,
            XYZ rayDirection)
        {
            return EfficientIntersectionService.FindWallIntersections(mepElement, mepLine, view3D, testPoints, rayDirection);
        }
    }
}
