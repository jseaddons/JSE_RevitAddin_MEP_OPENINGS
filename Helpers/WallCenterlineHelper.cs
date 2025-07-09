using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class WallCenterlineHelper
    {
        // Returns the centerline point of the wall at a given intersection point, using robust exterior normal
        public static XYZ GetWallCenterlinePoint(Wall wall, XYZ intersectionPoint)
        {
            DebugLogger.Log($"[CENTERLINE-DEBUG] Input intersectionPoint: {intersectionPoint}");
            if (wall == null)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Wall is null, returning input point");
                return intersectionPoint;
            }
            try
            {
                LocationCurve locCurve = wall.Location as LocationCurve;
                if (locCurve?.Curve is Line wallLine)
                {
                    XYZ projected = wallLine.Project(intersectionPoint).XYZPoint;
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Projected point on wall line: {projected}");
                    XYZ wallNormal = WallFaceHelper.GetWallExteriorNormal(wall, projected);
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Wall normal at projected point: {wallNormal}");
                    double halfWidth = wall.Width / 2.0;
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Wall width: {wall.Width}, halfWidth: {halfWidth}");
                    XYZ centerlinePoint = projected + wallNormal * (-halfWidth);
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Calculated centerline point: {centerlinePoint}");
                    return centerlinePoint;
                }
                else
                {
                    DebugLogger.Log($"[CENTERLINE-DEBUG] Wall location is not a line, returning input point");
                }
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[CENTERLINE-DEBUG] Exception: {ex.Message}");
            }
            return intersectionPoint;
        }
    }
}
