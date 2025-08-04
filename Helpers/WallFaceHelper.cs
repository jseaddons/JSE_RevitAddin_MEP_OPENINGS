using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    public static class WallFaceHelper
    {
        // Returns the normal of the exterior face of the wall at a given point
        public static XYZ GetWallExteriorNormal(Wall wall, XYZ nearPoint)
        {
            DebugLogger.Log($"[WALL-NORMAL-DEBUG] ===== WALL NORMAL CALCULATION =====");
            if (wall == null)
            {
                DebugLogger.Log("[WALL-NORMAL-DEBUG] Wall is null, returning BasisZ");
                return XYZ.BasisZ;
            }
            try
            {
                XYZ wallOrientation = wall.Orientation.Normalize();
                DebugLogger.Log($"[WALL-NORMAL-DEBUG] Wall orientation: {wallOrientation}");
                
                Options options = new Options { ComputeReferences = true, IncludeNonVisibleObjects = false };
                GeometryElement geomElem = wall.get_Geometry(options);
                Face? bestFace = null;
                double bestDot = double.MinValue;
                
                foreach (GeometryObject obj in geomElem)
                {
                    if (obj is Solid solid && solid.Volume > 0)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            // Only consider planar faces
                            if (face is not PlanarFace pf) continue;
                            // Compare face normal to wall orientation
                            double dot = pf.FaceNormal.Normalize().DotProduct(wallOrientation);
                            if (dot > bestDot)
                            {
                                bestDot = dot;
                                bestFace = pf;
                            }
                        }
                    }
                }
                if (bestFace is PlanarFace pfBest)
                {
                    XYZ normal = pfBest.FaceNormal.Normalize();
                    DebugLogger.Log($"[WALL-NORMAL-DEBUG] Best planar face normal: {normal}");
                    // If normal is nearly vertical, fallback to wall orientation
                    if (Math.Abs(normal.Z) > 0.7) // nearly vertical
                    {
                        DebugLogger.Log("[WALL-NORMAL-DEBUG] Normal is vertical, using wall orientation instead");
                        // Return a perpendicular in XY plane
                        XYZ perp = new XYZ(-wallOrientation.Y, wallOrientation.X, 0).Normalize();
                        DebugLogger.Log($"[WALL-NORMAL-DEBUG] Calculated XY perpendicular: {perp}");
                        DebugLogger.Log($"[WALL-NORMAL-DEBUG] ===== END WALL NORMAL CALCULATION =====");
                        return perp;
                    }
                    DebugLogger.Log($"[WALL-NORMAL-DEBUG] Using planar face normal: {normal}");
                    DebugLogger.Log($"[WALL-NORMAL-DEBUG] ===== END WALL NORMAL CALCULATION =====");
                    return normal;
                }
                else
                {
                    DebugLogger.Log("[WALL-NORMAL-DEBUG] No suitable planar face found, using wall orientation");
                    XYZ perp = new XYZ(-wallOrientation.Y, wallOrientation.X, 0).Normalize();
                    DebugLogger.Log($"[WALL-NORMAL-DEBUG] Calculated XY perpendicular: {perp}");
                    DebugLogger.Log($"[WALL-NORMAL-DEBUG] ===== END WALL NORMAL CALCULATION =====");
                    return perp;
                }
            }
            catch (System.Exception ex)
            {
                DebugLogger.Log($"[WALL-NORMAL-DEBUG] Exception: {ex.Message}");
            }
            // Final fallback
            DebugLogger.Log("[WALL-NORMAL-DEBUG] Final fallback, using BasisY");
            DebugLogger.Log($"[WALL-NORMAL-DEBUG] ===== END WALL NORMAL CALCULATION =====");
            return XYZ.BasisY;
        }
    }
}
