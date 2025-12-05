using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry
{
    /// <summary>
    /// Service for handling RCS (Relative Coordinate System) bounding box transformations for clustering.
    /// Follows SRP by centralizing RCS transformation logic.
    /// 
    /// Responsibilities:
    /// - Determine if RCS transformation is needed (walls/framing)
    /// - Transform WCS bounding boxes to RCS
    /// - Set RCS properties on ClashZone
    /// </summary>
    public class RcsBoundingBoxService
    {
        /// <summary>
        /// Process bounding box for a zone, applying RCS transformation if needed (walls/framing).
        /// </summary>
        /// <param name="zone">The clash zone</param>
        /// <param name="wcsBbox">Bounding box in World Coordinate System</param>
        /// <returns>True if RCS transformation was applied, false otherwise</returns>
        public bool ProcessBoundingBox(ClashZone zone, BoundingBoxXYZ wcsBbox)
        {
            if (zone == null || wcsBbox == null)
                return false;

            // Set WCS bounding box coordinates
            zone.SetSleeveBoundingBox(wcsBbox);

            // ✅ RCS BBOX: Transform to wall-aligned RCS for walls/framing
            // This eliminates rotation logic for walls - bounding boxes are already wall-aligned
            bool isWallHost = string.Equals(zone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(zone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
            bool isFramingHost = string.Equals(zone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);

            if ((isWallHost || isFramingHost) && zone.WallDirection != null && !zone.WallDirection.IsZeroLength())
            {
                try
                {
                    // ✅ CRITICAL: Use WallRcsTransformer to transform bounding box to RCS
                    var rcsBbox = WallRcsTransformer.TransformToRcs(wcsBbox, zone.WallDirection);
                    if (rcsBbox != null)
                    {
                        zone.SleeveBoundingBoxRCS_MinX = rcsBbox.Min.X;
                        zone.SleeveBoundingBoxRCS_MinY = rcsBbox.Min.Y;
                        zone.SleeveBoundingBoxRCS_MinZ = rcsBbox.Min.Z;
                        zone.SleeveBoundingBoxRCS_MaxX = rcsBbox.Max.X;
                        zone.SleeveBoundingBoxRCS_MaxY = rcsBbox.Max.Y;
                        zone.SleeveBoundingBoxRCS_MaxZ = rcsBbox.Max.Z;

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"[RcsBoundingBoxService] ✅ Applied RCS transformation for {zone.StructuralElementType} zone {zone.Id}");
                        }

                        return true;
                    }
                }
                catch (Exception rcsEx)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_errors.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [RcsBoundingBoxService] [RCS-TRANSFORM] Error transforming bbox to RCS for zone {zone.Id}: {rcsEx.Message}\n");
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Check if RCS transformation is needed for a zone.
        /// </summary>
        /// <param name="zone">The clash zone</param>
        /// <returns>True if RCS transformation should be applied</returns>
        public bool RequiresRcsTransformation(ClashZone zone)
        {
            if (zone == null)
                return false;

            bool isWallHost = string.Equals(zone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                              string.Equals(zone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
            bool isFramingHost = string.Equals(zone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);

            return (isWallHost || isFramingHost) && 
                   zone.WallDirection != null && 
                   !zone.WallDirection.IsZeroLength();
        }
    }
}

