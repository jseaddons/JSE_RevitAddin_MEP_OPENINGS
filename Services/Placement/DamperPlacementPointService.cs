using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Configuration;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ SRP COMPLIANCE: Dedicated service for damper placement point adjustment ONLY.
    /// Single Responsibility: Adjust damper placement point from centroid to wall centerline if offset.
    /// 
    /// This service ONLY handles:
    /// - Check if damper centroid is offset from wall centerline
    /// - If offset > threshold, move placement point to wall centerline
    /// - That's it - no other adjustments, no damper offsets, no complex logic
    /// </summary>
    public class DamperPlacementPointService
    {
        private readonly Document _doc;

        public DamperPlacementPointService(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// ✅ SRP: Adjust damper placement point from centroid to wall centerline if offset.
        /// This is the ONLY method that moves damper sleeves - all other classes delegate to this.
        /// </summary>
        /// <param name="zone">The clash zone (must be Duct Accessories category)</param>
        /// <param name="centroid">The damper centroid (original placement point)</param>
        /// <returns>Adjusted placement point (wall centerline if offset, otherwise original centroid)</returns>
        public XYZ AdjustDamperPlacementPoint(ClashZone zone, XYZ centroid)
        {
            if (zone == null)
                throw new ArgumentNullException(nameof(zone));
            
            if (centroid == null)
                throw new ArgumentNullException(nameof(centroid));

            // ✅ ONLY handle Duct Accessories (dampers)
            if (!string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                // Not a damper - return original point unchanged
                return centroid;
            }

            // ✅ ONLY handle walls and framing (floors don't need centerline adjustment)
            if (string.IsNullOrEmpty(zone.StructuralElementType) ||
                (!zone.StructuralElementType.Contains("Wall", StringComparison.OrdinalIgnoreCase) &&
                 !zone.StructuralElementType.Contains("Framing", StringComparison.OrdinalIgnoreCase)))
            {
                // Not a wall/framing - return original point unchanged
                return centroid;
            }

            // ✅ STEP 1: Check if wall centerline point is saved in database
            bool hasWallCenterline = (zone.WallCenterlinePointX != 0.0 || zone.WallCenterlinePointY != 0.0 || zone.WallCenterlinePointZ != 0.0);
            
            if (!hasWallCenterline)
            {
                // No saved centerline - return original point unchanged
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ⚠️ NO WALL CENTERLINE: Zone {zone.Id}, " +
                        $"WallCenterlinePointX={zone.WallCenterlinePointX:F6}ft, " +
                        $"WallCenterlinePointY={zone.WallCenterlinePointY:F6}ft, " +
                        $"WallCenterlinePointZ={zone.WallCenterlinePointZ:F6}ft - " +
                        $"KEEPING ORIGINAL CENTROID\n");
                }
                return centroid;
            }

            // ✅ STEP 2: Get saved wall centerline point
            XYZ wallCenterline = new XYZ(zone.WallCenterlinePointX, zone.WallCenterlinePointY, zone.WallCenterlinePointZ);
            
            // ✅ STEP 3: Intelligently merge coordinates based on wall orientation
            // For X-walls: Use wall centerline Y coordinate, keep centroid X and Z
            // For Y-walls: Use wall centerline X coordinate, keep centroid Y and Z
            // This avoids double-adjustment and ensures correct placement along wall length
            // ✅ CRITICAL: Always use saved centerline if available - no distance checks needed
            XYZ adjustedPoint = centroid;
            
            if (zone.HostOrientation == "X")
            {
                // X-wall: Wall runs along X axis, replace Y coordinate with wall centerline Y
                adjustedPoint = new XYZ(centroid.X, zone.WallCenterlinePointY, centroid.Z);
            }
            else if (zone.HostOrientation == "Y")
            {
                // Y-wall: Wall runs along Y axis, replace X coordinate with wall centerline X
                adjustedPoint = new XYZ(zone.WallCenterlinePointX, centroid.Y, centroid.Z);
            }
            else
            {
                // Unknown orientation - use full wall centerline point as fallback
                adjustedPoint = wallCenterline;
            }
            
            // ✅ DIAGNOSTIC: Log the adjustment (no distance threshold checks)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                XYZ difference = adjustedPoint - centroid;
                double diffX = Math.Abs(difference.X) * 304.8; // mm
                double diffY = Math.Abs(difference.Y) * 304.8; // mm
                double diffZ = Math.Abs(difference.Z) * 304.8; // mm
                double distance = centroid.DistanceTo(adjustedPoint);
                double distanceMm = distance * 304.8;
                
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] 📏 Zone {zone.Id}, " +
                    $"Orientation='{zone.HostOrientation}', " +
                    $"Centroid=({centroid.X:F6}ft, {centroid.Y:F6}ft, {centroid.Z:F6}ft), " +
                    $"WallCenterline=({wallCenterline.X:F6}ft, {wallCenterline.Y:F6}ft, {wallCenterline.Z:F6}ft), " +
                    $"Adjusted=({adjustedPoint.X:F6}ft, {adjustedPoint.Y:F6}ft, {adjustedPoint.Z:F6}ft), " +
                    $"Difference=({difference.X:F6}ft, {difference.Y:F6}ft, {difference.Z:F6}ft), " +
                    $"Distance={distanceMm:F1}mm, DiffX={diffX:F1}mm, DiffY={diffY:F1}mm, DiffZ={diffZ:F1}mm - " +
                    $"USING SAVED CENTERLINE (no distance check)\n");
            }
            
            // ✅ CRITICAL: Always use saved centerline if available - no distance threshold checks
            // The saved centerline was calculated during refresh when wall element was available
            // This ensures correct placement at wall centerline regardless of centroid offset
            return adjustedPoint;
        }
    }
}

