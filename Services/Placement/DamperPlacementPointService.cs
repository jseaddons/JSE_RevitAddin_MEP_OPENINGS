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
            
            // ✅ STEP 2A: Get host element bbox and center point for diagnostic logging
            BoundingBoxXYZ hostBbox = null;
            XYZ hostBboxCenter = null;
            try
            {
                if (zone.StructuralElementId != null && zone.StructuralElementId != ElementId.InvalidElementId)
                {
                    Element hostElement = _doc.GetElement(zone.StructuralElementId);
                    if (hostElement != null)
                    {
                        hostBbox = hostElement.get_BoundingBox(null);
                        if (hostBbox != null)
                        {
                            hostBboxCenter = new XYZ(
                                (hostBbox.Min.X + hostBbox.Max.X) / 2.0,
                                (hostBbox.Min.Y + hostBbox.Max.Y) / 2.0,
                                (hostBbox.Min.Z + hostBbox.Max.Z) / 2.0
                            );
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Ignore errors getting host bbox
            }
            
            // ✅ STEP 3: Intelligently merge coordinates based on wall orientation
            // For X-walls: Use wall centerline Y coordinate, keep centroid X and Z
            // For Y-walls: Use wall centerline X coordinate, keep centroid Y and Z
            // This avoids double-adjustment and ensures correct placement along wall length
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
            
            // ✅ STEP 4: Calculate the adjustment distance
            double distance = centroid.DistanceTo(adjustedPoint);
            
            // ✅ STEP 5: Calculate individual coordinate differences for detailed logging
            XYZ difference = adjustedPoint - centroid;
            double diffX = Math.Abs(difference.X) * 304.8; // mm
            double diffY = Math.Abs(difference.Y) * 304.8; // mm
            double diffZ = Math.Abs(difference.Z) * 304.8; // mm
            
            // ✅ COMPREHENSIVE DIAGNOSTIC: Log ALL points and calculations
            if (!DeploymentConfiguration.DeploymentMode)
            {
                string hostBboxStr = hostBbox != null 
                    ? $"Min=({hostBbox.Min.X:F6}ft, {hostBbox.Min.Y:F6}ft, {hostBbox.Min.Z:F6}ft), Max=({hostBbox.Max.X:F6}ft, {hostBbox.Max.Y:F6}ft, {hostBbox.Max.Z:F6}ft)" 
                    : "NULL";
                string hostCenterStr = hostBboxCenter != null 
                    ? $"({hostBboxCenter.X:F6}ft, {hostBboxCenter.Y:F6}ft, {hostBboxCenter.Z:F6}ft)" 
                    : "NULL";
                
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] 📏 Zone {zone.Id}, " +
                    $"Orientation='{zone.HostOrientation}', " +
                    $"Centroid=({centroid.X:F6}ft, {centroid.Y:F6}ft, {centroid.Z:F6}ft), " +
                    $"WallCenterline=({wallCenterline.X:F6}ft, {wallCenterline.Y:F6}ft, {wallCenterline.Z:F6}ft), " +
                    $"Adjusted=({adjustedPoint.X:F6}ft, {adjustedPoint.Y:F6}ft, {adjustedPoint.Z:F6}ft), " +
                    $"Difference=({difference.X:F6}ft, {difference.Y:F6}ft, {difference.Z:F6}ft), " +
                    $"Distance={distance:F6}ft ({distance * 304.8:F2}mm), " +
                    $"DiffX={diffX:F1}mm, DiffY={diffY:F1}mm, DiffZ={diffZ:F1}mm, " +
                    $"HostBbox={hostBboxStr}, HostBboxCenter={hostCenterStr}\n");
            }

            // ✅ STEP 6: Apply adjustment if distance is within acceptable range
            // MIN threshold: Prevents micro-adjustments (< 5mm = likely calculation precision)
            // MAX threshold: Prevents excessive adjustments (> 50mm = likely incorrect saved centerline)
            const double MIN_ADJUSTMENT_THRESHOLD_MM = 5.0; // 5mm minimum threshold
            const double MAX_ADJUSTMENT_THRESHOLD_MM = 50.0; // 50mm maximum threshold (saved centerline likely incorrect if > this)
            const double MIN_ADJUSTMENT_THRESHOLD_FT = MIN_ADJUSTMENT_THRESHOLD_MM / 304.8;
            const double MAX_ADJUSTMENT_THRESHOLD_FT = MAX_ADJUSTMENT_THRESHOLD_MM / 304.8;
            
            double distanceMm = distance * 304.8;
            
            if (distance > MIN_ADJUSTMENT_THRESHOLD_FT && distance <= MAX_ADJUSTMENT_THRESHOLD_FT)
            {
                // Centroid is offset from wall centerline within acceptable range - move to merged point
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ✅ ADJUSTING: Zone {zone.Id}, " +
                        $"Centroid offset {distanceMm:F1}mm from wall centerline ({MIN_ADJUSTMENT_THRESHOLD_MM}mm < {distanceMm:F1}mm <= {MAX_ADJUSTMENT_THRESHOLD_MM}mm) - " +
                        $"MOVING TO ADJUSTED POINT ({adjustedPoint.X:F6}ft, {adjustedPoint.Y:F6}ft, {adjustedPoint.Z:F6}ft)\n");
                }
                return adjustedPoint;
            }
            else if (distance <= MIN_ADJUSTMENT_THRESHOLD_FT)
            {
                // Centroid is already at wall centerline (within minimum threshold) - keep original
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ⏭️ SKIPPING ADJUSTMENT: Zone {zone.Id}, " +
                        $"Centroid offset {distanceMm:F1}mm < {MIN_ADJUSTMENT_THRESHOLD_MM}mm threshold - " +
                        $"KEEPING ORIGINAL CENTROID (difference too small to adjust)\n");
                }
                return centroid;
            }
            else if (distance > MAX_ADJUSTMENT_THRESHOLD_FT)
            {
                // ✅ FALLBACK: Difference too large - saved centerline is likely incorrect
                // ALWAYS recalculate wall centerline on-the-fly (like cluster sleeves - force placement between wall faces)
                // This is the MAIN reason to discard saved wall centerline - always use wall face method
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ⚠️ SAVED CENTERLINE UNRELIABLE: Zone {zone.Id}, " +
                        $"Centroid offset {distanceMm:F1}mm > {MAX_ADJUSTMENT_THRESHOLD_MM}mm threshold - " +
                        $"FORCING RECALCULATION using wall face method (like cluster sleeves)\n");
                }
                
                // ✅ FORCE RECALCULATION: Recalculate wall centerline using current centroid as starting point
                // This ALWAYS forces placement between wall faces (similar to cluster sleeve logic)
                // NO FALLBACK TO CENTROID - this is the main reason to discard saved centerline
                try
                {
                    // Retrieve wall element from zone
                    Element wallElement = null;
                    if (zone.StructuralElementId != null && zone.StructuralElementId != ElementId.InvalidElementId)
                    {
                        wallElement = _doc.GetElement(zone.StructuralElementId);
                    }
                    
                    if (wallElement != null && (wallElement is Wall || 
                        (wallElement is FamilyInstance framing && 
                         framing.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)))
                    {
                        // ✅ RECALCULATE: Use WallCenterlineHelper to calculate centerline from current centroid
                        // This projects the centroid onto the wall centerline (between two wall faces)
                        XYZ recalculatedCenterline;
                        if (wallElement is Wall wall)
                        {
                            recalculatedCenterline = WallCenterlineHelper.GetWallCenterlinePoint(wall, centroid, _doc);
                        }
                        else if (wallElement is FamilyInstance framingInstance)
                        {
                            recalculatedCenterline = WallCenterlineHelper.GetStructuralFramingCenterlinePoint(framingInstance, centroid, _doc);
                        }
                        else
                        {
                            // Should not happen, but if it does, use centroid
                            recalculatedCenterline = centroid;
                        }
                        
                        // ✅ MERGE COORDINATES: Apply same logic as saved centerline (orientation-based)
                        XYZ recalculatedAdjustedPoint = centroid;
                        if (zone.HostOrientation == "X")
                        {
                            // X-wall: Use recalculated Y coordinate, keep centroid X and Z
                            recalculatedAdjustedPoint = new XYZ(centroid.X, recalculatedCenterline.Y, centroid.Z);
                        }
                        else if (zone.HostOrientation == "Y")
                        {
                            // Y-wall: Use recalculated X coordinate, keep centroid Y and Z
                            recalculatedAdjustedPoint = new XYZ(recalculatedCenterline.X, centroid.Y, centroid.Z);
                        }
                        else
                        {
                            // Unknown orientation - use full recalculated centerline
                            recalculatedAdjustedPoint = recalculatedCenterline;
                        }
                        
                        double recalculatedDistance = centroid.DistanceTo(recalculatedAdjustedPoint);
                        double recalculatedDistanceMm = recalculatedDistance * 304.8;
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ✅ WALL FACE METHOD: Zone {zone.Id}, " +
                                $"Recalculated centerline=({recalculatedCenterline.X:F6}ft, {recalculatedCenterline.Y:F6}ft, {recalculatedCenterline.Z:F6}ft), " +
                                $"Adjusted point=({recalculatedAdjustedPoint.X:F6}ft, {recalculatedAdjustedPoint.Y:F6}ft, {recalculatedAdjustedPoint.Z:F6}ft), " +
                                $"Distance={recalculatedDistanceMm:F1}mm - " +
                                $"FORCING PLACEMENT BETWEEN WALL FACES (no fallback to centroid)\n");
                        }
                        
                        // ✅ ALWAYS USE RECALCULATED: Force placement between wall faces (like cluster sleeves)
                        // This is the main reason to discard saved centerline - always use wall face method
                        return recalculatedAdjustedPoint;
                    }
                    else
                    {
                        // Wall element not found or invalid - only fallback to centroid if we can't get wall
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("placement_debug.log",
                                $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ⚠️ WALL ELEMENT NOT FOUND: Zone {zone.Id}, " +
                                $"StructuralElementId={zone.StructuralElementId?.IntegerValue ?? -1} - " +
                                $"Cannot recalculate centerline (wall not found), KEEPING ORIGINAL CENTROID\n");
                        }
                        return centroid;
                    }
                }
                catch (Exception ex)
                {
                    // Error during recalculation - only fallback to centroid if calculation fails
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ❌ RECALCULATION ERROR: Zone {zone.Id}, " +
                            $"Error recalculating wall centerline: {ex.Message} - " +
                            $"KEEPING ORIGINAL CENTROID (calculation failed)\n");
                    }
                    return centroid;
                }
            }
            else
            {
                // Centroid is already at wall centerline (within minimum threshold) - keep original
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ✅ NO ADJUSTMENT: Zone {zone.Id}, " +
                        $"Centroid offset {distanceMm:F1}mm < {MIN_ADJUSTMENT_THRESHOLD_MM}mm threshold - " +
                        $"KEEPING ORIGINAL CENTROID (difference too small to adjust)\n");
                }
                return centroid;
            }
        }
    }
}

