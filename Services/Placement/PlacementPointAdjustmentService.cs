using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Configuration;
using JSE_RevitAddin_MEP_OPENINGS.Services; // For ElementRetrievalService

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ SRP COMPLIANCE: Service responsible for adjusting placement points for NON-DAMPER MEP elements.
    /// Handles:
    /// - Wall/Framing centerline adjustment for Ducts, Pipes, Cable Trays
    /// 
    /// ✅ CRITICAL: Dampers (Duct Accessories) are handled by DamperPlacementPointService - this service delegates to it.
    /// This service maintains SRP by separating placement point calculation from placement orchestration.
    /// 
    /// ✅ 28 FEATURES COMPLIANCE: Implements all relevant features from Comprehensive Architecture:
    /// - SOLID Architecture (SRP, DIP)
    /// - Diagnostic Logging (with DeploymentMode support)
    /// - Exception Handling (graceful degradation)
    /// - Safe File Operations (SafeFileLogger)
    /// - Element Validation (IsValidObject checks)
    /// - Performance Monitoring (optional, via PlacementPerformanceMonitor)
    /// - Version Versatility (standard Revit API)
    /// </summary>
    public class PlacementPointAdjustmentService
    {
        private readonly Document _doc;
        private readonly PlacementPerformanceMonitor? _performanceMonitor;

        /// <summary>
        /// ✅ DIP COMPLIANCE: Constructor accepts optional performance monitor for dependency injection.
        /// </summary>
        public PlacementPointAdjustmentService(Document doc, PlacementPerformanceMonitor? performanceMonitor = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _performanceMonitor = performanceMonitor;
        }

        /// <summary>
        /// ✅ SRP: Adjusts placement point based on host element type.
        /// This method handles placement point adjustment for ALL MEP elements EXCEPT dampers.
        /// 
        /// ✅ CRITICAL: Dampers (Duct Accessories) are handled by DamperPlacementPointService (SRP compliance).
        /// This service delegates to DamperPlacementPointService for dampers to maintain single responsibility.
        /// 
        /// ✅ 28 FEATURES: Implements performance monitoring, element validation, and diagnostic logging.
        /// </summary>
        /// <param name="zone">The clash zone containing host element information</param>
        /// <param name="originalPlacementPoint">The original intersection point (from refresh, at MEP element/wall intersection)</param>
        /// <param name="damperOffset">DEPRECATED: No longer used - damper placement is handled by DamperPlacementPointService</param>
        /// <returns>Adjusted placement point (at wall centerline for non-dampers)</returns>
        public XYZ AdjustPlacementPoint(ClashZone zone, XYZ originalPlacementPoint, XYZ damperOffset = null)
        {
            // ✅ PERFORMANCE MONITORING: Track operation time (Feature #11 from 28 features)
            using (var tracker = _performanceMonitor?.TrackOperation("Adjust Placement Point"))
            {
                if (zone == null)
                    throw new ArgumentNullException(nameof(zone));
                
                if (originalPlacementPoint == null)
                    throw new ArgumentNullException(nameof(originalPlacementPoint));

                // ✅ SRP COMPLIANCE: Delegate damper placement to dedicated service
                // Dampers (Duct Accessories) have their own placement logic - don't mix with other MEP elements
                if (string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                {
                    // ✅ DELEGATE TO DEDICATED SERVICE: DamperPlacementPointService handles ALL damper placement adjustments
                    // This maintains SRP - one class, one responsibility
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] 🔄 DELEGATING to DamperPlacementPointService: Zone {zone.Id}, " +
                            $"OriginalPoint=({originalPlacementPoint.X:F6}ft, {originalPlacementPoint.Y:F6}ft, {originalPlacementPoint.Z:F6}ft), " +
                            $"WallCenterline=({zone.WallCenterlinePointX:F6}ft, {zone.WallCenterlinePointY:F6}ft, {zone.WallCenterlinePointZ:F6}ft)\n");
                    }
                    
                    var damperService = new DamperPlacementPointService(_doc);
                    XYZ damperAdjustedPoint = damperService.AdjustDamperPlacementPoint(zone, originalPlacementPoint);
                    
                    // ✅ PERFORMANCE MONITORING: Record item count
                    tracker?.SetItemCount(1);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ RETURNED from DamperPlacementPointService: Zone {zone.Id}, " +
                            $"AdjustedPoint=({damperAdjustedPoint.X:F6}ft, {damperAdjustedPoint.Y:F6}ft, {damperAdjustedPoint.Z:F6}ft)\n");
                    }
                    
                    return damperAdjustedPoint;
                }

                // ✅ NON-DAMPER ELEMENTS: Adjust for wall/framing centerline only
                // This applies to Ducts, Pipes, Cable Trays - NOT dampers
                XYZ adjustedPoint = AdjustForHostCenterline(zone, originalPlacementPoint);

                // ✅ PERFORMANCE MONITORING: Record item count
                tracker?.SetItemCount(1);

                return adjustedPoint;
            }
        }

        /// <summary>
        /// ✅ SRP: Adjusts placement point to host element centerline for walls and structural framing.
        /// This ensures sleeves are placed at the center of walls/framing, not on the face or at MEP element center.
        /// 
        /// ✅ CRITICAL: This method is ONLY called for NON-DAMPER MEP elements (Ducts, Pipes, Cable Trays).
        /// Dampers (Duct Accessories) are handled by DamperPlacementPointService - this method should NEVER be called for dampers.
        /// </summary>
        private XYZ AdjustForHostCenterline(ClashZone zone, XYZ placementPoint)
        {
            // Check if this is a wall or structural framing element
            if (string.IsNullOrEmpty(zone.StructuralElementType) ||
                (!zone.StructuralElementType.Contains("Wall", StringComparison.OrdinalIgnoreCase) &&
                 !zone.StructuralElementType.Contains("Framing", StringComparison.OrdinalIgnoreCase)))
            {
                // Not a wall or framing, no adjustment needed
                return placementPoint;
            }

            // ✅ PERFORMANCE: Use pre-calculated wall centerline point from database (enables multi-threading)
            // ✅ CRITICAL: Wall centerline point is calculated during refresh when wall element is available
            // This avoids Revit API calls during placement, enabling parallel processing
            // ✅ CRITICAL FIX: Check X/Y/Z values directly (not just computed property) to handle cases where values are saved but property returns zero
            bool hasWallCenterline = (zone.WallCenterlinePointX != 0.0 || zone.WallCenterlinePointY != 0.0 || zone.WallCenterlinePointZ != 0.0);
            
            // ✅ DIAGNOSTIC: Log wall centerline values for debugging "half in and out" issue
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] 🔍 WALL CENTERLINE CHECK: Zone {zone.Id}, " +
                    $"HasWallCenterline={hasWallCenterline}, " +
                    $"WallCenterlinePointX={zone.WallCenterlinePointX:F6}ft ({zone.WallCenterlinePointX * 304.8:F1}mm), " +
                    $"WallCenterlinePointY={zone.WallCenterlinePointY:F6}ft ({zone.WallCenterlinePointY * 304.8:F1}mm), " +
                    $"WallCenterlinePointZ={zone.WallCenterlinePointZ:F6}ft ({zone.WallCenterlinePointZ * 304.8:F1}mm), " +
                    $"PlacementPoint (Centroid)=({placementPoint.X:F6}ft, {placementPoint.Y:F6}ft, {placementPoint.Z:F6}ft)\n");
            }
            
            if (hasWallCenterline)
            {
                // ✅ USE SAVED WALL CENTERLINE POINT: Pre-calculated during refresh (no Revit API calls needed)
                // This enables multi-threading because it's just data access, not Revit API calls
                // ✅ CRITICAL: Construct XYZ from saved X/Y/Z values (more reliable than computed property)
                XYZ savedCenterlinePoint = new XYZ(zone.WallCenterlinePointX, zone.WallCenterlinePointY, zone.WallCenterlinePointZ);
                
                // ✅ CALCULATE OFFSET: If placement point is offset from wall centerline, use saved centerline
                // ✅ NOTE: This is for NON-DAMPER elements only - dampers are handled by DamperPlacementPointService
                // ✅ CRITICAL FIX: Only adjust if placement point is SIGNIFICANTLY offset (not just > 1e-6)
                // This prevents moving sleeves that are already at wall centerline to the wrong position
                // Threshold: 1mm (0.00328 feet) - only adjust if offset is more than 1mm
                const double MIN_ADJUSTMENT_THRESHOLD = 0.00328; // 1mm in feet
                double distance = placementPoint.DistanceTo(savedCenterlinePoint);
                
                // ✅ DIAGNOSTIC: Always log the distance calculation for debugging
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    XYZ difference = savedCenterlinePoint - placementPoint;
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] 📏 DISTANCE CALCULATION: Zone {zone.Id}, " +
                        $"PlacementPoint=({placementPoint.X:F6}ft, {placementPoint.Y:F6}ft, {placementPoint.Z:F6}ft), " +
                        $"WallCenterline=({savedCenterlinePoint.X:F6}ft, {savedCenterlinePoint.Y:F6}ft, {savedCenterlinePoint.Z:F6}ft), " +
                        $"Difference=({difference.X:F6}ft, {difference.Y:F6}ft, {difference.Z:F6}ft), " +
                        $"Distance={distance:F6}ft ({distance * 304.8:F2}mm), " +
                        $"Threshold={MIN_ADJUSTMENT_THRESHOLD:F6}ft (1.0mm)\n");
                }
                
                // ✅ CRITICAL FIX: Only adjust if distance is significant (> 1mm)
                // This prevents moving sleeves that are already at wall centerline
                if (distance > MIN_ADJUSTMENT_THRESHOLD)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ USING SAVED WALL CENTERLINE: Zone {zone.Id}, " +
                            $"Centroid=({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3}), " +
                            $"WallCenterline=({savedCenterlinePoint.X:F3}, {savedCenterlinePoint.Y:F3}, {savedCenterlinePoint.Z:F3}), " +
                            $"Distance={distance * 304.8:F1}mm > 1mm threshold - ADJUSTING TO WALL CENTERLINE\n");
                    }
                    return savedCenterlinePoint; // Use pre-calculated wall centerline point
                }
                else
                {
                    // Centroid is already at wall centerline (or very close, < 1mm), no adjustment needed
                    // ✅ CRITICAL: Don't move sleeves that are already centered (prevents "half in and out" issue)
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ CENTROID ALREADY AT WALL CENTERLINE: Zone {zone.Id}, " +
                            $"No adjustment needed (distance={distance * 304.8:F1}mm < 1mm threshold) - KEEPING ORIGINAL PLACEMENT POINT\n");
                    }
                    return placementPoint;
                }
            }
            else if (!DeploymentConfiguration.DeploymentMode)
            {
                // ✅ DIAGNOSTIC: Log when wall centerline is NOT available (this causes "half in and out" issue)
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️⚠️⚠️ NO WALL CENTERLINE SAVED: Zone {zone.Id}, " +
                    $"WallCenterlinePointX={zone.WallCenterlinePointX:F6}ft, " +
                    $"WallCenterlinePointY={zone.WallCenterlinePointY:F6}ft, " +
                    $"WallCenterlinePointZ={zone.WallCenterlinePointZ:F6}ft - " +
                    $"Will use FALLBACK calculation (may cause 'half in and out' issue)\n");
            }

            // ✅ FALLBACK: If wall centerline point is not saved (old data or calculation failed during refresh)
            // Calculate it now (requires Revit API calls, cannot be parallelized)
            // This should rarely happen if refresh was successful
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️ FALLBACK: Wall centerline point not saved for Zone {zone.Id}, calculating now (requires Revit API, cannot be parallelized)\n");
            }

            // Check if we have a valid host element ID
            if (zone.StructuralElementId == null || zone.StructuralElementId.IntegerValue <= 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️ NO HOST ELEMENT ID: Zone {zone.Id}, HostType={zone.StructuralElementType}, Cannot adjust to centerline\n");
                }
                return placementPoint;
            }

            try
            {
                // ✅ CRITICAL FIX: Use ElementRetrievalService to find element in active OR linked documents
                // Walls/framing can be in linked documents, so we need to search both active and linked docs
                Element hostElement = ElementRetrievalService.GetElementFromDocumentOrLinked(_doc, zone.StructuralElementId);
                if (hostElement == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️ HOST ELEMENT NOT FOUND: Zone {zone.Id}, HostId={zone.StructuralElementId.IntegerValue}, Cannot adjust to centerline (searched active and linked documents)\n");
                    }
                    return placementPoint;
                }

                // ✅ ELEMENT VALIDATION: Check if element is still valid (Feature #4 from Transaction & Safety Features)
                // This prevents "referenced object is not valid" errors
                if (!hostElement.IsValidObject)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️ HOST ELEMENT INVALID: Zone {zone.Id}, HostId={zone.StructuralElementId.IntegerValue}, Element is not valid (may have been deleted), Cannot adjust to centerline\n");
                    }
                    return placementPoint;
                }

                // ✅ DELEGATE TO HELPER: Use WallCenterlineHelper to adjust placement point to centerline
                // This maintains SRP by delegating geometric calculation to a specialized helper
                // ✅ CRITICAL FIX: Pass host document so helper can find transform for linked document walls
                // ✅ NOTE: This method is ONLY for non-damper MEP elements - dampers are handled by DamperPlacementPointService
                
                XYZ centerlinePoint = WallCenterlineHelper.GetElementCenterlinePoint(hostElement, placementPoint, _doc);

                if (centerlinePoint != null && centerlinePoint != placementPoint)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ HOST CENTERLINE ADJUSTED: Zone {zone.Id}, HostType={zone.StructuralElementType}, HostId={zone.StructuralElementId.IntegerValue}, Original=({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3}), Adjusted=({centerlinePoint.X:F3}, {centerlinePoint.Y:F3}, {centerlinePoint.Z:F3})\n");
                    }
                    return centerlinePoint;
                }
            }
            catch (Exception ex)
            {
                // Log error but continue with original placement point
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️ CENTERLINE ADJUSTMENT FAILED: Zone {zone.Id}, Error={ex.Message}, Using original placement point\n");
                }
            }

            return placementPoint;
        }
    }
}

