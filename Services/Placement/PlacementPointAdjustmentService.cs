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
        private readonly bool _isForceDetectionMode;
        private readonly PlacementPerformanceMonitor? _performanceMonitor;
        private readonly HostPropertyCache? _hostPropertyCache; // ✅ Phase 2: Host Cache

        /// <summary>
        /// ✅ DIP COMPLIANCE: Constructor accepts optional performance monitor for dependency injection.
        /// </summary>
        public PlacementPointAdjustmentService(
            Document doc, 
            PlacementPerformanceMonitor? performanceMonitor = null, 
            bool isForceDetectionMode = false,
            HostPropertyCache? hostPropertyCache = null) // ✅ Phase 2: Host Cache Injection
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _performanceMonitor = performanceMonitor;
            _isForceDetectionMode = isForceDetectionMode;
            _hostPropertyCache = hostPropertyCache;
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

                // ✅ PHASE 2: Use HostPropertyCache for Robust Wall Centering
                if (_hostPropertyCache != null && zone.StructuralElementIdValue > 0)
                {
                    try
                    {
                        // ✅ CRITICAL FIX: Ensure we use the correct element ID for cache lookup.
                        // The cache stores properties keyed by the element itself.
                        ElementId hostId = new ElementId(zone.StructuralElementIdValue);
                        
                        // ✅ CRITICAL FIX: Use ElementRetrievalService to check if element is in linked doc
                        // The cache logic assumes local coordinates, which fails for linked instances with transforms
                        // For linked elements, we MUST skip the cache and fall back to the robust WallCenterlineHelper
                        Element hostElement = ElementRetrievalService.GetElementFromDocumentOrLinked(_doc, hostId);
                        
                        if (hostElement != null)
                        {
                             // ✅ LINKED MODEL CHECK: If element is from a linked document, skip cache
                             // HostPropertyCache computes properties in local space, but we need Host Space adjustments
                             // The WallCenterlineHelper (fallback below) handles inverse transforms correctly.
                             if (hostElement.Document.Title != _doc.Title)
                             {
                                 if (!DeploymentConfiguration.DeploymentMode)
                                 {
                                     SafeFileLogger.SafeAppendText("placement_debug.log",
                                         $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] 🔄 SKIPPING CACHE (Linked Element): Zone {zone.Id}, HostId={zone.StructuralElementIdValue} is in linked doc '{hostElement.Document.Title}' - using robust fallback\n");
                                 }
                                 goto SkipCache; // Jump to fallback
                             }

                             // Get cached properties (calculates once per host)
                             // This is the authoritative source for the wall's geometry in the current session.
                             HostProperties props = _hostPropertyCache.GetOrCompute(hostElement);
                         
                         if (props != null && props.Normal != null && !props.Normal.IsZeroLength())
                         {
                             // ✅ CALCULATE CENTERLINE PROJECTION:
                             // AdjustedPoint = Point - ((Point - Center) dot Normal) * Normal
                             // This projects the point onto the centerline plane defined by Center and Normal
                             
                             XYZ centerPoint = new XYZ(props.CenterlineX, props.CenterlineY, props.CenterlineZ);
                             XYZ vectorToCenter = placementPoint - centerPoint;
                             double distanceToCenterPlane = vectorToCenter.DotProduct(props.Normal);
                             
                             XYZ adjustedPoint = placementPoint - (props.Normal * distanceToCenterPlane);
                             
                             if (!DeploymentConfiguration.DeploymentMode)
                             {
                                 SafeFileLogger.SafeAppendText("placement_debug.log",
                                     $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ CACHE HIT: Zone {zone.Id}, HostId={zone.StructuralElementIdValue}\n" +
                                     $"  Inputs: Point=({placementPoint.X:F3}, {placementPoint.Y:F3}, {placementPoint.Z:F3}), Normal=({props.Normal.X:F3}, {props.Normal.Y:F3}, {props.Normal.Z:F3})\n" +
                                     $"  Result: Offset={distanceToCenterPlane * 304.8:F1}mm, Adjusted=({adjustedPoint.X:F3}, {adjustedPoint.Y:F3}, {adjustedPoint.Z:F3})\n");
                             }
                             
                             return adjustedPoint;
                         }
                    }
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️ CACHE ERROR: {ex.Message} - Falling back to standard logic\n");
                    }
                }
            }
            
            SkipCache:

            // ✅ PERFORMANCE: Use pre-calculated sleeve placement point from database (enables multi-threading)
            // ✅ CRITICAL: SleevePlacementPoint is now calculated during refresh using bbox method (same as dampers)
            // This is the final placement point at wall centerline, no adjustment needed
            // ✅ CRITICAL FIX: Check X/Y/Z values directly (not just computed property) to handle cases where values are saved but property returns zero
            bool hasSleevePlacementPoint = (zone.SleevePlacementPointX != 0.0 || zone.SleevePlacementPointY != 0.0 || zone.SleevePlacementPointZ != 0.0);
            
            // ✅ DIAGNOSTIC: Log sleeve placement point values
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] 🔍 SLEEVE PLACEMENT POINT CHECK: Zone {zone.Id}, " +
                    $"HasSleevePlacementPoint={hasSleevePlacementPoint}, " +
                    $"SleevePlacementPointX={zone.SleevePlacementPointX:F6}ft ({zone.SleevePlacementPointX * 304.8:F1}mm), " +
                    $"SleevePlacementPointY={zone.SleevePlacementPointY:F6}ft ({zone.SleevePlacementPointY * 304.8:F1}mm), " +
                    $"SleevePlacementPointZ={zone.SleevePlacementPointZ:F6}ft ({zone.SleevePlacementPointZ * 304.8:F1}mm), " +
                    $"PlacementPoint (Input)=({placementPoint.X:F6}ft, {placementPoint.Y:F6}ft, {placementPoint.Z:F6}ft)\n");
            }
            
            // ✅ CRITICAL FIX: Check WallCenterlinePoint FIRST (always populated by Refresh)
            // SleevePlacementPoint may be zeros if placement hasn't happened yet, but WallCenterlinePoint is always set during Refresh
            bool hasWallCenterline = (zone.WallCenterlinePointX != 0.0 || zone.WallCenterlinePointY != 0.0 || zone.WallCenterlinePointZ != 0.0);
            
            // ✅ FORCE DETECTION MODE OR NORMAL MODE: Use WallCenterlinePoint if available (source of truth from Refresh)
            if (hasWallCenterline)
            {
                XYZ savedCenterlinePoint = new XYZ(zone.WallCenterlinePointX, zone.WallCenterlinePointY, zone.WallCenterlinePointZ);
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ USING WALL CENTERLINE POINT: Zone {zone.Id}, " +
                        $"WallCenterlinePoint=({savedCenterlinePoint.X:F6}ft, {savedCenterlinePoint.Y:F6}ft, {savedCenterlinePoint.Z:F6}ft), " +
                        $"ForceMode={_isForceDetectionMode}\n");
                }
                return savedCenterlinePoint; // ✅ CRITICAL: Actually RETURN the point!
            }
            
            // ✅ FALLBACK 1: Try SleevePlacementPoint if WallCenterlinePoint is not available
            if (hasSleevePlacementPoint)
            {
                XYZ savedPlacementPoint = new XYZ(zone.SleevePlacementPointX, zone.SleevePlacementPointY, zone.SleevePlacementPointZ);
                
                // Validate saved placement point against intersection point
                double distance = placementPoint.DistanceTo(savedPlacementPoint);
                const double MAX_VALID_OFFSET = 5.0; // 5 feet to accommodate thick walls
                
                if (distance <= MAX_VALID_OFFSET)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ USING SLEEVE PLACEMENT POINT: Zone {zone.Id}, " +
                            $"SleevePlacementPoint=({savedPlacementPoint.X:F6}ft, {savedPlacementPoint.Y:F6}ft, {savedPlacementPoint.Z:F6}ft), " +
                            $"Distance={distance * 304.8:F1}mm\n");
                    }
                    return savedPlacementPoint;
                }
                else if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️ STALE PLACEMENT POINT: Zone {zone.Id}, " +
                        $"Distance={distance * 304.8:F1}mm > {MAX_VALID_OFFSET * 304.8:F1}mm - RECALCULATING\n");
                }
            }
            
            // ✅ FALLBACK 2: No saved data available - calculate fresh (requires Revit API)
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("placement_debug.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️⚠️⚠️ NO SAVED CENTERLINE: Zone {zone.Id}, " +
                    $"WallCenterlinePoint=({zone.WallCenterlinePointX:F6}ft, {zone.WallCenterlinePointY:F6}ft, {zone.WallCenterlinePointZ:F6}ft), " +
                    $"SleevePlacementPoint=({zone.SleevePlacementPointX:F6}ft, {zone.SleevePlacementPointY:F6}ft, {zone.SleevePlacementPointZ:F6}ft) - " +
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
                // ✅ FORCE DETECTION: When active, this uses the fresh intersection point (not saved data)
                
                if (!DeploymentConfiguration.DeploymentMode && _isForceDetectionMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] 🔄 CALCULATING FRESH CENTERLINE: Zone {zone.Id}, " +
                        $"IntersectionPoint=({placementPoint.X:F6}ft, {placementPoint.Y:F6}ft, {placementPoint.Z:F6}ft) - Will adjust to wall centerline\n");
                }
                
                XYZ centerlinePoint = WallCenterlineHelper.GetElementCenterlinePoint(hostElement, placementPoint, _doc);

                if (centerlinePoint != null && centerlinePoint != placementPoint)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double deltaX = Math.Abs(centerlinePoint.X - placementPoint.X);
                        double deltaY = Math.Abs(centerlinePoint.Y - placementPoint.Y);
                        double deltaZ = Math.Abs(centerlinePoint.Z - placementPoint.Z);
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ HOST CENTERLINE ADJUSTED: Zone {zone.Id}, HostType={zone.StructuralElementType}, HostId={zone.StructuralElementId.IntegerValue}\n" +
                            $"  Input (Intersection): ({placementPoint.X:F6}ft, {placementPoint.Y:F6}ft, {placementPoint.Z:F6}ft)\n" +
                            $"  Output (Centerline):  ({centerlinePoint.X:F6}ft, {centerlinePoint.Y:F6}ft, {centerlinePoint.Z:F6}ft)\n" +
                            $"  Delta: ΔX={deltaX * 304.8:F1}mm, ΔY={deltaY * 304.8:F1}mm, ΔZ={deltaZ * 304.8:F1}mm\n");
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

