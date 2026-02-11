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
        
        // Cache to store invariant wall data (Orientation, Center Coordinate, Width) by Wall ID
        private readonly Dictionary<int, (bool IsXWall, double CenterCoordinate, double Width, bool Success)> _wallCenterlineCache 
            = new Dictionary<int, (bool IsXWall, double CenterCoordinate, double Width, bool Success)>();

        // Cache for Structural Framing (Normal, Thickness)
        private readonly Dictionary<int, (XYZ Normal, double Thickness, bool Success)> _framingCenterlineCache 
            = new Dictionary<int, (XYZ Normal, double Thickness, bool Success)>();

        // Cache for Floors (Normal, Thickness)
        private readonly Dictionary<int, (XYZ Normal, double Thickness, bool Success)> _floorCenterlineCache 
            = new Dictionary<int, (XYZ Normal, double Thickness, bool Success)>();

        /// <summary>
        /// ✅ DIP COMPLIANCE: Constructor accepts optional performance monitor for dependency injection.
        /// </summary>
        public PlacementPointAdjustmentService(Document doc, PlacementPerformanceMonitor? performanceMonitor = null, bool isForceDetectionMode = false)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _performanceMonitor = performanceMonitor;
            _isForceDetectionMode = isForceDetectionMode;
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
                // Re-enabled: WallCenterlineHelper is now fixed to use RayTrace (Midpoint)
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

            // ✅ PERFORMANCE: Use pre-calculated sleeve placement point from database (enables multi-threading)
            // ✅ CRITICAL: SleevePlacementPoint is now calculated during refresh using bbox method (same as dampers)
            // This is the final placement point at wall centerline, no adjustment needed
            // ✅ CRITICAL FIX: Check X/Y/Z values directly (not just computed property) to handle cases where values are saved but property returns zero
            bool hasSleevePlacementPoint = (zone.SleevePlacementPointX != 0.0 || zone.SleevePlacementPointY != 0.0 || zone.SleevePlacementPointZ != 0.0);
            
            // ✅ DIAGNOSTIC: Log sleeve placement point values
            // ✅ PERFORMANCE: Skipped verbose logging for placement point check to improve speed (12ms -> <1ms)
            // if (!DeploymentConfiguration.DeploymentMode) { ... }
            
            // ✅ FORCE DETECTION MODE: If active, ALWAYS use intersection point directly without any saved data
            // This ensures force detection uses fresh calculated intersection points from clash detection
            // Even if saved placement point exists, it's from a previous operation and must be recalculated
            if (_isForceDetectionMode)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] 🔄🔄🔄 FORCE DETECTION MODE ACTIVE: Zone {zone.Id}\n" +
                        $"  Saved SleevePlacementPoint (IGNORED): ({zone.SleevePlacementPointX:F6}ft, {zone.SleevePlacementPointY:F6}ft, {zone.SleevePlacementPointZ:F6}ft)\n" +
                        $"  Fresh IntersectionPoint (WILL USE): ({placementPoint.X:F6}ft, {placementPoint.Y:F6}ft, {placementPoint.Z:F6}ft)\n" +
                        $"  Will calculate fresh centerline from geometry and structural element geometry.\n");
                }
                // Use intersection point and calculate centerline fresh - do NOT return saved point
                // Continue to fallback calculation below
            }
            // ✅ DATABASE-FIRST FAST PATH: Use pre-calculated SleevePlacementPoint
            // This is the authoritative placement point calculated during refresh/planning phase
            if (hasSleevePlacementPoint && !_isForceDetectionMode)
            {
                XYZ savedPlacementPoint = new XYZ(zone.SleevePlacementPointX, zone.SleevePlacementPointY, zone.SleevePlacementPointZ);
                
                // ✅ PERFORMANCE: Skipped verbose logging for fast path success
                // if (!DeploymentConfiguration.DeploymentMode) { ... }
                return savedPlacementPoint; // Return immediately - no Revit queries, no validation needed if DB data exists
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
                // ✅ FORCE DETECTION: When active, this uses the fresh intersection point (not saved data)
                
                if (!DeploymentConfiguration.DeploymentMode && _isForceDetectionMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] 🔄 CALCULATING FRESH CENTERLINE: Zone {zone.Id}, " +
                        $"IntersectionPoint=({placementPoint.X:F6}ft, {placementPoint.Y:F6}ft, {placementPoint.Z:F6}ft) - Will adjust to wall centerline\n");
                }
                

                XYZ centerlinePoint = null;
                
                // ✅ OPTIMIZED: Use caching for Walls, Framing, and Floors
                if (hostElement is Wall wall)
                {
                    int wallId = wall.Id.IntegerValue;
                    if (!_wallCenterlineCache.TryGetValue(wallId, out var data))
                    {
                        data = WallCenterlineHelper.GetWallInvariantData(wall, _doc);
                        _wallCenterlineCache[wallId] = data;
                    }

                    if (data.Success)
                    {
                        if (data.IsXWall) centerlinePoint = new XYZ(placementPoint.X, data.CenterCoordinate, placementPoint.Z);
                        else centerlinePoint = new XYZ(data.CenterCoordinate, placementPoint.Y, placementPoint.Z);
                    }
                    else
                    {
                        centerlinePoint = WallCenterlineHelper.GetElementCenterlinePoint(hostElement, placementPoint, _doc);
                    }
                }
                else if (hostElement.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                {
                    int elemId = hostElement.Id.IntegerValue;
                    if (!_framingCenterlineCache.TryGetValue(elemId, out var data))
                    {
                         data = WallCenterlineHelper.GetFramingInvariantData(hostElement, _doc);
                         _framingCenterlineCache[elemId] = data;
                    }
                    
                    if (data.Success)
                    {
                        // Calculate center: intersection - normal * halfThickness
                        centerlinePoint = placementPoint + data.Normal * (-data.Thickness / 2.0);
                    }
                    else
                    {
                        centerlinePoint = WallCenterlineHelper.GetElementCenterlinePoint(hostElement, placementPoint, _doc);
                    }
                }
                else if (hostElement is Floor floor)
                {
                    // ✅ FLOORS: Use IntersectionPoint directly (which is the center of the clash).
                    // Do NOT apply thickness offset, as IntersectionPoint is already at the correct elevation
                    // (intersecting the MEP element at the floor's location).
                    // Applying `Normal * -Thickness/2` assumes the point is on a FACE, which may not be true if
                    // the intersection is in the middle of the floor (e.g. vertical pipe).
                    // For floors, we trust the Clash Point.
                    return placementPoint;
                }
                else
                {
                    // Fallback for other categories
                    centerlinePoint = WallCenterlineHelper.GetElementCenterlinePoint(hostElement, placementPoint, _doc);
                }

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

