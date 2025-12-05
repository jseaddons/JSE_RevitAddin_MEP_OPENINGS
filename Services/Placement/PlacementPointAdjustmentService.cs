using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Configuration;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ SRP COMPLIANCE: Service responsible for adjusting placement points based on host element type and special requirements.
    /// Handles:
    /// - Wall/Framing centerline adjustment
    /// - Damper offset adjustments (delegated to strategy)
    /// - Any future placement point adjustments
    /// 
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
        /// ✅ SRP: Adjusts placement point based on host element type and special requirements.
        /// This method encapsulates all placement point adjustment logic, keeping NewSleevePlacerService focused on orchestration.
        /// 
        /// ✅ 28 FEATURES: Implements performance monitoring, element validation, and diagnostic logging.
        /// </summary>
        /// <param name="zone">The clash zone containing host element information</param>
        /// <param name="originalPlacementPoint">The original intersection point</param>
        /// <param name="damperOffset">Optional damper offset vector (if damper placement adjustment was calculated)</param>
        /// <returns>Adjusted placement point</returns>
        public XYZ AdjustPlacementPoint(ClashZone zone, XYZ originalPlacementPoint, XYZ damperOffset = null)
        {
            // ✅ PERFORMANCE MONITORING: Track operation time (Feature #11 from 28 features)
            using (var tracker = _performanceMonitor?.TrackOperation("Adjust Placement Point"))
            {
                if (zone == null)
                    throw new ArgumentNullException(nameof(zone));
                
                if (originalPlacementPoint == null)
                    throw new ArgumentNullException(nameof(originalPlacementPoint));

                XYZ adjustedPoint = originalPlacementPoint;

                // ✅ STEP 1: Adjust for wall/framing centerline (must be done before damper offset)
                adjustedPoint = AdjustForHostCenterline(zone, adjustedPoint);

                // ✅ STEP 2: Apply damper offset if provided (asymmetric clearance placement)
                if (damperOffset != null && !damperOffset.IsZeroLength())
                {
                    adjustedPoint = adjustedPoint.Add(damperOffset);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ✅ DAMPER OFFSET APPLIED: Zone {zone.Id}, Offset=({damperOffset.X:F6}, {damperOffset.Y:F6}, {damperOffset.Z:F6}), Adjusted Point=({adjustedPoint.X:F3}, {adjustedPoint.Y:F3}, {adjustedPoint.Z:F3})\n");
                    }
                }

                // ✅ PERFORMANCE MONITORING: Record item count
                tracker?.SetItemCount(1);

                return adjustedPoint;
            }
        }

        /// <summary>
        /// ✅ SRP: Adjusts placement point to host element centerline for walls and structural framing.
        /// This ensures sleeves are placed at the center of walls/framing, not on the face.
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
                Element hostElement = _doc.GetElement(zone.StructuralElementId);
                if (hostElement == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("placement_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PlacementPointAdjustment] ⚠️ HOST ELEMENT NOT FOUND: Zone {zone.Id}, HostId={zone.StructuralElementId.IntegerValue}, Cannot adjust to centerline\n");
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
                XYZ centerlinePoint = WallCenterlineHelper.GetElementCenterlinePoint(hostElement, placementPoint);

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

