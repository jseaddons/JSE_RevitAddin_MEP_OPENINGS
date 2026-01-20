using System;
using System.Diagnostics;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Geometry
{
    /// <summary>
    /// ✅ SOLID COMPLIANCE (SRP): Service responsible for calculating rotated bounding box coordinates.
    /// Pure mathematical operations - no database or Revit API dependencies.
    /// Includes performance monitoring, safety features, and comprehensive validation.
    /// </summary>
    public class RotatedBoundingBoxCalculationService
    {
        /// <summary>
        /// ✅ SRP: Calculates rotated bounding box in LOCAL coordinates (centered at placement point).
        /// Includes performance monitoring and safety features.
        /// </summary>
        /// <param name="placementPoint">Placement point (center of sleeve)</param>
        /// <param name="width">Width of sleeve (in internal units, feet)</param>
        /// <param name="height">Height of sleeve (in internal units, feet)</param>
        /// <param name="depth">Depth of sleeve (in internal units, feet)</param>
        /// <returns>Rotated bounding box coordinates (minX, minY, minZ, maxX, maxY, maxZ) or null if invalid</returns>
        public (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)? CalculateRotatedBoundingBox(
            XYZ placementPoint, double width, double height, double depth)
        {
            // ✅ PERFORMANCE MONITORING: Track calculation time
            var timer = Stopwatch.StartNew();
            
            try
            {
                // ✅ SAFETY: Comprehensive input validation
                if (placementPoint == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning("[RotatedBoundingBoxCalculationService] Placement point is null");
                    return null;
                }

                if (width <= 0 || height <= 0 || depth <= 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RotatedBoundingBoxCalculationService] Invalid dimensions: W={width}, H={height}, D={depth}");
                    return null;
                }

                // ✅ SAFETY: Check for extreme values that might indicate calculation errors
                if (width > 1000 || height > 1000 || depth > 1000) // ~300m in feet
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Warning($"[RotatedBoundingBoxCalculationService] Suspiciously large dimensions: W={width}, H={height}, D={depth}");
                }

                // Calculate half dimensions
                double halfWidth = width / 2.0;
                double halfHeight = height / 2.0;
                double halfDepth = depth / 2.0;

                // Calculate rotated bounding box in LOCAL coordinates (centered at placement point)
                var rotatedLocalMinX = placementPoint.X - halfWidth;
                var rotatedLocalMinY = placementPoint.Y - halfHeight;
                var rotatedLocalMinZ = placementPoint.Z - halfDepth;
                var rotatedLocalMaxX = placementPoint.X + halfWidth;
                var rotatedLocalMaxY = placementPoint.Y + halfHeight;
                var rotatedLocalMaxZ = placementPoint.Z + halfDepth;

                timer.Stop();
                
                // ✅ PERFORMANCE MONITORING: Log slow calculations
                if (!DeploymentConfiguration.DeploymentMode && timer.ElapsedMilliseconds > 10)
                {
                    DebugLogger.Info($"[RotatedBoundingBoxCalculationService] Calculated rotated bbox in {timer.ElapsedMilliseconds}ms (W={width:F3}, H={height:F3}, D={depth:F3})");
                }

                return (rotatedLocalMinX, rotatedLocalMinY, rotatedLocalMinZ,
                        rotatedLocalMaxX, rotatedLocalMaxY, rotatedLocalMaxZ);
            }
            catch (Exception ex)
            {
                timer.Stop();
                // ✅ SAFETY: Comprehensive error handling
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[RotatedBoundingBoxCalculationService] Error calculating rotated bounding box: {ex.Message}");
                    DebugLogger.Error($"[RotatedBoundingBoxCalculationService] Stack trace: {ex.StackTrace}");
                }
                return null;
            }
        }

        /// <summary>
        /// ✅ SRP: Calculates rotated bounding box from a ClashZone (convenience method).
        /// </summary>
        public (double minX, double minY, double minZ, double maxX, double maxY, double maxZ)? CalculateRotatedBoundingBoxFromZone(
            ClashZone zone, double width, double height, double depth)
        {
            if (zone == null || zone.SleevePlacementPoint == null)
                return null;

            return CalculateRotatedBoundingBox(
                zone.SleevePlacementPoint,
                width,
                height,
                depth);
        }
    }
}

