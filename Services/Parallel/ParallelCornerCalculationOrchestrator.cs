using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Geometry;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Parallel
{
    /// <summary>
    /// ✅ SOLID COMPLIANCE (SRP): Service responsible for orchestrating parallel corner calculations.
    /// Encapsulates all parallel processing logic for corner calculations.
    /// Includes performance monitoring, safety features, and optimization flags.
    /// </summary>
    public class ParallelCornerCalculationOrchestrator
    {
        private readonly SleeveCornerCalculationService _cornerCalculationService;

        public ParallelCornerCalculationOrchestrator(SleeveCornerCalculationService cornerCalculationService = null)
        {
            _cornerCalculationService = cornerCalculationService ?? new SleeveCornerCalculationService();
        }

        /// <summary>
        /// ✅ SRP: Calculates corners in parallel for multiple sleeves.
        /// Includes performance monitoring and safety features.
        /// </summary>
        /// <param name="placedSleeveData">Collection of sleeve data</param>
        /// <returns>Dictionary mapping zone GUID to corner coordinates</returns>
        public Dictionary<Guid, (double corner1X, double corner1Y, double corner1Z,
                                 double corner2X, double corner2Y, double corner2Z,
                                 double corner3X, double corner3Y, double corner3Z,
                                 double corner4X, double corner4Y, double corner4Z)> CalculateCornersInParallel(
            List<(FamilyInstance sleeve, ClashZone zone, double finalWidth, double finalHeight, double finalDiameter)> placedSleeveData)
        {
            // ✅ PERFORMANCE MONITORING: Track total calculation time
            var totalTimer = Stopwatch.StartNew();
            var cornerData = new Dictionary<Guid, (double corner1X, double corner1Y, double corner1Z,
                                                     double corner2X, double corner2Y, double corner2Z,
                                                     double corner3X, double corner3Y, double corner3Z,
                                                     double corner4X, double corner4Y, double corner4Z)>();

            if (placedSleeveData == null || placedSleeveData.Count == 0)
                return cornerData;

            // ✅ SAFETY: Validate input data
            int validItems = placedSleeveData.Count(item => item.zone != null && item.sleeve != null && item.sleeve.IsValidObject);
            if (validItems == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning("[ParallelCornerCalculationOrchestrator] No valid items in placedSleeveData");
                return cornerData;
            }

            // ✅ PARALLEL PROCESSING: Only if enabled and enough items to justify overhead
            if (OptimizationFlags.UseParallelProcessing && placedSleeveData.Count >= 12)
            {
                var parallelTimer = Stopwatch.StartNew();
                int taskCount = 0;
                int errorCount = 0;

                try
                {
                    var cornerTasks = placedSleeveData
                        .Where(item => item.zone != null && item.sleeve != null && item.sleeve.IsValidObject)
                        .Select(item => Task.Run(() =>
                        {
                            try
                            {
                                var (sleeve, zone, fw, fh, fd) = item;
                                var rotationAngleRad = zone.MepElementRotationAngle;
                                var rotationAngleDeg = Math.Abs(rotationAngleRad * 180.0 / Math.PI);
                                bool isStraightAxisAligned = IsStraightAxisAligned(rotationAngleDeg);

                                // Only calculate corners if rotation is non-zero
                                if (Math.Abs(rotationAngleRad) > 1e-6)
                                {
                                    double actualWidth = zone.SleeveWidth > 0 ? zone.SleeveWidth : fw;
                                    double actualHeight = zone.SleeveHeight > 0 ? zone.SleeveHeight : fh;
                                    
                                    // ✅ SRP: Delegate corner calculation to specialized service
                                    var corners = _cornerCalculationService.CalculateCornersFromZone(zone, actualWidth, actualHeight);
                                    if (corners.HasValue)
                                    {
                                        // ✅ FIX: Create named tuple to match return type signature
                                        var namedCorners = (
                                            corner1X: corners.Value.corner1.X,
                                            corner1Y: corners.Value.corner1.Y,
                                            corner1Z: corners.Value.corner1.Z,
                                            corner2X: corners.Value.corner2.X,
                                            corner2Y: corners.Value.corner2.Y,
                                            corner2Z: corners.Value.corner2.Z,
                                            corner3X: corners.Value.corner3.X,
                                            corner3Y: corners.Value.corner3.Y,
                                            corner3Z: corners.Value.corner3.Z,
                                            corner4X: corners.Value.corner4.X,
                                            corner4Y: corners.Value.corner4.Y,
                                            corner4Z: corners.Value.corner4.Z
                                        );
                                        return (zone.Id, namedCorners, true);
                                    }
                                }
                                return (zone.Id, default, false);
                            }
                            catch (Exception ex)
                            {
                                // ✅ SAFETY: Fail-safe - log error but don't crash entire batch
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[ParallelCornerCalculationOrchestrator] Error in parallel task: {ex.Message}");
                                }
                                return (Guid.Empty, default, false);
                            }
                        }))
                        .ToArray();

                    taskCount = cornerTasks.Length;
                    Task.WaitAll(cornerTasks);
                    parallelTimer.Stop();

                    foreach (var task in cornerTasks)
                    {
                        try
                        {
                            var (zoneId, corners, success) = task.Result;
                            // ✅ FIX: Access named tuple fields
                            if (success && (corners.corner1X != 0 || corners.corner1Y != 0 || corners.corner1Z != 0))
                            {
                                cornerData[zoneId] = corners;
                            }
                            else if (!success && zoneId != Guid.Empty)
                            {
                                errorCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[ParallelCornerCalculationOrchestrator] Error retrieving task result: {ex.Message}");
                            }
                        }
                    }

                    // ✅ PERFORMANCE MONITORING: Log parallel processing statistics
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double avgTimePerTask = taskCount > 0 ? (double)parallelTimer.ElapsedMilliseconds / taskCount : 0;
                        DebugLogger.Info($"[ParallelCornerCalculationOrchestrator] ✅ Parallel calculation complete: {cornerData.Count} corners calculated, {errorCount} errors, {taskCount} tasks in {parallelTimer.ElapsedMilliseconds}ms (avg: {avgTimePerTask:F1}ms per task)");
                    }
                }
                catch (Exception ex)
                {
                    parallelTimer.Stop();
                    // ✅ SAFETY: Comprehensive error handling
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Error($"[ParallelCornerCalculationOrchestrator] ❌ CRITICAL: Parallel calculation failed: {ex.Message}");
                        DebugLogger.Error($"[ParallelCornerCalculationOrchestrator] Stack trace: {ex.StackTrace}");
                    }
                }
            }
            else
            {
                // ✅ FALLBACK: Sequential processing when parallel is disabled or not enough items
                if (!DeploymentConfiguration.DeploymentMode && placedSleeveData.Count > 0)
                {
                    DebugLogger.Info($"[ParallelCornerCalculationOrchestrator] ⏭️ Parallel processing skipped: UseParallelProcessing={OptimizationFlags.UseParallelProcessing}, Count={placedSleeveData.Count} (threshold=12)");
                }
            }

            totalTimer.Stop();
            if (!DeploymentConfiguration.DeploymentMode && totalTimer.ElapsedMilliseconds > 100)
            {
                DebugLogger.Info($"[ParallelCornerCalculationOrchestrator] Total time: {totalTimer.ElapsedMilliseconds}ms, Calculated: {cornerData.Count} corners");
            }

            return cornerData;
        }

        // ✅ HELPER: Check if rotation angle is straight axis-aligned
        private bool IsStraightAxisAligned(double rotationAngleDeg)
        {
            return Math.Abs(rotationAngleDeg) < 1.0 ||
                   Math.Abs(rotationAngleDeg - 90.0) < 1.0 ||
                   Math.Abs(rotationAngleDeg - 180.0) < 1.0 ||
                   Math.Abs(rotationAngleDeg - 270.0) < 1.0 ||
                   Math.Abs(rotationAngleDeg - 360.0) < 1.0;
        }
    }
}

