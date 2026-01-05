using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy
{
    /// <summary>
    /// Factory for creating appropriate clustering strategies based on sleeve properties.
    /// Phase 4A: Decision tree with logging for strategy selection.
    /// </summary>
    public class ClusteringStrategyFactory
    {
        public ClusteringStrategyFactory()
        {
        }

        /// <summary>
        /// Get appropriate clustering strategy for the given group key and sleeves.
        /// Decision logic: Floor (Circular/Rectangular) -> Wall/Framing (Rotated/Axis-Aligned).
        /// </summary>
        public IClusteringStrategy GetStrategy(SleeveGroupKey groupKey, List<ClusteringSleeveDto> sleeves)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeves == null || sleeves.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("strategy_factory_errors.log",
                        $"[ClusteringStrategyFactory] Invalid sleeves input - returning FloorRectangularClusteringStrategy");
                    return new FloorRectangularClusteringStrategy();
                }

                // ✅ Log strategy selection for diagnostics
                SafeFileLogger.SafeAppendText("clustering_strategy.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] ========== STRATEGY SELECTION ==========\n");
                SafeFileLogger.SafeAppendText("clustering_strategy.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] HostType={groupKey.hostType}, SystemType={groupKey.systemType}, Orientation={groupKey.orientation}\n");
                SafeFileLogger.SafeAppendText("clustering_strategy.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve count: {sleeves.Count}\n");

                // Check if sleeves are circular or rectangular
                bool isCircular = IsCircular(sleeves);
                bool hasRotated = HasRotated(sleeves);

                SafeFileLogger.SafeAppendText("clustering_strategy.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] IsCircular: {isCircular}, HasRotated: {hasRotated}\n");

                // ✅ Decision tree
                if (groupKey.hostType == "Floor")
                {
                    if (isCircular)
                    {
                        SafeFileLogger.SafeAppendText("clustering_strategy.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ✅ SELECTED: FloorCircularClusteringStrategy (2D X-Y, edge-to-edge)\n");
                        SafeFileLogger.SafeAppendText("clustering_strategy.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ========== END STRATEGY SELECTION ==========\n\n");
                        return new FloorCircularClusteringStrategy();
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("clustering_strategy.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ✅ SELECTED: FloorRectangularClusteringStrategy (2D X-Y, bbox overlap)\n");
                        SafeFileLogger.SafeAppendText("clustering_strategy.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ========== END STRATEGY SELECTION ==========\n\n");
                        return new FloorRectangularClusteringStrategy();
                    }
                }
                else if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
                {
                    if (hasRotated)
                    {
                        SafeFileLogger.SafeAppendText("clustering_strategy.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ✅ SELECTED: WallRotatedClusteringStrategy (rotation angle validation)\n");
                        SafeFileLogger.SafeAppendText("clustering_strategy.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ========== END STRATEGY SELECTION ==========\n\n");
                        return new WallRotatedClusteringStrategy();
                    }
                    else
                    {
                        SafeFileLogger.SafeAppendText("clustering_strategy.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ✅ SELECTED: WallAxisAlignedStrategy (2D {groupKey.orientation}-Z plane)\n");
                        SafeFileLogger.SafeAppendText("clustering_strategy.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] ========== END STRATEGY SELECTION ==========\n\n");
                        return new WallAxisAlignedStrategy();
                    }
                }

                // Fallback to floor rectangular strategy
                SafeFileLogger.SafeAppendText("clustering_strategy.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] ⚠️ FALLBACK: FloorRectangularClusteringStrategy (unknown host type)\n");
                SafeFileLogger.SafeAppendText("clustering_strategy.log",
                    $"[{DateTime.Now:HH:mm:ss.fff}] ========== END STRATEGY SELECTION ==========\n\n");
                return new FloorRectangularClusteringStrategy();
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("strategy_factory_errors.log",
                    $"[ClusteringStrategyFactory] Exception in GetStrategy: {ex.Message}, StackTrace: {ex.StackTrace}");
                // Safe fallback
                return new FloorRectangularClusteringStrategy();
            }
        }

        /// <summary>
        /// Check if any sleeve in the list is circular (round pipe/duct).
        /// </summary>
        private bool IsCircular(List<ClusteringSleeveDto> sleeves)
        {
            try
            {
                foreach (var sleeve in sleeves.Take(10)) // Check first 10 as sample
                {
                    var cz = sleeve?.ClashZone;
                    
                    // Check if it's a round pipe/duct by category and shape
                    if (cz != null)
                    {
                        bool isPipe = cz.MepElementCategory?.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0;
                        bool isRoundDuct = cz.MepElementCategory?.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                          (string.Equals(cz.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                           (cz.MepElementSizeData != null &&
                                            (string.Equals(cz.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                             string.Equals(cz.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase))));

                        if (isPipe || isRoundDuct)
                            return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Check if any sleeve in the list is rotated (non-axis-aligned).
        /// </summary>
        private bool HasRotated(List<ClusteringSleeveDto> sleeves)
        {
            try
            {
                foreach (var sleeve in sleeves.Take(10)) // Check first 10 as sample
                {
                    var cz = sleeve?.ClashZone;
                    if (cz != null)
                    {
                        double rotationAngle = cz.MepElementRotationAngle;
                        if (Math.Abs(rotationAngle) > 1e-6 && !IsAxisAlignedAngle(rotationAngle))
                            return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Check if rotation angle is axis-aligned (0°, 90°, 180°, 270°).
        /// </summary>
        private bool IsAxisAlignedAngle(double angleRad)
        {
            double angleDeg = angleRad * 180.0 / Math.PI;
            while (angleDeg < 0) angleDeg += 360;
            while (angleDeg >= 360) angleDeg -= 360;

            double thresholdDegrees = 2.0;
            double distTo0 = Math.Min(angleDeg, 360 - angleDeg);
            double distTo90 = Math.Abs(angleDeg - 90);
            double distTo180 = Math.Abs(angleDeg - 180);
            double distTo270 = Math.Abs(angleDeg - 270);

            return distTo0 < thresholdDegrees || distTo90 < thresholdDegrees ||
                   distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
        }
    }
}

