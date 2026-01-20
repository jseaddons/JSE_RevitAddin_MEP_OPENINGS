using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy
{
    /// <summary>
    /// Clustering strategy for wall rotated sleeves.
    /// Phase 4A: Extracted from UniversalClusterService with same-axis validation and rotation transforms.
    /// Uses pre-calculated rotation matrix (cos/sin) from database.
    /// </summary>
    public class WallRotatedClusteringStrategy : ClusteringStrategyBase
    {
        public WallRotatedClusteringStrategy()
            : base()
        {
        }

        public override string GetStrategyName() => "WallRotatedClusteringStrategy";

        public override bool CanHandle(SleeveGroupKey groupKey, System.Collections.Generic.List<dynamic> sleeves)
        {
            // ✅ Wall or Structural Framing hosts only
            if (groupKey.hostType != "Wall" && groupKey.hostType != "Structural Framing")
                return false;

            // ✅ Check if sleeves are rotated (non-axis-aligned)
            if (sleeves == null || sleeves.Count == 0)
                return false;

            foreach (var sleeve in sleeves.Take(5)) // Check first 5 as sample
            {
                var cz = sleeve?.ClashZone as ClashZone;
                if (cz != null)
                {
                    double rotationAngle = cz.MepElementRotationAngle;
                    if (Math.Abs(rotationAngle) > 1e-6 && !IsAxisAlignedAngle(rotationAngle))
                        return true;
                }
            }

            return false;
        }

        public override bool CheckProximity(dynamic sleeve1, dynamic sleeve2, double tolerance, string orientation, Document document = null)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeve1?.ClashZone == null || sleeve2?.ClashZone == null)
                {
                    SafeFileLogger.SafeAppendText("strategy_errors.log",
                        $"[{GetStrategyName()}] Invalid input: ClashZone is null - returning false");
                    return false;
                }

                var cz1 = sleeve1.ClashZone as ClashZone;
                var cz2 = sleeve2.ClashZone as ClashZone;

                if (cz1 == null || cz2 == null)
                    return false;

                // ✅ Check if both sleeves are rotated and on same axis
                double rotationAngle1 = cz1.MepElementRotationAngle;
                double rotationAngle2 = cz2.MepElementRotationAngle;

                bool isRotated1 = Math.Abs(rotationAngle1) > 1e-6 && !IsStraightAxisAlignedAngle(rotationAngle1);
                bool isRotated2 = Math.Abs(rotationAngle2) > 1e-6 && !IsStraightAxisAlignedAngle(rotationAngle2);

                if (!isRotated1 || !isRotated2)
                {
                    // Not both rotated - fallback to axis-aligned strategy
                    return new WallAxisAlignedStrategy().CheckProximity(sleeve1, sleeve2, tolerance, orientation, document);
                }

                // ✅ Check if axes are on same axis line (difference of 0°, 180°, or 360°)
                double angle1Deg = rotationAngle1 * 180.0 / Math.PI;
                double angle2Deg = rotationAngle2 * 180.0 / Math.PI;
                while (angle1Deg < 0) angle1Deg += 360.0;
                while (angle1Deg >= 360.0) angle1Deg -= 360.0;
                while (angle2Deg < 0) angle2Deg += 360.0;
                while (angle2Deg >= 360.0) angle2Deg -= 360.0;

                double angleDiffDeg = Math.Abs(angle1Deg - angle2Deg);
                if (angleDiffDeg > 180.0) angleDiffDeg = 360.0 - angleDiffDeg;

                double axisToleranceDeg = 1.0; // 1 degree tolerance
                bool isSameAxis = angleDiffDeg <= axisToleranceDeg || Math.Abs(angleDiffDeg - 180.0) <= axisToleranceDeg;

                if (!isSameAxis)
                {
                    // Different axes - fallback to axis-aligned strategy
                    return new WallAxisAlignedStrategy().CheckProximity(sleeve1, sleeve2, tolerance, orientation, document);
                }

                // ✅ DECISION: Delegate to factory to ensure corner-based logic is used for rectangular sleeves
                var checker = ProximityCheckerFactory.CreateChecker(sleeve1, sleeve2, rotationAngle1, true);
                return checker.CheckProximity(sleeve1, sleeve2, tolerance);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("strategy_errors.log",
                    $"[{GetStrategyName()}] Exception in CheckProximity: {ex.Message}, StackTrace: {ex.StackTrace}");
                return false;
            }
        }

        public override double CalculateProximityDistance(dynamic sleeve1, dynamic sleeve2, string orientation)
        {
            try
            {
                var cz1 = sleeve1?.ClashZone as ClashZone;
                var cz2 = sleeve2?.ClashZone as ClashZone;

                if (cz1 == null || cz2 == null)
                    return double.MaxValue;

                // ✅ Get rotation angle for RotatedProximityChecker constructor
                double rotationAngle = cz1.MepElementRotationAngle;
                var checker = new RotatedProximityChecker(rotationAngle);
                double? distance = checker.CalculateDistance(sleeve1, sleeve2);
                return distance ?? double.MaxValue;
            }
            catch
            {
                return double.MaxValue;
            }
        }

        /// <summary>
        /// Helper method to check if an angle is axis-aligned (0°, 90°, 180°, 270°)
        /// </summary>
        private bool IsAxisAlignedAngle(double angleRad)
        {
            double angleDeg = angleRad * 180.0 / Math.PI;
            // Normalize to 0-360 range
            while (angleDeg < 0) angleDeg += 360;
            while (angleDeg >= 360) angleDeg -= 360;
            
            double thresholdDegrees = 2.0; // 2 degree tolerance
            double distTo0 = Math.Min(angleDeg, 360 - angleDeg);
            double distTo90 = Math.Abs(angleDeg - 90);
            double distTo180 = Math.Abs(angleDeg - 180);
            double distTo270 = Math.Abs(angleDeg - 270);
            
            return distTo0 < thresholdDegrees || distTo90 < thresholdDegrees || 
                   distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
        }
    }
}

