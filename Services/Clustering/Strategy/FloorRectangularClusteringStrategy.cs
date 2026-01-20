using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy
{
    /// <summary>
    /// Clustering strategy for floor rectangular sleeves.
    /// Phase 4A: Extracted from UniversalClusterService with 2D X-Y bounding box distance calculation.
    /// Uses pre-calculated bounding boxes from database.
    /// </summary>
    public class FloorRectangularClusteringStrategy : ClusteringStrategyBase
    {
        public FloorRectangularClusteringStrategy()
            : base()
        {
        }

        public override string GetStrategyName() => "FloorRectangularClusteringStrategy";

        public override bool CanHandle(SleeveGroupKey groupKey, System.Collections.Generic.List<dynamic> sleeves)
        {
            // ✅ Floor hosts only
            return groupKey.hostType == "Floor";
        }

        public override bool CheckProximity(dynamic sleeve1, dynamic sleeve2, double tolerance, string orientation, Document document = null)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeve1?.BoundingBox == null || sleeve2?.BoundingBox == null)
                {
                    SafeFileLogger.SafeAppendText("strategy_errors.log",
                        $"[{GetStrategyName()}] Invalid input: BoundingBox is null - returning false");
                    return false;
                }

                // ✅ DECISION: Delegate to factory to ensure corner-based logic is used for rectangular sleeves
                var cz1 = sleeve1.ClashZone as ClashZone;
                double angle1 = cz1?.MepElementRotationAngle ?? 0.0;
                bool isRotated = !IsStraightAxisAlignedAngle(angle1);

                var checker = ProximityCheckerFactory.CreateChecker(sleeve1, sleeve2, angle1, isRotated);
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
                var bbox1 = sleeve1?.BoundingBox;
                var bbox2 = sleeve2?.BoundingBox;

                if (bbox1 == null || bbox2 == null)
                    return double.MaxValue;

                // ✅ Use DistanceCalculator from Phase 1 (2D X-Y plane, ignore Z)
                double? distance = DistanceCalculator.CalculateMinimumDistance2D(
                    bbox1.Min.X, bbox1.Min.Y, bbox1.Max.X, bbox1.Max.Y,
                    bbox2.Min.X, bbox2.Min.Y, bbox2.Max.X, bbox2.Max.Y);

                return distance ?? double.MaxValue;
            }
            catch
            {
                return double.MaxValue;
            }
        }
    }
}

