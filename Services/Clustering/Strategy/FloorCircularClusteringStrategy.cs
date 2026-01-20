using System;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy
{
    /// <summary>
    /// Clustering strategy for floor circular sleeves (round pipes/ducts).
    /// Phase 4A: Extracted from UniversalClusterService with edge-to-edge distance calculation.
    /// Uses 2D X-Y plane (ignores Z) and edge-to-edge distance: centerToCenter - (radius1 + radius2).
    /// </summary>
    public class FloorCircularClusteringStrategy : ClusteringStrategyBase
    {
        public FloorCircularClusteringStrategy()
            : base()
        {
        }

        public override string GetStrategyName() => "FloorCircularClusteringStrategy";

        public override bool CanHandle(SleeveGroupKey groupKey, System.Collections.Generic.List<dynamic> sleeves)
        {
            // ✅ Floor hosts only
            if (groupKey.hostType != "Floor")
                return false;

            // ✅ Check if sleeves are circular (round pipes/ducts)
            if (sleeves == null || sleeves.Count == 0)
                return false;

            // Check if any sleeve is circular (round pipe/duct)
            foreach (var sleeve in sleeves.Take(5)) // Check first 5 as sample
            {
                if (IsCircular(sleeve))
                    return true;
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

                if (cz1 == null || cz2 == null || cz1.SleeveDiameter <= 0 || cz2.SleeveDiameter <= 0)
                {
                    SafeFileLogger.SafeAppendText("strategy_errors.log",
                        $"[{GetStrategyName()}] Invalid diameters - falling back to rectangular strategy");
                    // Fallback to rectangular strategy if not truly round
                    return new FloorRectangularClusteringStrategy().CheckProximity(sleeve1, sleeve2, tolerance, orientation, document);
                }

                // ✅ Get placement points from database
                var placement1 = GetPlacementPointFromSleeve(sleeve1);
                var placement2 = GetPlacementPointFromSleeve(sleeve2);

                if (placement1 == null || placement2 == null)
                {
                    SafeFileLogger.SafeAppendText("strategy_errors.log",
                        $"[{GetStrategyName()}] Invalid placement points - returning false");
                    return false;
                }

                // ✅ Calculate edge-to-edge distance (2D X-Y plane, ignore Z)
                double dx = placement2.Value.X - placement1.Value.X;
                double dy = placement2.Value.Y - placement1.Value.Y;
                double centerToCenterDistance = Math.Sqrt(dx * dx + dy * dy);

                double radius1 = cz1.SleeveDiameter / 2.0;
                double radius2 = cz2.SleeveDiameter / 2.0;
                double edgeToEdgeDistance = centerToCenterDistance - (radius1 + radius2);
                double minDistance = Math.Max(0, edgeToEdgeDistance);

                bool withinTolerance = minDistance <= tolerance;

                // ✅ Use EdgeToEdgeProximityChecker from Phase 2
                var checker = new EdgeToEdgeProximityChecker();
                // ✅ Use correct IProximityChecker interface signature
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

                if (cz1 == null || cz2 == null || cz1.SleeveDiameter <= 0 || cz2.SleeveDiameter <= 0)
                    return double.MaxValue;

                var placement1 = GetPlacementPointFromSleeve(sleeve1);
                var placement2 = GetPlacementPointFromSleeve(sleeve2);

                if (placement1 == null || placement2 == null)
                    return double.MaxValue;

                // 2D X-Y distance (ignore Z)
                double dx = placement2.Value.X - placement1.Value.X;
                double dy = placement2.Value.Y - placement1.Value.Y;
                double centerToCenterDistance = Math.Sqrt(dx * dx + dy * dy);

                double radius1 = cz1.SleeveDiameter / 2.0;
                double radius2 = cz2.SleeveDiameter / 2.0;
                double edgeToEdgeDistance = centerToCenterDistance - (radius1 + radius2);

                return Math.Max(0, edgeToEdgeDistance);
            }
            catch
            {
                return double.MaxValue;
            }
        }
    }
}

