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
    /// Clustering strategy for wall axis-aligned sleeves.
    /// Phase 4A: Extracted from UniversalClusterService with cross-wall prevention.
    /// Uses 2D X-Z or Y-Z plane (ignores wall depth) based on orientation.
    /// </summary>
    public class WallAxisAlignedStrategy : ClusteringStrategyBase
    {
        public WallAxisAlignedStrategy()
            : base()
        {
        }

        public override string GetStrategyName() => "WallAxisAlignedStrategy";

        public override bool CanHandle(SleeveGroupKey groupKey, System.Collections.Generic.List<dynamic> sleeves)
        {
            // ✅ Wall or Structural Framing hosts only
            return groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing";
        }

        /// <summary>
        /// ✅ CRITICAL: Check wall hosts BEFORE proximity check to prevent cross-wall clustering.
        /// </summary>
        protected override bool ShouldCluster(dynamic sleeve1, dynamic sleeve2, Document document)
        {
            try
            {
                // ✅ CRITICAL FIX: Check wall hosts BEFORE bounding box overlap
                if (sleeve1?.HostType == "Wall" && sleeve2?.HostType == "Wall")
                {
                    var host1Id = sleeve1.ClashZone?.StructuralElementIdValue ?? -1;
                    var host2Id = sleeve2.ClashZone?.StructuralElementIdValue ?? -1;

                    if (host1Id > 0 && host2Id > 0 && host1Id != host2Id)
                    {
                        // Different walls - do NOT cluster
                        SafeFileLogger.SafeAppendText("wall_clustering.log",
                            $"[{GetStrategyName()}] Rejected clustering: sleeves on different walls ({host1Id} vs {host2Id})");
                        return false;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
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
                bool isRotated = Math.Abs(angle1) > 1e-6 && !IsStraightAxisAlignedAngle(angle1);
                
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

                // ✅ Calculate 2D distance based on orientation
                double? distance = null;

                if (orientation == "X")
                {
                    // X-oriented walls: 2D distance in X,Z plane (ignore Y)
                    distance = DistanceCalculator.CalculateMinimumDistance2D(
                        bbox1.Min.X, bbox1.Min.Z, bbox1.Max.X, bbox1.Max.Z,
                        bbox2.Min.X, bbox2.Min.Z, bbox2.Max.X, bbox2.Max.Z);
                }
                else
                {
                    // Y-oriented walls: 2D distance in Y,Z plane (ignore X)
                    distance = DistanceCalculator.CalculateMinimumDistance2D(
                        bbox1.Min.Y, bbox1.Min.Z, bbox1.Max.Y, bbox1.Max.Z,
                        bbox2.Min.Y, bbox2.Min.Z, bbox2.Max.Y, bbox2.Max.Z);
                }

                return distance ?? double.MaxValue;
            }
            catch
            {
                return double.MaxValue;
            }
        }
    }
}

