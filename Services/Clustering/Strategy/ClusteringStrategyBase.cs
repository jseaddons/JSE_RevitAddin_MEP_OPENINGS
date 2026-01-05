using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Data;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Safety;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Strategy
{
    /// <summary>
    /// Base class for clustering strategies with common functionality.
    /// Phase 4A: Provides helper methods and proximity checker integration.
    /// </summary>
    public abstract class ClusteringStrategyBase : IClusteringStrategy
    {
        protected ClusteringStrategyBase()
        {
        }

        public abstract bool CanHandle(SleeveGroupKey groupKey, List<ClusteringSleeveDto> sleeves);
        public abstract string GetStrategyName();
        public abstract bool CheckProximity(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2, double tolerance, string orientation, Document document = null);
        public abstract double CalculateProximityDistance(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2, string orientation);

        /// <summary>
        /// Default FormClusters implementation using iterative expansion (flood-fill).
        /// </summary>
        public virtual List<List<ClusteringSleeveDto>> FormClusters(List<ClusteringSleeveDto> sleeves, double tolerance, string orientation, Document document = null)
        {
            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (sleeves == null || sleeves.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("strategy_errors.log",
                        $"[{GetStrategyName()}] Invalid sleeves input - returning empty clusters");
                    return new List<List<ClusteringSleeveDto>>();
                }

                var clusters = new List<List<ClusteringSleeveDto>>();
                var processed = new HashSet<int>();

                foreach (var sleeve in sleeves)
                {
                    if (processed.Contains(sleeve.SleeveInstanceId))
                        continue;

                    var cluster = new List<ClusteringSleeveDto> { sleeve };
                    processed.Add(sleeve.SleeveInstanceId);

                    // ✅ Iterative expansion to find all connected sleeves
                    bool foundNewNeighbors = true;
                    while (foundNewNeighbors)
                    {
                        foundNewNeighbors = false;
                        var currentClusterSize = cluster.Count;

                        // Find neighbors for all sleeves in current cluster
                        foreach (var clusterSleeve in cluster.ToList())
                        {
                            foreach (var otherSleeve in sleeves)
                            {
                                if (processed.Contains(otherSleeve.SleeveInstanceId))
                                    continue;

                                // Check cross-wall prevention (for wall strategies)
                                if (!ShouldCluster(clusterSleeve, otherSleeve, document))
                                    continue;

                                if (CheckProximity(clusterSleeve, otherSleeve, tolerance, orientation, document))
                                {
                                    cluster.Add(otherSleeve);
                                    processed.Add(otherSleeve.SleeveInstanceId);
                                    foundNewNeighbors = true;
                                }
                            }
                        }
                    }

                    // Only add clusters with more than 1 sleeve
                    if (cluster.Count > 1)
                    {
                        clusters.Add(cluster);
                    }
                }

                return clusters;
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("strategy_errors.log",
                    $"[{GetStrategyName()}] Exception in FormClusters: {ex.Message}, StackTrace: {ex.StackTrace}");
                return new List<List<ClusteringSleeveDto>>();
            }
        }

        /// <summary>
        /// Check if two sleeves should be clustered (pre-filtering, e.g., cross-wall prevention).
        /// Override in subclasses for host-specific validation.
        /// </summary>
        protected virtual bool ShouldCluster(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2, Document document)
        {
            // Default: allow clustering (no special restrictions)
            return true;
        }

        /// <summary>
        /// Get placement point from sleeve using ClashZone data.
        /// </summary>
        protected XYZ? GetPlacementPointFromSleeve(ClusteringSleeveDto sleeve)
        {
            try
            {
                if (sleeve?.ClashZone == null)
                    return null;

                var cz = sleeve.ClashZone;
                if (cz == null)
                    return null;

                // ✅ Use active document placement point (priority)
                if (cz.SleevePlacementPointActiveDocumentX != 0 ||
                    cz.SleevePlacementPointActiveDocumentY != 0 ||
                    cz.SleevePlacementPointActiveDocumentZ != 0)
                {
                    return new XYZ(
                        cz.SleevePlacementPointActiveDocumentX,
                        cz.SleevePlacementPointActiveDocumentY,
                        cz.SleevePlacementPointActiveDocumentZ
                    );
                }

                // Fallback to regular placement point
                if (cz.SleevePlacementPointX != 0 ||
                    cz.SleevePlacementPointY != 0 ||
                    cz.SleevePlacementPointZ != 0)
                {
                    return new XYZ(
                        cz.SleevePlacementPointX,
                        cz.SleevePlacementPointY,
                        cz.SleevePlacementPointZ
                    );
                }

                // Fallback: Calculate center from bounding box
                var bbox = sleeve.BoundingBox;
                if (bbox != null)
                {
                    return new XYZ(
                        (bbox.Min.X + bbox.Max.X) / 2.0,
                        (bbox.Min.Y + bbox.Max.Y) / 2.0,
                        (bbox.Min.Z + bbox.Max.Z) / 2.0
                    );
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Get sleeve radii from ClashZone data (only for round pipes/ducts).
        /// </summary>
        protected (double radius1, double radius2) GetSleeveRadii(ClusteringSleeveDto sleeve1, ClusteringSleeveDto sleeve2)
        {
            double radius1 = 0;
            double radius2 = 0;

            try
            {
                if (sleeve1?.ClashZone != null)
                {
                    var cz1 = sleeve1.ClashZone;
                    if (cz1 != null && cz1.SleeveDiameter > 0)
                    {
                        radius1 = cz1.SleeveDiameter / 2.0;
                    }
                }

                if (sleeve2?.ClashZone != null)
                {
                    var cz2 = sleeve2.ClashZone;
                    if (cz2 != null && cz2.SleeveDiameter > 0)
                    {
                        radius2 = cz2.SleeveDiameter / 2.0;
                    }
                }
            }
            catch { }

            return (radius1, radius2);
        }

        /// <summary>
        /// Check if sleeve is circular (round pipe/duct).
        /// Uses ClashZone category, DuctShape, MepElementSizeData.Shape, and SleeveDiameter.
        /// </summary>
        protected bool IsCircular(ClusteringSleeveDto sleeve)
        {
            try
            {
                var cz = sleeve?.ClashZone;
                if (cz == null)
                    return false;

                // ✅ Pipes are always circular/round
                bool isPipe = cz.MepElementCategory?.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0;
                if (isPipe)
                    return true;

                // ✅ Check if it's a round duct by DuctShape or MepElementSizeData.Shape
                bool isRoundDuct = cz.MepElementCategory?.IndexOf("Duct", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                  (string.Equals(cz.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                   (cz.MepElementSizeData != null &&
                                    (string.Equals(cz.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                     string.Equals(cz.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase))));

                // ✅ Also check SleeveDiameter > 0 (indicates circular sleeve)
                bool hasDiameter = cz.SleeveDiameter > 0;

                return isRoundDuct || hasDiameter;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Check if rotation angle is straight axis-aligned to WCS (0°, 90°, 180°, 270°).
        /// Returns false for rotated axis-aligned (non-straight) angles (45°, 135°, 225°, 315°, etc.).
        /// </summary>
        protected bool IsStraightAxisAlignedAngle(double angleRad)
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

