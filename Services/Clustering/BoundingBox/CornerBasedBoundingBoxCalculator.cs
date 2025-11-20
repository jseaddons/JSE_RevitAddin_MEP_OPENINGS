using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox
{
    /// <summary>
    /// Corner-based watertight algorithm for calculating cluster bounding boxes.
    /// Phase 3: Extracted from UniversalClusterService with crash-safe guards.
    /// This algorithm works for ALL scenarios: single, stacked, inline, diagonal, grid.
    /// </summary>
    public static class CornerBasedBoundingBoxCalculator
    {
        /// <summary>
        /// Calculate bounding box using corner-based watertight algorithm.
        /// For each sleeve: Calculate 4 corners in world space, transform to rotated coordinate system, find min/max extents.
        /// </summary>
        /// <param name="cluster">List of sleeves in the cluster</param>
        /// <param name="rotationAngle">Cluster's intended rotated axis angle (NOT average of sleeve angles)</param>
        /// <param name="origin">Reference point for coordinate transformation (first sleeve center)</param>
        /// <param name="getClashZoneBySleeveInstanceId">Function to retrieve ClashZone by sleeve instance ID</param>
        /// <param name="xmlFilePath">Optional XML file path for ClashZone lookup</param>
        /// <returns>Tuple of (width, height, minX, minY, maxX, maxY, origin) or null if calculation failed</returns>
        public static (double width, double height, double minX, double minY, double maxX, double maxY, XYZ origin)? CalculateFromCorners(
            List<dynamic> cluster,
            double rotationAngle,
            out XYZ origin,
            Func<int, string, dynamic> getClashZoneBySleeveInstanceId,
            string xmlFilePath = null)
        {
            origin = XYZ.Zero;

            try
            {
                // ✅ CRASH-SAFE: Validate inputs
                if (cluster == null || cluster.Count == 0 || getClashZoneBySleeveInstanceId == null)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CornerBasedBoundingBoxCalculator] Invalid inputs - returning null");
                    return null;
                }

                // Step 1: Collect sleeve data (center, width, height, rotation, pre-calculated corners) from ClashZone
                // ✅ OPTIMIZATION: Use pre-calculated corners from database (dump once use many times)
                var sleeveDataList = new List<(ClashZone cz, XYZ center, double width, double height, double sleeveRotation, double? cosSleeve, double? sinSleeve, XYZ[] preCalculatedCorners)>();

                foreach (var sleeveData in cluster)
                {
                    var clashZone = getClashZoneBySleeveInstanceId(sleeveData.SleeveInstanceId, xmlFilePath);
                    if (clashZone == null) continue;

                    var cz = clashZone as ClashZone;
                    if (cz == null) continue;

                    // ✅ Get sleeve center from Active document coordinates (where sleeve is actually placed)
                    XYZ center = new XYZ(
                        cz.SleevePlacementPointActiveDocumentX,
                        cz.SleevePlacementPointActiveDocumentY,
                        cz.SleevePlacementPointActiveDocumentZ
                    );

                    // Get sleeve dimensions
                    double sleeveWidth = cz.SleeveWidth > 0 ? cz.SleeveWidth : 0;
                    double sleeveHeight = cz.SleeveHeight > 0 ? cz.SleeveHeight : 0;

                    // Get sleeve rotation angle
                    double sleeveRotation = cz.MepElementRotationAngle;

                    // ✅ ROTATION MATRIX: Use pre-calculated cos/sin from database (dump once use many times)
                    double? cosSleeve = cz.MepRotationCos;
                    double? sinSleeve = cz.MepRotationSin;

                    // ✅ PRE-CALCULATED CORNERS: Load from database (dump once use many times)
                    XYZ[] preCalculatedCorners = null;
                    if (cz.SleeveCorner1X.HasValue && cz.SleeveCorner1Y.HasValue && cz.SleeveCorner1Z.HasValue &&
                        cz.SleeveCorner2X.HasValue && cz.SleeveCorner2Y.HasValue && cz.SleeveCorner2Z.HasValue &&
                        cz.SleeveCorner3X.HasValue && cz.SleeveCorner3Y.HasValue && cz.SleeveCorner3Z.HasValue &&
                        cz.SleeveCorner4X.HasValue && cz.SleeveCorner4Y.HasValue && cz.SleeveCorner4Z.HasValue)
                    {
                        preCalculatedCorners = new XYZ[]
                        {
                            new XYZ(cz.SleeveCorner1X.Value, cz.SleeveCorner1Y.Value, cz.SleeveCorner1Z.Value),  // Corner 1: Bottom-left
                            new XYZ(cz.SleeveCorner2X.Value, cz.SleeveCorner2Y.Value, cz.SleeveCorner2Z.Value),  // Corner 2: Bottom-right
                            new XYZ(cz.SleeveCorner3X.Value, cz.SleeveCorner3Y.Value, cz.SleeveCorner3Z.Value),  // Corner 3: Top-left
                            new XYZ(cz.SleeveCorner4X.Value, cz.SleeveCorner4Y.Value, cz.SleeveCorner4Z.Value)   // Corner 4: Top-right
                        };
                    }

                    sleeveDataList.Add((cz, center, sleeveWidth, sleeveHeight, sleeveRotation, cosSleeve, sinSleeve, preCalculatedCorners));
                }

                if (sleeveDataList.Count == 0)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CornerBasedBoundingBoxCalculator] No sleeve data found - returning null");
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ❌ CornerBasedBoundingBoxCalculator: No sleeve data found (cluster.Count={cluster?.Count ?? 0})\n");
                    return null;
                }
                
                // ✅ DIAGNOSTIC: Log how many sleeves have pre-calculated corners
                int sleevesWithCorners = sleeveDataList.Count(s => s.preCalculatedCorners != null);
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] CornerBasedBoundingBoxCalculator: {sleevesWithCorners}/{sleeveDataList.Count} sleeves have pre-calculated corners\n");

                // Step 2: Choose reference point (first sleeve center as origin)
                origin = sleeveDataList[0].center;

                // Step 3: Pre-calculate cluster rotation matrix components
                // ⚠️ CRITICAL: rotationAngle is the cluster's INTENDED ROTATED AXIS, not average of sleeve angles
                var rotationMatrix = RotationMatrixCalculator.CreateRotationMatrix(rotationAngle);
                if (rotationMatrix == null)
                {
                    SafeFileLogger.SafeAppendText("geometry_errors.log",
                        $"[CornerBasedBoundingBoxCalculator] Failed to create rotation matrix for angle {rotationAngle * 180.0 / Math.PI:F2}° - returning null");
                    return null;
                }
                double cosCluster = rotationMatrix.Value.cos;
                double sinCluster = rotationMatrix.Value.sin;

                // Step 4: For each sleeve, use pre-calculated corners OR recalculate if missing
                var allTransformedCorners = new List<XYZ>();

                // ✅ OPTIMIZATION: Pre-calculate corner offset pattern (for fallback recalculation only)
                var cornerOffsetPattern = new[]
                {
                    (-1.0, -1.0),  // Bottom-left
                    (1.0, -1.0),   // Bottom-right
                    (-1.0, 1.0),   // Top-left
                    (1.0, 1.0)     // Top-right
                };

                for (int i = 0; i < sleeveDataList.Count; i++)
                {
                    var (cz, center, sleeveWidth, sleeveHeight, sleeveRotation, cosSleevePreCalc, sinSleevePreCalc, preCalculatedCorners) = sleeveDataList[i];

                    XYZ[] worldCorners;

                    // ✅ USE PRE-CALCULATED CORNERS: If available, use them directly (dump once use many times)
                    if (preCalculatedCorners != null)
                    {
                        worldCorners = preCalculatedCorners;
                    }
                    else
                    {
                        // ✅ FALLBACK: Recalculate corners if pre-calculated ones are missing
                        // Step 4a: Calculate 4 corners of sleeve in its LOCAL coordinate system (before sleeve rotation)
                        double halfW = sleeveWidth / 2.0;
                        double halfH = sleeveHeight / 2.0;

                        var localCorners = new XYZ[4];
                        for (int cornerIdx = 0; cornerIdx < 4; cornerIdx++)
                        {
                            localCorners[cornerIdx] = new XYZ(
                                cornerOffsetPattern[cornerIdx].Item1 * halfW,
                                cornerOffsetPattern[cornerIdx].Item2 * halfH,
                                0
                            );
                        }

                        // Step 4b: Rotate corners by sleeve's rotation angle to get world-space corners
                        // ✅ ROTATION MATRIX: Use pre-calculated cos/sin from database if available, otherwise calculate
                        double cosSleeve = cosSleevePreCalc ?? Math.Cos(sleeveRotation);
                        double sinSleeve = sinSleevePreCalc ?? Math.Sin(sleeveRotation);

                        worldCorners = new XYZ[4];
                        for (int j = 0; j < 4; j++)
                        {
                            double localX = localCorners[j].X;
                            double localY = localCorners[j].Y;

                            // Rotate corner by sleeve rotation
                            double worldX = localX * cosSleeve - localY * sinSleeve;
                            double worldY = localX * sinSleeve + localY * cosSleeve;

                            // Translate to sleeve center
                            worldCorners[j] = new XYZ(
                                center.X + worldX,
                                center.Y + worldY,
                                center.Z
                            );
                        }
                    }

                    // Step 4c: Transform world-space corners to cluster's INTENDED ROTATED AXIS coordinate system
                    for (int j = 0; j < 4; j++)
                    {
                        // Translate relative to origin
                        double relX = worldCorners[j].X - origin.X;
                        double relY = worldCorners[j].Y - origin.Y;

                        // Rotate to cluster's intended axis coordinate system
                        double clusterX = relX * cosCluster - relY * sinCluster;
                        double clusterY = relX * sinCluster + relY * cosCluster;

                        allTransformedCorners.Add(new XYZ(
                            clusterX,
                            clusterY,
                            worldCorners[j].Z
                        ));
                    }
                }

                // Step 5: Find min/max extents of all transformed corners
                double minX = allTransformedCorners.Min(p => p.X);
                double minY = allTransformedCorners.Min(p => p.Y);
                double maxX = allTransformedCorners.Max(p => p.X);
                double maxY = allTransformedCorners.Max(p => p.Y);
                double width = maxX - minX;
                double height = maxY - minY;

                return (width, height, minX, minY, maxX, maxY, origin);
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("geometry_errors.log",
                    $"[CornerBasedBoundingBoxCalculator] Exception in CalculateFromCorners: {ex.Message}, StackTrace: {ex.StackTrace}");
                return null;
            }
        }
    }
}

