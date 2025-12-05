using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation.Services
{
    /// <summary>
    /// ✅ SOLID SRP: Service responsible ONLY for calculating cluster placement points
    /// Single Responsibility: Calculate optimal placement point for cluster sleeve
    /// </summary>
    public class ClusterPlacementPointService : IClusterPlacementPointService
    {
        private readonly IClusterRotationCacheService _cacheService;

        public ClusterPlacementPointService(IClusterRotationCacheService cacheService)
        {
            _cacheService = cacheService ?? throw new ArgumentNullException(nameof(cacheService));
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Calculate cluster placement point from intersection points (centroid)
        /// Individual sleeves are placed at intersection points, so cluster should be at average of intersection points
        /// </summary>
        public XYZ CalculatePlacementPoint(List<dynamic> cluster, string? xmlFilePath = null)
        {
            if (cluster == null || cluster.Count == 0)
                return XYZ.Zero;

            double sumX = 0.0;
            double sumY = 0.0;
            double sumZ = 0.0;
            int validCount = 0;

            foreach (var sleeveData in cluster)
            {
                try
                {
                    int sleeveInstanceId = sleeveData.SleeveInstanceId;
                    if (sleeveInstanceId <= 0)
                        continue;

                    var cz = _cacheService.GetCachedClashZone(sleeveInstanceId, xmlFilePath);
                    if (cz == null)
                        continue;

                    // ✅ Use intersection point coordinates (where MEP element intersects structural element)
                    double ipX = cz.IntersectionPointX;
                    double ipY = cz.IntersectionPointY;
                    double ipZ = cz.IntersectionPointZ;

                    // Validate coordinates are non-zero and not NaN
                    if (ipX != 0.0 || ipY != 0.0 || ipZ != 0.0)
                    {
                        if (!double.IsNaN(ipX) && !double.IsInfinity(ipX) &&
                            !double.IsNaN(ipY) && !double.IsInfinity(ipY) &&
                            !double.IsNaN(ipZ) && !double.IsInfinity(ipZ))
                        {
                            sumX += ipX;
                            sumY += ipY;
                            sumZ += ipZ;
                            validCount++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log error but continue with other sleeves
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error getting intersection point for sleeve in cluster: {ex.Message}\n");
                    }
                }
            }

            if (validCount == 0)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ No valid intersection points found in cluster ({cluster.Count} sleeves)\n");
                }
                return XYZ.Zero;
            }

            // Calculate centroid (average) of intersection points
            XYZ centroid = new XYZ(sumX / validCount, sumY / validCount, sumZ / validCount);

            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ✅ PLACEMENT POINT: Calculated from {validCount}/{cluster.Count} intersection points: ({centroid.X:F6}, {centroid.Y:F6}, {centroid.Z:F6})\n");
            }

            return centroid;
        }
    }
}

