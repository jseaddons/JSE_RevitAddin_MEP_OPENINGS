using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Placement
{
    /// <summary>
    /// ✅ SRP COMPLIANCE: Dedicated service for calculating cluster placement coordinates.
    /// Extracted from ClusterRotationService to separate placement logic from rotation/sizing logic.
    /// Handles category-specific overrides (e.g., Dampers), hybrid wall logic, and DB-first lookup.
    /// </summary>
    public class ClusterPlacementCalculationService : IClusterPlacementCalculationService
    {
        private readonly Func<int, string, ClashZone> _getClashZoneFunc;
        private readonly Func<int, XYZ> _getClusterPlacementFunc;
        
        // Thread-safe set to track processed cluster IDs and avoid stacking
        private static readonly ConcurrentDictionary<int, byte> _processedClusterIds = new ConcurrentDictionary<int, byte>();

        public ClusterPlacementCalculationService(
            Func<int, string, ClashZone> getClashZoneFunc,
            Func<int, XYZ> getClusterPlacementFunc = null)
        {
            _getClashZoneFunc = getClashZoneFunc ?? throw new ArgumentNullException(nameof(getClashZoneFunc));
            _getClusterPlacementFunc = getClusterPlacementFunc;
        }

        /// <summary>
        /// ✅ CRITICAL: Authoritative placement coordinate determination.
        /// Priority:
        /// 1. DB LOOKUP: If cluster already exists in DB, trust its stored placement (Source of Truth).
        /// 2. DAMPER OVERRIDE: Strictly use intersection centroid for dampers to fix wall alignment.
        /// 3. HYBRID WALL LOGIC: For non-damper walls, use First Sleeve for thickness axis and Centroid for width/height.
        /// 4. FLOOR/FALLBACK: Use geometric center or centroid.
        /// </summary>
        public XYZ CalculatePlacementPoint(
            List<dynamic> cluster, 
            double bboxWidth, 
            double bboxHeight, 
            double bboxDepth,
            XYZ? rotatedMin = null,
            XYZ? rotatedMax = null,
            XYZ? explicitCenter = null,
            string? xmlFilePath = null)
        {
            if (cluster == null || cluster.Count == 0)
                return XYZ.Zero;

            // ✅ PHASE 1: Try DB Lookup (Authoritative Source of Truth)
            var dbPoint = TryGetDatabasePlacement(cluster, xmlFilePath);
            if (dbPoint != null) return dbPoint;

            // ✅ PHASE 2: Dispatch by Host Type
            var firstCz = GetFirstClashZone(cluster, xmlFilePath);
            if (firstCz == null) return explicitCenter ?? CalculateFromIntersections(cluster, xmlFilePath);

            bool isFloor = (firstCz.StructuralElementType ?? "").IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0;

            if (isFloor)
            {
                return CalculateFloorPlacementPoint(cluster, explicitCenter, xmlFilePath);
            }
            else
            {
                // Wall or Structural Framing
                return CalculateWallFramingPlacementPoint(cluster, firstCz, explicitCenter, xmlFilePath);
            }
        }

        private XYZ? TryGetDatabasePlacement(List<dynamic> cluster, string? xmlFilePath)
        {
            if (_getClusterPlacementFunc == null || cluster[0] == null) return null;

            try
            {
                int clusterId = GetClusterInstanceId(cluster[0], xmlFilePath);
                if (clusterId > 0 && _processedClusterIds.TryAdd(clusterId, 0))
                {
                    var dbPlacement = _getClusterPlacementFunc(clusterId);
                    if (dbPlacement != null)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🎯 DB SOURCE OF TRUTH: Using stored placement for Cluster {clusterId}: ({dbPlacement.X:F4}, {dbPlacement.Y:F4}, {dbPlacement.Z:F4})\n");
                        return dbPlacement;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    SafeFileLogger.SafeAppendText("cluster_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ DB Lookup Failed: {ex.Message}\n");
            }
            return null;
        }

        /// <summary>
        /// ✅ SPECIALIZED FLOOR LOGIC: Strictly separated from wall logic.
        /// Floors use Centroid or Explicit Center (BBox Geometric Center). No corner-based logic.
        /// </summary>
        private XYZ CalculateFloorPlacementPoint(List<dynamic> cluster, XYZ? explicitCenter, string? xmlFilePath)
        {
            XYZ result = explicitCenter ?? CalculateFromIntersections(cluster, xmlFilePath);
            
            if (!DeploymentConfiguration.DeploymentMode)
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] 🎯 FLOOR PLACEMENT: Using {(explicitCenter != null ? "Geometric Center" : "Centroid")}: ({result.X:F4}, {result.Y:F4}, {result.Z:F4})\n");
            
            return result;
        }

        /// <summary>
        /// ✅ SPECIALIZED WALL/FRAMING LOGIC: Strictly separated from floor logic.
        /// Uses Corners for non-dampers, and Centroid + WallAlignment for dampers.
        /// </summary>
        private XYZ CalculateWallFramingPlacementPoint(List<dynamic> cluster, ClashZone firstCz, XYZ? explicitCenter, string? xmlFilePath)
        {
            XYZ placementPoint;
            bool isDamperCluster = CheckIfDamperCluster(cluster, xmlFilePath);

            // ✅ USER RULE (2026-02-04):
            // For ALL wall / structural framing clusters (including dampers), placement must be
            // driven purely by sleeve corners:
            // - Width: extreme left/right corners along wall axis (Y for Y‑wall, X for X‑wall)
            // - Height: extreme top/bottom corners (Z)
            // - Placement point: midpoint of those extremes, with thickness axis snapped to
            //   the first sleeve's wall centerline.
            //
            // Centroid-based placement is NO LONGER used on walls/framing, even for dampers.
            // That is exactly what CalculateFromCornersForWallCluster already implements.
            placementPoint = CalculateFromCornersForWallCluster(cluster, firstCz, xmlFilePath);

            if (!DeploymentConfiguration.DeploymentMode)
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] 🎯 WALL/FRAMING PLACEMENT (Damper={isDamperCluster}, Hybrid=True): ({placementPoint.X:F4}, {placementPoint.Y:F4}, {placementPoint.Z:F4})\n");

            return placementPoint;
        }

        private XYZ ApplyDamperAlignment(XYZ currentPoint, List<dynamic> cluster, string? xmlFilePath)
        {
            if (currentPoint == null) return XYZ.Zero;

            bool isDamper = CheckIfDamperCluster(cluster, xmlFilePath);
            if (!isDamper) return currentPoint;
            if (cluster == null || cluster.Count == 0) return currentPoint;

            // ✅ Duct Accessories on walls now use same corner-based placement as Ducts (Phase 2); do not overwrite with WallCenterline snap
            ClashZone firstCz = null;
            var firstItem = cluster[0];
            
            if (firstItem is ClashZone cz)
                firstCz = cz;
            else
            {
                try { firstCz = firstItem.ClashZone; } catch { }
            }
            
            if (firstCz == null)
            {
                int firstId = GetSleeveId(firstItem);
                if (firstId > 0)
                {
                    firstCz = _getClashZoneFunc(firstId, xmlFilePath);
                }
            }
            
            // ✅ Skip WallCenterline snap for wall-hosted Duct Accessories: they now use same corner-based placement as Ducts (Phase 2)
            if (firstCz != null && IsDamperCategory(firstCz.MepElementCategory) && IsWallHosted(firstCz))
            {
                return currentPoint;
            }
            
            return currentPoint;
        }

        /// <summary>
        /// ✅ USER-REQUESTED WALL LOGIC:
        /// Use ONLY sleeve corners to compute cluster placement for wall-hosted, non-damper clusters.
        /// - Width  axis: extreme left/right corners of all sleeves (along wall length).
        /// - Height axis: extreme top/bottom corners of all sleeves (Z).
        /// - Thickness axis: first sleeve's wall centerline / intersection (keeps opening in wall core).
        /// </summary>
        private XYZ CalculateFromCornersForWallCluster(List<dynamic> cluster, ClashZone firstCz, string? xmlFilePath)
        {
            if (cluster == null || cluster.Count == 0 || firstCz == null)
                return XYZ.Zero;

            double minX = double.PositiveInfinity, maxX = double.NegativeInfinity;
            double minY = double.PositiveInfinity, maxY = double.NegativeInfinity;
            double minZ = double.PositiveInfinity, maxZ = double.NegativeInfinity;
            bool hasCorner = false;

            int sleeveIdx = 0;
            foreach (var sleeve in cluster)
            {
                sleeveIdx++;
                ClashZone cz = null;

                if (sleeve is ClashZone directCz)
                    cz = directCz;
                else
                {
                    try { cz = sleeve.ClashZone; } catch { }
                }

                if (cz == null)
                {
                    int id = GetSleeveId(sleeve);
                    if (id > 0)
                    {
                        cz = _getClashZoneFunc(id, xmlFilePath);
                    }
                }

                if (cz == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"    [{sleeveIdx}] ⚠️ CORNERS: ClashZone is NULL, skipping.\n");
                    continue;
                }

                double[] xs = { cz.SleeveCorner1X ?? 0, cz.SleeveCorner2X ?? 0, cz.SleeveCorner3X ?? 0, cz.SleeveCorner4X ?? 0 };
                double[] ys = { cz.SleeveCorner1Y ?? 0, cz.SleeveCorner2Y ?? 0, cz.SleeveCorner3Y ?? 0, cz.SleeveCorner4Y ?? 0 };
                double[] zs = { cz.SleeveCorner1Z ?? 0, cz.SleeveCorner2Z ?? 0, cz.SleeveCorner3Z ?? 0, cz.SleeveCorner4Z ?? 0 };

                // Consider corners invalid only if *all* coordinates are zero
                bool allZero = xs.All(v => v == 0) && ys.All(v => v == 0) && zs.All(v => v == 0);
                if (allZero)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"    [{sleeveIdx}] ⚠️ CORNERS: All corner coordinates are zero, skipping.\n");
                    continue;
                }

                hasCorner = true;

                minX = Math.Min(minX, xs.Min());
                maxX = Math.Max(maxX, xs.Max());
                minY = Math.Min(minY, ys.Min());
                maxY = Math.Max(maxY, ys.Max());
                minZ = Math.Min(minZ, zs.Min());
                maxZ = Math.Max(maxZ, zs.Max());
            }

            if (!hasCorner || double.IsInfinity(minX) || double.IsInfinity(minZ))
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"    ⚠️ CORNERS: No valid corners found, falling back to intersection centroid.\n");
                return CalculateFromIntersections(cluster, xmlFilePath);
            }

            string orientation = (firstCz.HostOrientation ?? "").ToUpper();

            // Midpoints from corner extents
            double centerX = (minX + maxX) / 2.0;
            double centerY = (minY + maxY) / 2.0;
            double centerZ = (minZ + maxZ) / 2.0;

            XYZ result;

            if (orientation.Contains("X"))
            {
                // X-wall: Length along X, thickness along Y.
                // - Width  (X) and Height (Z) from cluster corner midpoints.
                // - Thickness (Y) locked to first instance's wall centerline point.
                double yCenterline = firstCz.WallCenterlinePointY != 0 ? firstCz.WallCenterlinePointY : firstCz.IntersectionPointY;
                result = new XYZ(centerX, yCenterline, centerZ);
            }
            else if (orientation.Contains("Y"))
            {
                // Y-wall: Length along Y, thickness along X.
                // - Width  (Y) and Height (Z) from cluster corner midpoints.
                // - Thickness (X) locked to first instance's wall centerline point.
                double xCenterline = firstCz.WallCenterlinePointX != 0 ? firstCz.WallCenterlinePointX : firstCz.IntersectionPointX;
                result = new XYZ(xCenterline, centerY, centerZ);
            }
            else
            {
                // Fallback: use full 3D midpoint if orientation is unknown
                result = new XYZ(centerX, centerY, centerZ);
            }

            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"    🎯 CORNER-BASED PLACEMENT (WALL): " +
                    $"X=[{minX:F6},{maxX:F6}], Y=[{minY:F6},{maxY:F6}], Z=[{minZ:F6},{maxZ:F6}] -> " +
                    $"Center=({result.X:F6}, {result.Y:F6}, {result.Z:F6})\n");
            }

            return result;
        }

        private XYZ CalculateFromIntersections(List<dynamic> cluster, string? xmlFilePath)
        {
            double sumX = 0, sumY = 0, sumZ = 0;
            int validCount = 0;
            ClashZone? firstValidCz = null;
            bool isDamper = false;

            int sleeveIdx = 0;
            foreach (var sleeve in cluster)
            {
                sleeveIdx++;
                ClashZone cz = null;
                
                if (sleeve is ClashZone)
                    cz = sleeve as ClashZone;
                else
                {
                    try { cz = sleeve.ClashZone; } catch { }
                }

                if (cz == null)
                {
                    int id = GetSleeveId(sleeve);
                    if (id > 0)
                    {
                        cz = _getClashZoneFunc(id, xmlFilePath);
                    }
                }

                if (cz == null) 
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log", $"    [{sleeveIdx}] ⚠️ SKIPPING: ClashZone is NULL\n");
                    continue;
                }

                if (firstValidCz == null) firstValidCz = cz;
                
                if (IsDamperCategory(cz.MepElementCategory)) isDamper = true;

                if (IsValidPoint(cz.IntersectionPointX, cz.IntersectionPointY, cz.IntersectionPointZ))
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                        $"    [{sleeveIdx}] GUID={cz.ClashZoneGuid}: Intersection=({cz.IntersectionPointX:F4}, {cz.IntersectionPointY:F4}, {cz.IntersectionPointZ:F4})\n");
                    
                    sumX += cz.IntersectionPointX;
                    sumY += cz.IntersectionPointY;
                    sumZ += cz.IntersectionPointZ;
                    validCount++;
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                        $"    [{sleeveIdx}] GUID={cz.ClashZoneGuid}: ⚠️ INVALID Intersection Point\n");
                }
            }

            if (validCount == 0 || firstValidCz == null) 
            {
                SafeFileLogger.SafeAppendText("cluster_sizing.log", "    ❌ FAILED: No valid intersection points found for averaging.\n");
                return XYZ.Zero;
            }

            XYZ centroid = new XYZ(sumX / validCount, sumY / validCount, sumZ / validCount);
            SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                $"    🧮 MATH: SumX={sumX:F4}, SumY={sumY:F4}, SumZ={sumZ:F4}, Count={validCount}\n" +
                $"    🎯 COMPUTED CENTROID: ({centroid.X:F4}, {centroid.Y:F4}, {centroid.Z:F4})\n");

            if (isDamper)
            {
                // ✅ LATERAL SHIFT FIX: Use WallCenterline for alignment
                if (IsWallHosted(firstValidCz))
                {
                    double cX = centroid.X;
                    double cY = centroid.Y;
                    double cZ = centroid.Z;

                    string orientation = (firstValidCz.HostOrientation ?? "").ToUpper();
                    if (orientation.Contains("Y")) // Normal along X, Thickness along X
                    {
                        cX = firstValidCz.WallCenterlinePointX;
                         if (!DeploymentConfiguration.DeploymentMode)
                             SafeFileLogger.SafeAppendText("cluster_sizing.log", $"[{DateTime.Now:HH:mm:ss}] 🎯 PLACEMENT (DAMPER): Y-Wall Fixed X={cX:F4}, Centroid Y={cY:F4}, Z={cZ:F4}\n");
                    }
                    else if (orientation.Contains("X")) // Normal along Y, Thickness along Y
                    {
                        cY = firstValidCz.WallCenterlinePointY;
                        if (!DeploymentConfiguration.DeploymentMode)
                             SafeFileLogger.SafeAppendText("cluster_sizing.log", $"[{DateTime.Now:HH:mm:ss}] 🎯 PLACEMENT (DAMPER): X-Wall Fixed Y={cY:F4}, Centroid X={cX:F4}, Z={cZ:F4}\n");
                    }
                    
                    XYZ placementPoint = new XYZ(cX, cY, cZ);

                    // ✅ Z-AXIS DIAGNOSTIC: Log actual Z-coordinates to identify shift source
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        // For walls, we want to know the vertical range (Z)
                        // SumZ / validCount is our centroid Z
                        SafeFileLogger.SafeAppendText("cluster_placement_debug.log", 
                            $"[{DateTime.Now:HH:mm:ss}] 🎯 Z-AXIS DIAGNOSTIC (WALL CENTROID): " +
                            $"CentroidZ={cZ:F6}, PlacedZ={cZ:F6}\n");
                    }

                    return placementPoint;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                    SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                        $"[{DateTime.Now:HH:mm:ss}] 🎯 PLACEMENT (DAMPER): Using strict intersection centroid ({centroid.X:F6}, {centroid.Y:F6}, {centroid.Z:F6})\n");
                return centroid;
            }

            // Hybrid Wall Logic
            if (IsWallHosted(firstValidCz))
            {
                string orientation = (firstValidCz.HostOrientation ?? "").ToUpper();
                if (orientation.Contains("X")) 
                    return new XYZ(centroid.X, firstValidCz.IntersectionPointY, centroid.Z);
                if (orientation.Contains("Y")) 
                    return new XYZ(firstValidCz.IntersectionPointX, centroid.Y, centroid.Z);
            }

            return centroid;
        }

        private int GetSleeveId(dynamic sleeve)
        {
            try { return sleeve.SleeveInstanceId; } catch { return 0; }
        }

        private ClashZone? GetFirstClashZone(List<dynamic> cluster, string? xmlPath)
        {
            if (cluster == null || cluster.Count == 0)
                return null;

            var firstItem = cluster[0];
            ClashZone cz = null;

            if (firstItem is ClashZone directCz)
                cz = directCz;
            else
            {
                try { cz = firstItem.ClashZone; } catch { }
            }

            if (cz != null)
                return cz;

            int id = GetSleeveId(firstItem);
            if (id > 0)
            {
                return _getClashZoneFunc(id, xmlPath);
            }

            return null;
        }

        private int GetClusterInstanceId(dynamic sleeve, string? xmlPath)
        {
            try { return sleeve.ClusterInstanceId; } 
            catch { 
                try { return sleeve.ClashZone.ClusterInstanceId; }
                catch { return _getClashZoneFunc(GetSleeveId(sleeve), xmlPath)?.ClusterInstanceId ?? -1; }
            }
        }

        private bool IsDamperCategory(string category) => 
            (category ?? "").ToLower().Contains("duct accessory") || (category ?? "").ToLower().Contains("duct accessories");

        private bool CheckIfDamperCluster(List<dynamic> cluster, string? xmlPath) =>
            cluster.Any(s => IsDamperCategory(_getClashZoneFunc(GetSleeveId(s), xmlPath)?.MepElementCategory));

        private bool IsWallHosted(ClashZone cz) => 
            (cz.StructuralElementType ?? "").StartsWith("Wall", StringComparison.OrdinalIgnoreCase) || 
            (cz.StructuralElementType ?? "").Equals("Structural Framing", StringComparison.OrdinalIgnoreCase);

        private bool IsValidPoint(double x, double y, double z) => 
            (x != 0 || y != 0 || z != 0) && !double.IsNaN(x) && !double.IsInfinity(x);
    }
}
