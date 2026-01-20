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

            XYZ? placementPoint = null;

            // ✅ PHASE 1: Try DB Lookup (Authoritative Source of Truth)
            if (_getClusterPlacementFunc != null && cluster[0] != null)
            {
                try
                {
                    int clusterId = GetClusterInstanceId(cluster[0], xmlFilePath);
                    
                    if (clusterId > 0)
                    {
                        // Deduplication: Only use DB placement once per run
                        if (_processedClusterIds.TryAdd(clusterId, 0))
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
                }
                catch (Exception ex)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_errors.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ DB Lookup Failed: {ex.Message}\n");
                }
            }

            // ✅ PHASE 2: Calculate from Intersections if DB lookup skipped or failed
            // Usage of explicitCenter (Geometric Center from Rotation Service) if available
            placementPoint = explicitCenter ?? CalculateFromIntersections(cluster, xmlFilePath);

            // ✅ PHASE 2b: Apply Alignment Constraints (Lateral Shift Fix / Wall Centerline)
            // This applies to both calculated and explicitly passed centers
            placementPoint = ApplyDamperAlignment(placementPoint, cluster, xmlFilePath);

            // ✅ PHASE 3: Re-centering / Geometric Fallback (Category Dependent)
            // For circular elements or mixed sizes, adjusting the point based on the bounding box range.
            // DO NOT OVERWRITE if explicitCenter was provided (it is the authoritative Geometric Center).
            if (explicitCenter == null && rotatedMin != null && rotatedMax != null)
            {
                bool isDamper = CheckIfDamperCluster(cluster, xmlFilePath);
                
                if (!isDamper)
                {
                    // For non-dampers (Pipes/Ducts), we restore geometric center centering
                    // NOTE: If explicitCenter (Geometric Center) was passed, this might be redundant or slightly different?
                    // But for consistency with legacy flow, we keep this re-centering logic for non-dampers.
                    placementPoint = new XYZ(
                        (rotatedMin.X + rotatedMax.X) / 2.0,
                        (rotatedMin.Y + rotatedMax.Y) / 2.0,
                        (rotatedMin.Z + rotatedMax.Z) / 2.0
                    );
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] 🎯 PLACEMENT (DEFAULT): Using geometric center (re-centered): ({placementPoint.X:F2}, {placementPoint.Y:F2}, {placementPoint.Z:F2})\n");
                }
            }
 
            XYZ finalPoint = placementPoint ?? XYZ.Zero;
            SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                $"    🏁 FINAL PLACEMENT POINT: ({finalPoint.X:F6}, {finalPoint.Y:F6}, {finalPoint.Z:F6})\n\n");
            return finalPoint;
        }

        private XYZ ApplyDamperAlignment(XYZ currentPoint, List<dynamic> cluster, string? xmlFilePath)
        {
            if (currentPoint == null) return XYZ.Zero;

            bool isDamper = CheckIfDamperCluster(cluster, xmlFilePath);
            if (!isDamper) return currentPoint;
            if (cluster == null || cluster.Count == 0) return currentPoint;

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
            
            if (firstCz != null && IsDamperCategory(firstCz.MepElementCategory) && IsWallHosted(firstCz))
            {
                double cX = currentPoint.X;
                double cY = currentPoint.Y;
                double cZ = currentPoint.Z;

                string orientation = (firstCz.HostOrientation ?? "").ToUpper();
                
                if (orientation.Contains("Y")) // Normal along X, Thickness along X
                {
                    double oldX = cX;
                    cX = firstCz.WallCenterlinePointX;
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                            $"    🔄 SHIFT REASON (DAMPER/WALL): Smapped X to Wall Core. Changed {oldX:F4} -> {cX:F4} (Orientation=Y-Wall)\n");
                }
                else if (orientation.Contains("X")) // Normal along Y, Thickness along Y
                {
                    double oldY = cY;
                    cY = firstCz.WallCenterlinePointY;
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                            $"    🔄 SHIFT REASON (DAMPER/WALL): Snapped Y to Wall Core. Changed {oldY:F4} -> {cY:F4} (Orientation=X-Wall)\n");
                }
                
                return new XYZ(cX, cY, cZ);
            }
            
            return currentPoint;
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
                    
                    return new XYZ(cX, cY, cZ);
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
