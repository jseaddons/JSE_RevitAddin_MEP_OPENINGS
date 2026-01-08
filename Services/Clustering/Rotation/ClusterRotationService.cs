using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation
{
    /// <summary>
    /// Service for handling cluster rotation calculations and rotated bounding box determinations
    /// Extracted from UniversalClusterService for better separation of concerns
    /// </summary>
    public class ClusterRotationService : IClusterRotationService
    {
        // ✅ THREAD-SAFETY: Use ThreadLocal storage for parallel processing isolation
        // Each thread gets its own isolated dictionary to prevent race conditions during parallel execution
        
        // Store rotation angle and rotated bounding box for each cluster sleeve
        private readonly System.Threading.ThreadLocal<Dictionary<int, (double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)>> _clusterRotationData;

        // ✅ PERFORMANCE: Cache rotated bounding box calculations by cluster signature + rotation angle
        private readonly System.Threading.ThreadLocal<Dictionary<string, (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)>> _rotatedBboxCache;
        private const int MAX_ROTATED_BBOX_CACHE_SIZE = 1000; // Limit cache size to prevent memory growth

        // ✅ PERFORMANCE: Cache individual ClashZone lookups to avoid repeated database queries for same sleeve
        private readonly System.Threading.ThreadLocal<Dictionary<int, ClashZone>> _clashZoneCache;
        private const int MAX_CLASHZONE_CACHE_SIZE = 5000; // Limit cache size to prevent memory growth

        // Delegate for getting ClashZone by sleeve instance ID (injected dependency)
        private readonly Func<int, string, ClashZone> _getClashZoneFunc;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="getClashZoneFunc">Function to retrieve ClashZone by sleeve instance ID</param>
        public ClusterRotationService(Func<int, string, ClashZone> getClashZoneFunc)
        {
            // ✅ VALIDATION: Check for null before assignment (better error message)
            if (getClashZoneFunc == null)
            {
                throw new ArgumentNullException(nameof(getClashZoneFunc), 
                    "getClashZoneFunc cannot be null. Provide a function that returns ClashZone (can return null if not found).");
            }
            
            // ✅ THREAD-SAFETY: Initialize ThreadLocal factories
            _clusterRotationData = new System.Threading.ThreadLocal<Dictionary<int, (double, bool, XYZ, XYZ, double, double, double)>>(() => new Dictionary<int, (double, bool, XYZ, XYZ, double, double, double)>());
            _rotatedBboxCache = new System.Threading.ThreadLocal<Dictionary<string, (double, double, double, XYZ, double?, double?, double?, double?, double?, double?)>>(() => new Dictionary<string, (double, double, double, XYZ, double?, double?, double?, double?, double?, double?)>());
            _clashZoneCache = new System.Threading.ThreadLocal<Dictionary<int, ClashZone>>(() => new Dictionary<int, ClashZone>());
            
            _getClashZoneFunc = getClashZoneFunc;
        }

        /// <summary>
        /// Determine the dominant rotation angle for a cluster of sleeves
        /// Returns 0 for straight axis-aligned clusters (aligned to WCS: 0°, 90°, 180°, 270°)
        /// Returns rotation angle for rotated axis-aligned clusters (non-straight: 45°, 225°, etc.)
        /// 
        /// ✅ CRITICAL: Circular elements (Pipes and Round Ducts) always return 0.0° - no rotation needed
        ///    MEP orientation is meaningless for circular elements, so always use straight axis (0°)
        ///    This applies to both floors and walls - circular elements don't need rotation alignment
        /// </summary>
        public double DetermineRotationAngle(List<dynamic> cluster, string? xmlFilePath = null)
        {
            try
            {
                if (cluster == null || cluster.Count == 0)
                    return 0.0;

                // ✅ CRITICAL: Check if cluster contains circular elements (Pipes and Round Ducts)
                // Circular elements should always be placed straight to WCS (axis-aligned), no rotation needed
                // MEP orientation is meaningless for circular elements - they have the same dimensions in all directions
                // This applies to BOTH floors and walls - circular elements don't need rotation for alignment
                bool isCircularElementCluster = false;
                string circularElementType = "";
                foreach (var sleeveData in cluster)
                {
                    if (sleeveData?.ClashZone == null)
                        continue;

                    var clashZone = sleeveData.ClashZone as ClashZone;
                    if (clashZone == null)
                        continue;

                    // ✅ DIAGNOSTIC: Log actual category name to identify pipe detection issues
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [PIPE-CATEGORY-DEBUG] Sleeve {sleeveData.SleeveInstanceId}: MepElementCategory='{clashZone.MepElementCategory}'\n");
                    }

                    // ✅ Check if this is a pipe (all pipes are circular)
                    bool isPipe = string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(clashZone.MepElementCategory, "Pipe Accessories", StringComparison.OrdinalIgnoreCase) ||
                                  (clashZone.MepElementCategory != null && clashZone.MepElementCategory.IndexOf("Pipe", StringComparison.OrdinalIgnoreCase) >= 0);
                    
                    if (isPipe)
                    {
                        isCircularElementCluster = true;
                        circularElementType = "PIPE";
                        break;
                    }

                    // ✅ Check if this is a round duct (circular element)
                    bool isDuct = string.Equals(clashZone.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) ||
                                  string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
                    if (isDuct)
                    {
                        // Check if duct is round/circular
                        bool isRoundDuct = string.Equals(clashZone.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(clashZone.DuctShape, "Circular", StringComparison.OrdinalIgnoreCase) ||
                                          (clashZone.MepElementSizeData != null && 
                                           (string.Equals(clashZone.MepElementSizeData.Shape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                            string.Equals(clashZone.MepElementSizeData.Shape, "Circular", StringComparison.OrdinalIgnoreCase)));
                        
                        if (isRoundDuct)
                        {
                            isCircularElementCluster = true;
                            circularElementType = "ROUND DUCT";
                            break;
                        }
                    }
                }

                // ✅ WALL/STRUCTURAL FRAMING: Rotation based on X-wall vs Y-wall (same as individual sleeves)
                // Individual sleeves: X-walls get +90°, Y-walls get 0°
                // Cluster sleeves must match individual sleeve rotation to maintain correct orientation
                // ⚠️⚠️⚠️ CRITICAL FIX: Check wall orientation BEFORE returning 0.0° for circular elements
                // Pipes on X-walls need 90° rotation, pipes on Y-walls need 0° rotation
                // Get host type and orientation from first clash zone
                ClashZone? firstClashZone = null;
                foreach (var sleeveData in cluster)
                {
                    if (sleeveData?.ClashZone == null)
                        continue;

                    firstClashZone = sleeveData.ClashZone as ClashZone;
                    if (firstClashZone != null)
                        break;
                }

                if (firstClashZone != null)
                {
                    // Check if this is a wall or structural framing host
                    bool isWallHost = string.Equals(firstClashZone.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                     string.Equals(firstClashZone.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                    bool isFramingHost = string.Equals(firstClashZone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);

                    if (isWallHost || isFramingHost)
                    {
                        // ✅ WALL ROTATION: Match individual sleeve rotation logic
                        // Individual sleeves: X-walls get +90° rotation, Y-walls get 0° rotation
                        // Cluster sleeves must use the SAME rotation to match individual sleeve orientation
                        // ⚠️⚠️⚠️ CRITICAL FIX: Even circular elements (pipes) need rotation on X-walls
                        string wallDirectionType = firstClashZone.WallDirectionType ?? "";
                        string hostOrientation = firstClashZone.HostOrientation ?? "";
                        
                        // LOGGING FOR USER VERIFICATION
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE-DEBUG] Sleeve {firstClashZone.SleeveInstanceId} WallDirectionType='{wallDirectionType}', HostOrientation='{hostOrientation}'\n");
                            DebugLogger.Info($"[ClusterRotationService] Sleeve {firstClashZone.SleeveInstanceId} WallDirectionType='{wallDirectionType}', HostOrientation='{hostOrientation}'");
                        }
                                                
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE-DEBUG] WALL HOST: StructuralType='{firstClashZone.StructuralElementType}', WallDirectionType='{wallDirectionType}', HostOrientation='{hostOrientation}', IsCircular={isCircularElementCluster}\n");
                        }
                        
                        // Check WallDirectionType or HostOrientation to determine X-wall vs Y-wall
                        bool isXWall = wallDirectionType.Contains("X-WALL", StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase);
                        bool isYWall = wallDirectionType.Contains("Y-WALL", StringComparison.OrdinalIgnoreCase) ||
                                      string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase);
                        
                        if (isXWall)
                        {
                            // ⚠️⚠️⚠️ CRITICAL FIX: X-wall needs 90° rotation even for circular elements (pipes)
                            // Individual sleeves apply 90° rotation for X-walls regardless of element shape
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                string elementType = isCircularElementCluster ? $"{circularElementType} on" : "";
                                DebugLogger.Info($"[CLUSTER-ANGLE] WALL ROTATION: {elementType} {firstClashZone.StructuralElementType} - X-WALL → Returning 90.0° (matches individual sleeve rotation)");
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE] ✅ X-WALL DETECTED → 90.0° (matches individual sleeve, applies to circular elements too)\n");
                            }
                            return Math.PI / 2.0; // 90° rotation for X-walls (matches individual sleeves, even for pipes)
                        }
                        else if (isYWall)
                        {
                            // Y-wall: 0° rotation (matches individual sleeve rotation)
                            // Circular elements on Y-walls also get 0° (no rotation needed)
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                string elementType = isCircularElementCluster ? $"{circularElementType} on" : "";
                                DebugLogger.Info($"[CLUSTER-ANGLE] WALL ROTATION: {elementType} {firstClashZone.StructuralElementType} - Y-WALL → Returning 0.0° (matches individual sleeve rotation)");
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE] ✅ Y-WALL DETECTED → 0.0° (matches individual sleeve)\n");
                            }
                            return 0.0; // No rotation for Y-walls (matches individual sleeves)
                        }
                        else
                        {
                            // Fallback: Default to 0° if cannot determine
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE] ⚠️ WALL HOST but cannot determine X/Y → Defaulting to 0.0°\n");
                            }
                            return 0.0;
                        }
                    }
                }

                // ✅ CIRCULAR ELEMENT FIX: For floors (non-wall hosts), return 0.0° for circular elements (Pipes and Round Ducts)
                // MEP orientation is meaningless for circular elements on floors - they have the same dimensions in all directions
                // ⚠️⚠️⚠️ CRITICAL: This only applies to floors, NOT walls (walls handled above)
                if (isCircularElementCluster)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTER-ANGLE] {circularElementType} cluster on FLOOR detected: Returning 0.0° (straight axis-aligned to WCS, MEP orientation meaningless for circular elements on floors)");
                        SafeFileLogger.SafeAppendText("cluster_debug.log",
                            $"[{DateTime.Now:HH:mm:ss.fff}] [CIRCULAR-RETURN] ✅ {circularElementType} cluster on FLOOR → Returning 0.0° (straight axis, MEP orientation meaningless)\n");
                    }
                    return 0.0;
                }

                var rotationAngles = new List<double>();

                foreach (var sleeveData in cluster)
                {
                    if (sleeveData?.ClashZone == null)
                        continue;

                    var clashZone = sleeveData.ClashZone as ClashZone;
                    if (clashZone == null)
                        continue;

                    // ✅ ROTATION DATA FROM DB: Get rotation angle from clash zone (loaded from database).
                    // MepElementRotationAngle is saved to database during refresh (InsertOrUpdateClashZones)
                    // and loaded by GetClashZonesByCategory -> MepRotationAngleRad column.
                    double angle = clashZone.MepElementRotationAngle;
                    
                    // Normalize angle to 0-2π range
                    while (angle < 0) angle += 2 * Math.PI;
                    while (angle >= 2 * Math.PI) angle -= 2 * Math.PI;
                    
                    rotationAngles.Add(angle);
                }

                if (rotationAngles.Count == 0)
                    return 0.0;

                // ✅ STRATEGY: Use average angle (works well for similar angles)
                double averageAngle = rotationAngles.Average();
                
                // Check if angles wrap around 0°/360° boundary
                double minAngle = rotationAngles.Min();
                double maxAngle = rotationAngles.Max();
                if (maxAngle - minAngle > Math.PI)
                {
                    // Angles wrap around - adjust by adding 2π to angles < π
                    var adjustedAngles = rotationAngles.Select(a => a < Math.PI ? a + 2 * Math.PI : a).ToList();
                    averageAngle = adjustedAngles.Average();
                    if (averageAngle >= 2 * Math.PI)
                        averageAngle -= 2 * Math.PI;
                }

                // ✅ CRITICAL FIX: Check if angle is essentially straight axis-aligned to WCS (0°, 90°, 180°, 270°)
                // If so, return 0 to use straight axis-aligned bounding box logic
                // Otherwise, return rotation angle for rotated axis-aligned (non-straight) clusters
                double thresholdDegrees = 2.0; // 2 degree tolerance
                
                // Helper function to check if an angle is straight axis-aligned to WCS
                bool IsStraightAxisAligned(double angleRad)
                {
                    double angleDeg = angleRad * 180 / Math.PI;
                    while (angleDeg < 0) angleDeg += 360;
                    while (angleDeg >= 360) angleDeg -= 360;
                    
                    double distTo0 = Math.Min(angleDeg, 360 - angleDeg);
                    double distTo90 = Math.Abs(angleDeg - 90);
                    double distTo180 = Math.Abs(angleDeg - 180);
                    double distTo270 = Math.Abs(angleDeg - 270);
                    
                    return distTo0 < thresholdDegrees || distTo90 < thresholdDegrees || 
                           distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
                }
                
                // Check if ALL individual angles are straight axis-aligned to WCS
                bool allStraightAxisAligned = rotationAngles.All(IsStraightAxisAligned);
                
                if (allStraightAxisAligned)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        string angleList = string.Join(", ", rotationAngles.Select(a => $"{a * 180 / Math.PI:F1}°"));
                        DebugLogger.Info($"[CLUSTER-ANGLE] All {rotationAngles.Count} angles are straight axis-aligned to WCS: [{angleList}], using straight axis-aligned bounding box");
                    }
                    return 0.0;
                }
                
                // Check if average angle is straight axis-aligned to WCS
                if (IsStraightAxisAligned(averageAngle))
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        double angleDegrees = averageAngle * 180 / Math.PI;
                        string angleList = string.Join(", ", rotationAngles.Select(a => $"{a * 180 / Math.PI:F1}°"));
                        DebugLogger.Info($"[CLUSTER-ANGLE] Average angle {angleDegrees:F1}° is straight axis-aligned to WCS (angles: [{angleList}]), using straight axis-aligned bounding box");
                    }
                    return 0.0;
                }

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CLUSTER-ANGLE] Determined dominant rotation angle: {averageAngle * 180 / Math.PI:F1}° from {rotationAngles.Count} sleeves (rotated axis-aligned/non-straight)");
                }

                return averageAngle;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[ClusterRotationService] Error determining dominant rotation angle: {ex.Message}");
                return 0.0;
            }
        }

        /// <summary>
        /// Calculate cluster bounding box using rotated coordinates from ClashZone data
        /// This is a simplified version - full implementation should be migrated from UniversalClusterService
        /// </summary>
        public (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)
        CalculateRotatedBoundingBox(List<dynamic> cluster, List<FamilyInstance>? actualSleeves, double rotationAngle, string? xmlFilePath = null)
        {
            // ✅ PERFORMANCE: Track calculation time (initialize at start for all code paths)
            var calcStopwatch = System.Diagnostics.Stopwatch.StartNew();
            
            // ✅ PERFORMANCE: Check cache first (cache key = sorted sleeve IDs + rotation angle)
            if (cluster != null && cluster.Count > 0)
            {
                try
                {
                    // Create cache key from sorted sleeve IDs + rotation angle
                    var sleeveIds = cluster.Select(s => s?.SleeveInstanceId ?? 0).Where(id => id > 0).OrderBy(id => id).ToList();
                    if (sleeveIds.Count == cluster.Count) // Only cache if all sleeves have valid IDs
                    {
                        string cacheKey = $"RBB_{string.Join("_", sleeveIds)}_{rotationAngle:F6}";
                        
                        if (_rotatedBboxCache.Value.TryGetValue(cacheKey, out var cachedResult))
                        {
                            calcStopwatch.Stop();
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ CACHE HIT: Rotated bounding box for {cluster.Count} sleeves, rotation={rotationAngle * 180 / Math.PI:F1}° (saved {calcStopwatch.ElapsedMilliseconds}ms), cacheKey={cacheKey.Substring(0, Math.Min(50, cacheKey.Length))}\n");
                                DebugLogger.Info($"[BBOX-CACHE] ✅ CACHE HIT: Rotated bounding box for {cluster.Count} sleeves, rotation={rotationAngle * 180 / Math.PI:F1}°");
                            }
                            return cachedResult;
                        }
                        else
                        {
                            // ✅ DIAGNOSTIC: Log cache miss for debugging
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 💾 CACHE MISS: Rotated bounding box for {cluster.Count} sleeves, rotation={rotationAngle * 180 / Math.PI:F1}°, cacheKey={cacheKey.Substring(0, Math.Min(50, cacheKey.Length))}, cacheSize={_rotatedBboxCache.Value.Count}\n");
                                DebugLogger.Info($"[BBOX-CACHE] 💾 CACHE MISS: Rotated bounding box for {cluster.Count} sleeves, rotation={rotationAngle * 180 / Math.PI:F1}°, cacheSize={_rotatedBboxCache.Value.Count}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // If cache key generation fails, continue with calculation
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cache key generation failed: {ex.Message}\n");
                    }
                }
            }
            
            // Simplified rotated bounding box calculation (union + optional corner refinement) without external ambiguous loggers.
            // ✅ OPTIMIZATION: actualSleeves is now optional - database data is primary source
            if (cluster == null || cluster.Count == 0)
                return (0,0,0,XYZ.Zero,null,null,null,null,null,null);

            // ✅ CRITICAL FIX: Calculate placement point from intersection points (centroid), not bounding box midpoint
            // Individual sleeves are placed at intersection points, so cluster should be at average of intersection points
            XYZ placementPoint = CalculatePlacementPointFromIntersections(cluster, xmlFilePath);
            
            if (placementPoint.IsZeroLength() || double.IsNaN(placementPoint.X))
            {
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to calculate placement point from intersections, falling back to first sleeve intersection point\n");
                // Fallback: Use first sleeve's intersection point
                var firstCz = GetCachedClashZone(cluster[0].SleeveInstanceId, xmlFilePath);
                if (firstCz != null)
                {
                    placementPoint = new XYZ(firstCz.IntersectionPointX, firstCz.IntersectionPointY, firstCz.IntersectionPointZ);
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}]   ✅ FALLBACK: Using first sleeve intersection point: ({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
                    }
                }
                else
                {
                    placementPoint = XYZ.Zero;
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}]   ❌ CRITICAL: First sleeve ClashZone is NULL, placementPoint set to XYZ.Zero!\n");
                    }
                }
            }
            else
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}]   ✅ Initial placement point from intersections: ({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
                }
            }

            // ✅ ROUND PIPE/DUCT FIX: Detect if this is a round pipe or round duct cluster and use bounding box extents instead of corners
            // For circular sleeves (round pipes/ducts), corners don't represent the actual extent properly
            // Instead, we need to use the bounding box min/max to get the true cluster width
            bool isRoundPipeCluster = false;
            foreach (var sleeveData in cluster)
            {
                if (sleeveData == null) continue;
                try
                {
                    int sleeveId = sleeveData.SleeveInstanceId;
                    var cz = GetCachedClashZone(sleeveId, xmlFilePath);
                    if (cz != null)
                    {
                        // Check for pipes (all pipes are circular)
                        bool isPipe = string.Equals(cz.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
                        
                        // Check for round ducts
                        bool isRoundDuct = false;
                        if (string.Equals(cz.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(cz.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
                        {
                            // Check if duct is round/circular
                            isRoundDuct = string.Equals(cz.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                         string.Equals(cz.DuctShape, "Circular", StringComparison.OrdinalIgnoreCase);
                        }
                        
                        if (isPipe || isRoundDuct)
                        {
                            isRoundPipeCluster = true;
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔵 CIRCULAR ELEMENT DETECTED: Sleeve {sleeveId}, Category={cz.MepElementCategory}, DuctShape={cz.DuctShape ?? "N/A"}, IsPipe={isPipe}, IsRoundDuct={isRoundDuct}\n");
                            break;
                        }
                    }
                }
                catch { continue; }
            }

            if (isRoundPipeCluster)
            {
                // ✅ ROUND PIPE/DUCT CLUSTER: Use bounding box extents instead of corners
                // For round pipes/ducts, the cluster width should be calculated from the min/max extents of all sleeve bounding boxes
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] 🔵 CIRCULAR CLUSTER DETECTED: Using bounding box extents instead of corners (applies to round pipes and round ducts)\n");

                double circularMinX = double.MaxValue, circularMaxX = double.MinValue;
                double circularMinY = double.MaxValue, circularMaxY = double.MinValue;
                double circularMinZ = double.MaxValue, circularMaxZ = double.MinValue;

                foreach (var sleeveData in cluster)
                {
                    if (sleeveData == null) continue;
                    try
                    {
                        int sleeveId = sleeveData.SleeveInstanceId;
                        var cz = GetCachedClashZone(sleeveId, xmlFilePath);
                        if (cz != null)
                        {
                            // Use sleeve bounding box (which includes clearance)
                            circularMinX = Math.Min(circularMinX, cz.SleeveBoundingBoxMinX);
                            circularMaxX = Math.Max(circularMaxX, cz.SleeveBoundingBoxMaxX);
                            circularMinY = Math.Min(circularMinY, cz.SleeveBoundingBoxMinY);
                            circularMaxY = Math.Max(circularMaxY, cz.SleeveBoundingBoxMaxY);
                            circularMinZ = Math.Min(circularMinZ, cz.SleeveBoundingBoxMinZ);
                            circularMaxZ = Math.Max(circularMaxZ, cz.SleeveBoundingBoxMaxZ);
                            
                            // ✅ CRITICAL FIX: Also track corner X/Y/Z values for tracking (horizontal AND vertical)
                            // SleeveBoundingBox* only gives individual sleeve size
                            // SleeveCorner* gives actual position in document coordinates
                            // This captures tracking where pipes are at different X/Y/Z levels
                            if (cz.SleeveCorner1X.HasValue && cz.SleeveCorner1Y.HasValue)
                            {
                                circularMinX = Math.Min(circularMinX, cz.SleeveCorner1X.Value);
                                circularMaxX = Math.Max(circularMaxX, cz.SleeveCorner1X.Value);
                                circularMinY = Math.Min(circularMinY, cz.SleeveCorner1Y.Value);
                                circularMaxY = Math.Max(circularMaxY, cz.SleeveCorner1Y.Value);
                            }
                            if (cz.SleeveCorner2X.HasValue && cz.SleeveCorner2Y.HasValue)
                            {
                                circularMinX = Math.Min(circularMinX, cz.SleeveCorner2X.Value);
                                circularMaxX = Math.Max(circularMaxX, cz.SleeveCorner2X.Value);
                                circularMinY = Math.Min(circularMinY, cz.SleeveCorner2Y.Value);
                                circularMaxY = Math.Max(circularMaxY, cz.SleeveCorner2Y.Value);
                            }
                            if (cz.SleeveCorner3X.HasValue && cz.SleeveCorner3Y.HasValue)
                            {
                                circularMinX = Math.Min(circularMinX, cz.SleeveCorner3X.Value);
                                circularMaxX = Math.Max(circularMaxX, cz.SleeveCorner3X.Value);
                                circularMinY = Math.Min(circularMinY, cz.SleeveCorner3Y.Value);
                                circularMaxY = Math.Max(circularMaxY, cz.SleeveCorner3Y.Value);
                            }
                            if (cz.SleeveCorner4X.HasValue && cz.SleeveCorner4Y.HasValue)
                            {
                                circularMinX = Math.Min(circularMinX, cz.SleeveCorner4X.Value);
                                circularMaxX = Math.Max(circularMaxX, cz.SleeveCorner4X.Value);
                                circularMinY = Math.Min(circularMinY, cz.SleeveCorner4Y.Value);
                                circularMaxY = Math.Max(circularMaxY, cz.SleeveCorner4Y.Value);
                            }
                            // Z tracking for vertical pipes
                            if (cz.SleeveCorner1Z.HasValue)
                            {
                                circularMinZ = Math.Min(circularMinZ, cz.SleeveCorner1Z.Value);
                                circularMaxZ = Math.Max(circularMaxZ, cz.SleeveCorner1Z.Value);
                            }
                            if (cz.SleeveCorner2Z.HasValue)
                            {
                                circularMinZ = Math.Min(circularMinZ, cz.SleeveCorner2Z.Value);
                                circularMaxZ = Math.Max(circularMaxZ, cz.SleeveCorner2Z.Value);
                            }
                            if (cz.SleeveCorner3Z.HasValue)
                            {
                                circularMinZ = Math.Min(circularMinZ, cz.SleeveCorner3Z.Value);
                                circularMaxZ = Math.Max(circularMaxZ, cz.SleeveCorner3Z.Value);
                            }
                            if (cz.SleeveCorner4Z.HasValue)
                            {
                                circularMinZ = Math.Min(circularMinZ, cz.SleeveCorner4Z.Value);
                                circularMaxZ = Math.Max(circularMaxZ, cz.SleeveCorner4Z.Value);
                            }

                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Sleeve {sleeveId}: BBox Min=({cz.SleeveBoundingBoxMinX:F6},{cz.SleeveBoundingBoxMinY:F6},{cz.SleeveBoundingBoxMinZ:F6}), Max=({cz.SleeveBoundingBoxMaxX:F6},{cz.SleeveBoundingBoxMaxY:F6},{cz.SleeveBoundingBoxMaxZ:F6}), CornerZ=({cz.SleeveCorner1Z ?? 0:F3},{cz.SleeveCorner3Z ?? 0:F3})\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}]   ⚠️ Error processing sleeve: {ex.Message}\n");
                        continue;
                    }
                }

                double circularWidth = circularMaxX - circularMinX;
                double circularHeight = circularMaxY - circularMinY;
                double circularDepth = circularMaxZ - circularMinZ;

                // ✅ CRITICAL FIX: For pipes, use actual sleeve diameter for Height, not bounding box Z-range
                // Get the maximum sleeve diameter from all sleeves in the cluster
                double maxSleeveDiameter = 0.0;
                double maxStructuralThickness = 0.0;
                foreach (var sleeveData in cluster)
                {
                    if (sleeveData == null) continue;
                    try
                    {
                        int sleeveId = sleeveData.SleeveInstanceId;
                        var cz = GetCachedClashZone(sleeveId, xmlFilePath);
                        if (cz != null)
                        {
                            // Use SleeveDiameter if available (this is the actual calculated diameter)
                            if (cz.SleeveDiameter > maxSleeveDiameter)
                                maxSleeveDiameter = cz.SleeveDiameter;
                            
                            // Get structural thickness for depth
                            // ✅ CRITICAL FIX: For Floors, ALWAYS use StructuralElementThickness (actual slab thickness). 
                            // WallThickness often contains default garbage (0.1) for floors.
                            // For Walls, prefer WallThickness if present.
                            bool isFloorElement = false;
                            if (!string.IsNullOrEmpty(cz.StructuralElementType))
                            {
                                isFloorElement = cz.StructuralElementType.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0;
                            }

                            double currentThickness;
                            if (isFloorElement)
                            {
                                currentThickness = cz.StructuralElementThickness;
                            }
                            else
                            {
                                currentThickness = cz.WallThickness > 0 ? cz.WallThickness : cz.StructuralElementThickness;
                            }

                            if (currentThickness > maxStructuralThickness)
                                maxStructuralThickness = currentThickness;
                            
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}]   Sleeve {sleeveId}: W={cz.SleeveWidth * 304.8:F1}mm, H={cz.SleeveHeight * 304.8:F1}mm, Depth={cz.StructuralElementThickness * 304.8:F1}mm, Diameter={cz.SleeveDiameter * 304.8:F1}mm\n");
                        }
                    }
                    catch { continue; }
                }

                // ✅ ROBUST SIZING: Determine wall direction explicitly from database properties
                // Do NOT guess based on aspect ratio (Math.Abs(dY) > Math.Abs(dX)) because for square grids or combined clusters it fails.
                bool isYWall = false;
                bool isXWall = false;
                bool isFloor = false; // ✅ DETECT FLOORS
                
                // Inspect first few elements to find authoritative wall direction / host type
                foreach (var sleeveData in cluster)
                {
                    if (sleeveData == null) continue;
                    try
                    {
                        int sleeveId = sleeveData.SleeveInstanceId;
                        var cz = GetCachedClashZone(sleeveId, xmlFilePath);
                        if (cz != null)
                        {
                            string wDir = cz.WallDirectionType ?? "";
                            string hOri = cz.HostOrientation ?? "";
                            string structType = cz.StructuralElementType ?? "";

                            // Check for Floor
                            if (structType.Equals("Floor", StringComparison.OrdinalIgnoreCase) || 
                                structType.Equals("Floors", StringComparison.OrdinalIgnoreCase))
                            {
                                isFloor = true;
                                break;
                            }
                            
                            // Check explicit flags first
                            isXWall = wDir.Contains("X-WALL", StringComparison.OrdinalIgnoreCase) || hOri.Equals("X", StringComparison.OrdinalIgnoreCase);
                            isYWall = wDir.Contains("Y-WALL", StringComparison.OrdinalIgnoreCase) || hOri.Equals("Y", StringComparison.OrdinalIgnoreCase);
                            
                            if (isXWall || isYWall) break; // Found authoritative direction
                        }
                    }
                    catch { }
                }

                // Fallback heuristic ONLY if explicit flags missing
                if (!isXWall && !isYWall)
                {
                    // If height is dominant (vertical/stacked), aspect ratio of X/Y might be misleading about wall direction
                    // But we have no choice but to guess if DB lacks info.
                    isYWall = (circularMaxY - circularMinY) > (circularMaxX - circularMinX); 
                    isXWall = !isYWall;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ Wall direction unknown, guessing: IsYWall={isYWall} (based on aspect ratio)\n");
                }
                
                // Calculate raw bounding box dimensions
                double rawWidthX = circularMaxX - circularMinX;
                double rawWidthY = circularMaxY - circularMinY;
                double rawHeightZ = circularMaxZ - circularMinZ;
                
                // ✅ Z-HEIGHT LOGIC (COMMON):
                // If Z-range is larger than max diameter (vertical stack), use Z-range.
                // Otherwise use MaxDiameter (for precision on horizontal runs).
                // Factor 1.05 provides a 5% tolerance.
                // ✅ CRITICAL FIX: For Floors, Height is Plan Y (or X), not Z. Only use this Z-logic for WALLS.
                if (!isFloor)
                {
                    circularHeight = rawHeightZ > (maxSleeveDiameter * 1.05) ? rawHeightZ : maxSleeveDiameter;
                    // Safety clamp
                    if (circularHeight < maxSleeveDiameter) circularHeight = maxSleeveDiameter;
                }

                if (isXWall)
                {
                    // X-WALL (wall runs ALONG X-axis, normal is along Y):
                    // Width is the span ALONG the wall = X-axis = rawWidthX
                    // Depth is THROUGH the wall = Y-axis = rawWidthY (wall thickness)
                    circularWidth = rawWidthX;
                    
                    // Depth logic: Use structural thickness if valid, else fallback to BBox Y-depth
                    double bboxDepth = rawWidthY;
                    circularDepth = maxStructuralThickness > 0 ? maxStructuralThickness : bboxDepth;
                }

                else if (isFloor)
                {
                    // ✅ FLOOR LOGIC:
                    // Width = Plan X (rawWidthX)
                    // Height = Plan Y (rawWidthY)
                    // Depth = Thickness (maxStructuralThickness or rawHeightZ)
                    
                    circularWidth = rawWidthX;
                    circularHeight = rawWidthY; // Plan Y
                    
                    // Depth logic (Thickness)
                    // Use maxStructuralThickness if valid (preferred), else Z-bbox (likely pipe stick-out, be careful)
                    // For floors, Z-bbox might be the pipe length. We strictly want slab thickness.
                    circularDepth = maxStructuralThickness > 0 ? maxStructuralThickness : rawHeightZ;

                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ FLOOR HOST: Mapped Width=X ({circularWidth*304.8:F1}), Height=Y ({circularHeight*304.8:F1}, Depth=Z/Thick ({circularDepth*304.8:F1})\n");
                }
                else
                {
                    // Y-WALL (wall runs ALONG Y-axis, normal is along X):
                    // Width is the span ALONG the wall = Y-axis = rawWidthY
                    // Depth is THROUGH the wall = X-axis = rawWidthX (wall thickness)
                    // ✅ FIX: Was backwards - had Width=rawWidthX, Depth=rawWidthY (wrong!)
                    circularWidth = rawWidthY;

                    // Depth logic: Use structural thickness if valid, else fallback to BBox X-depth
                    double bboxDepth = rawWidthX;
                    circularDepth = maxStructuralThickness > 0 ? maxStructuralThickness : bboxDepth;
                }

                // ✅ FAILSAFE: Swap Width and Depth if Depth is unrealistically large compared to Width
                // This catches cases where wall direction is misidentified (e.g. Y-Wall identified as X-Wall)
                // A sleeve is typically Wider (along wall) than it is Deep (wall thickness).
                // If Depth > 1.5 * Width AND Depth > 300mm (1ft), it's likely the dimensions are swapped.
                // Exception: If structural thickness is explicitly large (e.g. > 500mm), don't swap.
                bool explicitLargeThickness = maxStructuralThickness > 0.5; // > 500mm
                if (!explicitLargeThickness && circularDepth > (circularWidth * 1.5) && circularDepth > 0.3)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                         SafeFileLogger.SafeAppendText("cluster_sizing.log", $"[{DateTime.Now:HH:mm:ss}] ⚠️ SUSPICIOUS DIMENSIONS: Depth ({circularDepth*304.8:F0}mm) > Width ({circularWidth*304.8:F0}mm). Swapping Width/Depth (assuming misidentified wall direction).\n");
                    
                    double temp = circularWidth;
                    circularWidth = circularDepth;
                    // For the new depth, use the smaller dimension (was width) OR rely on maxStructuralThickness if valid
                    circularDepth = maxStructuralThickness > 0 ? maxStructuralThickness : temp;
                }

                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] 🔵 CIRCULAR CLUSTER RESULT (Pipes/Round Ducts): Width={circularWidth * 304.8:F1}mm, Height={circularHeight * 304.8:F1}mm, Depth={circularDepth * 304.8:F1}mm\n");
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}]   IsXWall={isXWall}, IsYWall={isYWall}, RawBBox=({rawWidthX*304.8:F0}x{rawWidthY*304.8:F0}x{rawHeightZ*304.8:F0})\n");
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}]   MaxSleeveDiameter={maxSleeveDiameter * 304.8:F1}mm, MaxStructuralThickness={maxStructuralThickness * 304.8:F1}mm\n");

                // Cache the result
                try
                {
                    var sleeveIds = cluster.Select(s => s?.SleeveInstanceId ?? 0).Where(id => id > 0).OrderBy(id => id).ToList();
                    if (sleeveIds.Count == cluster.Count)
                    {
                        string cacheKey = $"RBB_{string.Join("_", sleeveIds)}_{rotationAngle:F6}";
                        var circularResult = (circularWidth, circularHeight, circularDepth, placementPoint, (double?)circularMinX, (double?)circularMinY, (double?)circularMinZ, (double?)circularMaxX, (double?)circularMaxY, (double?)circularMaxZ);
                        
                        if (_rotatedBboxCache.Value.Count < MAX_ROTATED_BBOX_CACHE_SIZE)
                        {
                            _rotatedBboxCache.Value[cacheKey] = circularResult;
                        }
                        
                        calcStopwatch.Stop();
                        return circularResult;
                    }
                }
                catch { }

                calcStopwatch.Stop();
                return (circularWidth, circularHeight, circularDepth, placementPoint, circularMinX, circularMinY, circularMinZ, circularMaxX, circularMaxY, circularMaxZ);
            }



            // ✅ CRITICAL FIX: Use corner-based calculation for ALL clusters (not just rotated ones)
            // Corners are always saved and are the authoritative source for accurate sizing
            // For walls/framing, we'll transform corners to RCS before calculating dimensions
            // This ensures consistent sizing regardless of rotation angle (straight or rotated axis)
            var firstSleeveData = cluster[0];
            bool isWallOrFraming = false;
            XYZ? wallDirection = null;
            XYZ? wallOrigin = null;
            
            if (firstSleeveData != null)
            {
                try
                {
                    int firstSleeveInstanceId = firstSleeveData.SleeveInstanceId;
                    if (firstSleeveInstanceId > 0)
                    {
                        var firstCz = GetCachedClashZone(firstSleeveInstanceId, xmlFilePath);
                        if (firstCz != null)
                        {
                            bool isWallHost = string.Equals(firstCz.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                              string.Equals(firstCz.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                            bool isFramingHost = string.Equals(firstCz.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
                            
                            if ((isWallHost || isFramingHost) && firstCz.WallDirection != null && !firstCz.WallDirection.IsZeroLength())
                            {
                                isWallOrFraming = true;
                                wallDirection = firstCz.WallDirection;
                                wallOrigin = new XYZ(firstCz.SleevePlacementPointActiveDocumentX, 
                                                    firstCz.SleevePlacementPointActiveDocumentY, 
                                                    firstCz.SleevePlacementPointActiveDocumentZ);
                                
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🔍 WALL/FRAMING DETECTED: Will use corner-based calculation with RCS transformation (straight or rotated axis)\n");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error checking host type: {ex.Message}\n");
                }
            }

            var rotatedBboxes = new List<(XYZ min, XYZ max)>();
            foreach (var sleeveData in cluster)
            {
                // ✅ CRASH-SAFE: Validate sleeveData and SleeveInstanceId before calling function
                if (sleeveData == null)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Sleeve data is NULL in cluster\n");
                    continue;
                }
                
                // ✅ CRASH-SAFE: Get SleeveInstanceId safely (handle dynamic binding failure)
                int sleeveInstanceId = 0;
                try
                {
                    sleeveInstanceId = sleeveData.SleeveInstanceId;
                    if (sleeveInstanceId <= 0)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Invalid SleeveInstanceId: {sleeveInstanceId}\n");
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ❌ Error getting SleeveInstanceId: {ex.Message}\n");
                    continue;
                }
                
                // ✅ PERFORMANCE: Use cached ClashZone lookup
                var clashZone = GetCachedClashZone(sleeveInstanceId, xmlFilePath);
                if (clashZone == null)
                if (clashZone == null)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Sleeve {sleeveInstanceId}: ClashZone cast failed\n");
                    continue;
                }
                
                // ✅ DIAGNOSTIC: Log all available data from database
                bool hasRot = clashZone.RotatedBoundingBoxMinX.HasValue && clashZone.RotatedBoundingBoxMaxX.HasValue && 
                              clashZone.RotatedBoundingBoxMinY.HasValue && clashZone.RotatedBoundingBoxMaxY.HasValue;
                bool hasCorners = clashZone.SleeveCorner1X.HasValue && clashZone.SleeveCorner1Y.HasValue &&
                                  clashZone.SleeveCorner2X.HasValue && clashZone.SleeveCorner2Y.HasValue &&
                                  clashZone.SleeveCorner3X.HasValue && clashZone.SleeveCorner3Y.HasValue &&
                                  clashZone.SleeveCorner4X.HasValue && clashZone.SleeveCorner4Y.HasValue;
                bool hasCosSin = clashZone.MepRotationCos.HasValue && clashZone.MepRotationSin.HasValue;
                
                bool hasAxisAligned = clashZone.SleeveBoundingBoxMinX != 0.0 || clashZone.SleeveBoundingBoxMinY != 0.0 || clashZone.SleeveBoundingBoxMaxX != 0.0 || clashZone.SleeveBoundingBoxMaxY != 0.0;
                
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] 📊 Sleeve {sleeveInstanceId} DATA CHECK: HasRotatedBbox={hasRot}, HasCorners={hasCorners}, HasCosSin={hasCosSin}, HasAxisAligned={hasAxisAligned}, Rotation={clashZone.MepElementRotationAngle * 180 / Math.PI:F1}°\n");
                
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}]   Sleeve {sleeveInstanceId}: HasRotatedBbox={hasRot}, HasAxisAlignedBbox={hasAxisAligned}\n");
                if (hasRot)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}]     RotatedBbox: MinX={clashZone.RotatedBoundingBoxMinX.Value:F6}, MinY={clashZone.RotatedBoundingBoxMinY.Value:F6}, MaxX={clashZone.RotatedBoundingBoxMaxX.Value:F6}, MaxY={clashZone.RotatedBoundingBoxMaxY.Value:F6}\n");
                    rotatedBboxes.Add((
                        new XYZ(clashZone.RotatedBoundingBoxMinX!.Value, clashZone.RotatedBoundingBoxMinY!.Value, clashZone.RotatedBoundingBoxMinZ ?? clashZone.SleeveBoundingBoxMinZ),
                        new XYZ(clashZone.RotatedBoundingBoxMaxX!.Value, clashZone.RotatedBoundingBoxMaxY!.Value, clashZone.RotatedBoundingBoxMaxZ ?? clashZone.SleeveBoundingBoxMaxZ)
                    ));
                }
                else
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}]     ⚠️ RotatedBbox values: MinX={clashZone.RotatedBoundingBoxMinX.HasValue}, MinY={clashZone.RotatedBoundingBoxMinY.HasValue}, MaxX={clashZone.RotatedBoundingBoxMaxX.HasValue}, MaxY={clashZone.RotatedBoundingBoxMaxY.HasValue}\n");
                }
            }
            
            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                $"[{DateTime.Now:HH:mm:ss}] 📊 Collected rotatedBboxes: {rotatedBboxes.Count} out of {cluster.Count} sleeves\n");

            // ✅ CRITICAL FIX: Use corner-based calculation for ALL clusters (not just rotated ones)
            // Corners are always saved and are the authoritative source for accurate sizing
            // This ensures consistent sizing regardless of rotation angle or host type
            // ✅ APPLIES TO ALL CATEGORIES: Cable Trays, Ducts, Pipes, Walls, Framing (all rectangular elements)
            // This was originally debugged for cable trays but works universally for all categories
            // For walls/framing, we transform corners to RCS before calculating dimensions
            bool useCornerBasedCalculation = true; // Always use corner-based for accurate sizing
            if (useCornerBasedCalculation)
            {
                // ✅ WATERTIGHT ALGORITHM: Use corner-based calculation for accurate sizing
                // This works even if rotated bounding boxes aren't in database, as long as corners are available
                // ✅ UNIVERSAL: Applies to ALL categories - Cable Trays, Ducts, Pipes (rectangular elements with rotated axis)
                // Corners are saved during individual sleeve placement for all categories, so this works universally
                
                // ✅ CRITICAL FIX: Extract SleeveInstanceId BEFORE try block so catch block can access it
                var sleeveIdsInCluster = new List<int>();
                foreach (var sleeveData in cluster)
                {
                    try
                    {
                        int sleeveId;
                        dynamic dynSleeve = sleeveData;
                        object sleeveIdObj = dynSleeve.SleeveInstanceId;
                        if (sleeveIdObj == null) continue;
                        
                        if (sleeveIdObj is int id)
                            sleeveId = id;
                        else if (sleeveIdObj is long longId)
                            sleeveId = (int)longId;
                        else
                            sleeveId = Convert.ToInt32(sleeveIdObj);
                        
                        sleeveIdsInCluster.Add(sleeveId);
                    }
                    catch
                    {
                        continue;
                    }
                }
                
                try
                {
                    // ✅ DIAGNOSTIC: Log before attempting corner-based calculation
                    double rotationDeg = rotationAngle * 180.0 / Math.PI;
                    string hostType = isWallOrFraming ? "Wall/Framing" : "Floor/Other";
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 ATTEMPTING corner-based calculation: rotation={rotationDeg:F1}°, hostType={hostType}, clusterSize={cluster.Count}, rotatedBboxesCount={rotatedBboxes.Count}\n");
                    
                    // ✅ PRE-CHECK: Verify at least one sleeve has corners before attempting calculation
                    // Corners are ALWAYS saved for all sleeves during individual placement
                    // Note: sleeveIdsInCluster was already extracted before the try block
                    bool hasAnyCorners = false;
                    int sleevesWithCorners = 0;
                    
                    // Now check corners using the extracted IDs (already extracted before try block)
                    foreach (int sleeveId in sleeveIdsInCluster)
                    {
                        var clashZone = GetCachedClashZone(sleeveId, xmlFilePath);
                        if (clashZone != null)
                        {
                            double sleeveRotationDeg = clashZone.MepElementRotationAngle * 180.0 / Math.PI;
                            
                            // ✅ CRITICAL FIX: Explicitly type nullable properties to avoid dynamic dispatch errors
                            double? corner1X = clashZone.SleeveCorner1X;
                            double? corner1Y = clashZone.SleeveCorner1Y;
                            double? corner1Z = clashZone.SleeveCorner1Z;
                            double? corner2X = clashZone.SleeveCorner2X;
                            double? corner2Y = clashZone.SleeveCorner2Y;
                            double? corner2Z = clashZone.SleeveCorner2Z;
                            double? corner3X = clashZone.SleeveCorner3X;
                            double? corner3Y = clashZone.SleeveCorner3Y;
                            double? corner3Z = clashZone.SleeveCorner3Z;
                            double? corner4X = clashZone.SleeveCorner4X;
                            double? corner4Y = clashZone.SleeveCorner4Y;
                            double? corner4Z = clashZone.SleeveCorner4Z;
                            
                            bool hasCorners = corner1X.HasValue && corner1Y.HasValue &&
                                              corner2X.HasValue && corner2Y.HasValue &&
                                              corner3X.HasValue && corner3Y.HasValue &&
                                              corner4X.HasValue && corner4Y.HasValue;
                            
                            bool hasCornerZ = corner1Z.HasValue && corner2Z.HasValue && 
                                              corner3Z.HasValue && corner4Z.HasValue;
                            
                            // ✅ DIAGNOSTIC: Log corner Z availability
                            if (!hasCornerZ)
                            {
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}]   ⚠️ Sleeve {sleeveId}: Corner Z coordinates NOT saved in database! " +
                                    $"Corner1Z={corner1Z.HasValue}, Corner2Z={corner2Z.HasValue}, Corner3Z={corner3Z.HasValue}, Corner4Z={corner4Z.HasValue}, " +
                                    $"Will use SleeveBoundingBoxMinZ={clashZone.SleeveBoundingBoxMinZ:F6} as fallback\n");
                            }
                            else
                            {
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ Sleeve {sleeveId}: Corner Z coordinates found in database! " +
                                        $"Corner1Z={corner1Z.GetValueOrDefault():F6}, Corner2Z={corner2Z.GetValueOrDefault():F6}, Corner3Z={corner3Z.GetValueOrDefault():F6}, Corner4Z={corner4Z.GetValueOrDefault():F6}, " +
                                        $"PlacementPointZ={clashZone.SleevePlacementPointActiveDocumentZ:F6}, BBoxMinZ={clashZone.SleeveBoundingBoxMinZ:F6}\n");

                            }
                            
                            if (hasCorners)
                            {
                                sleevesWithCorners++;
                                hasAnyCorners = true;
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}]   ✅ Sleeve {sleeveId}: Has corners, Rotation={sleeveRotationDeg:F1}°\n");
                            }
                            else
                            {
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}]   ⚠️ Sleeve {sleeveId}: Missing corners, Rotation={sleeveRotationDeg:F1}°\n");
                            }
                        }
                    }
                    
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] 📊 Corner check: {sleevesWithCorners}/{cluster.Count} sleeves have corners\n");
                    
                    if (hasAnyCorners)
                    {
                        // ✅ MANUAL CORNER-BASED CALCULATION: Calculate directly here to avoid dynamic type issues
                        // ✅ CRITICAL: Uses pre-saved corners from database (NO recalculation needed)
                        // - Corners were batch calculated and saved to database during individual sleeve placement
                        // - After regeneration, corners were batch saved via SleevePersistenceService.PersistSleeveData
                        // - Cluster calculation reads corners directly from database (SleeveCorner1X/Y/Z through Corner4X/Y/Z)
                        // - This ensures accurate cluster sizing using pre-calculated corner coordinates
                        var allCorners = new List<XYZ>();
                        
                        foreach (int sleeveId in sleeveIdsInCluster)
                        {
                            var cz = GetCachedClashZone(sleeveId, xmlFilePath);
                            if (cz != null)
                            {
                                // ✅ CRITICAL FIX: Explicitly type nullable properties to avoid dynamic dispatch errors
                                // ✅ READ FROM DATABASE: These corners were batch saved after regeneration
                                // - No recalculation needed - corners are already in database
                                double? corner1X = cz.SleeveCorner1X;
                                double? corner1Y = cz.SleeveCorner1Y;
                                double? corner1Z = cz.SleeveCorner1Z;
                                double? corner2X = cz.SleeveCorner2X;
                                double? corner2Y = cz.SleeveCorner2Y;
                                double? corner2Z = cz.SleeveCorner2Z;
                                double? corner3X = cz.SleeveCorner3X;
                                double? corner3Y = cz.SleeveCorner3Y;
                                double? corner3Z = cz.SleeveCorner3Z;
                                double? corner4X = cz.SleeveCorner4X;
                                double? corner4Y = cz.SleeveCorner4Y;
                                double? corner4Z = cz.SleeveCorner4Z;
                                
                                if (corner1X.HasValue && corner1Y.HasValue &&
                                    corner2X.HasValue && corner2Y.HasValue &&
                                    corner3X.HasValue && corner3Y.HasValue &&
                                    corner4X.HasValue && corner4Y.HasValue)
                                {
                                    // ✅ CRITICAL FIX: Use stored corner Z coordinates if available, otherwise fall back to bounding box MinZ
                                    // The stored corner Z coordinates are the authoritative source - they were calculated from the actual sleeve placement
                                    double corner1ZValue = corner1Z.HasValue ? corner1Z.Value : cz.SleeveBoundingBoxMinZ;
                                    double corner2ZValue = corner2Z.HasValue ? corner2Z.Value : cz.SleeveBoundingBoxMinZ;
                                    double corner3ZValue = corner3Z.HasValue ? corner3Z.Value : cz.SleeveBoundingBoxMinZ;
                                    double corner4ZValue = corner4Z.HasValue ? corner4Z.Value : cz.SleeveBoundingBoxMinZ;
                                    
                                    allCorners.Add(new XYZ(corner1X.Value, corner1Y.Value, corner1ZValue));
                                    allCorners.Add(new XYZ(corner2X.Value, corner2Y.Value, corner2ZValue));
                                    allCorners.Add(new XYZ(corner3X.Value, corner3Y.Value, corner3ZValue));
                                    allCorners.Add(new XYZ(corner4X.Value, corner4Y.Value, corner4ZValue));
                                    
                                    bool usingStoredZ = corner1Z.HasValue && corner2Z.HasValue && corner3Z.HasValue && corner4Z.HasValue;
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ Added 4 corners for sleeve {sleeveId}: " +
                                        $"C1=({corner1X.Value:F6},{corner1Y.Value:F6},{corner1ZValue:F6}), " +
                                        $"C2=({corner2X.Value:F6},{corner2Y.Value:F6},{corner2ZValue:F6}), " +
                                        $"C3=({corner3X.Value:F6},{corner3Y.Value:F6},{corner3ZValue:F6}), " +
                                        $"C4=({corner4X.Value:F6},{corner4Y.Value:F6},{corner4ZValue:F6}) " +
                                        $"[UsingStoredZ={usingStoredZ}]\n");
                                }
                            }
                        }
                        
                        if (allCorners.Count >= 4)
                        {
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] 🔧 Starting corner-based calculation with {allCorners.Count} corners, rotation={rotationAngle * 180 / Math.PI:F1}°, isWallOrFraming={isWallOrFraming}\n");
                            
                            double cornerWidth = 0, cornerHeight = 0, cornerDepth = 0; // Initialize to avoid unassigned variable error
                            double rotatedMinX = 0, rotatedMinY = 0, rotatedMaxX = 0, rotatedMaxY = 0;
                            double originX = 0, originY = 0;
                            
                            // ✅ METHODOLOGY COMPLIANCE: According to SLEEVE_PLACEMENT_METHODOLOGY.md (line 850+),
                            //    RCS transformation is NOT used for walls. Instead:
                            //    - For straight axis (rotationAngle = 0°): Use world-space corners directly
                            //    - For rotated axis: Transform corners to cluster's rotated coordinate system using rotation matrix
                            if (isWallOrFraming && wallDirection != null && wallOrigin != null)
                            {
                                // ✅ WALL/FRAMING: Use rotation matrix approach (NOT RCS) per methodology
                                // For straight axis (rotationAngle ≈ 0°): Use world-space corners directly
                                // For rotated axis: Transform corners to cluster's rotated coordinate system
                                
                                // Check if rotation angle is near zero (straight axis)
                                bool isStraightAxis = Math.Abs(rotationAngle) < 1e-6;
                                
                                if (isStraightAxis)
                                {
                                    // ✅ STRAIGHT AXIS: Use world-space corners directly (no transformation needed)
                                    // For Y-wall: Width = Y range, Height = Z range
                                    // For X-wall: Width = X range, Height = Z range
                                    double wcsMinX = allCorners.Min(c => c.X);
                                    double wcsMaxX = allCorners.Max(c => c.X);
                                    double wcsMinY = allCorners.Min(c => c.Y);
                                    double wcsMaxY = allCorners.Max(c => c.Y);
                                    
                                    // ✅ CRITICAL: For walls, width is ALWAYS along the wall direction
                                    // Y-wall: Wall runs along Y-axis → Width = Y range (along the wall)
                                    // X-wall: Wall runs along X-axis → Width = X range (along the wall)
                                    bool isYWall = Math.Abs(wallDirection.Y) > Math.Abs(wallDirection.X);
                                    
                                    if (isYWall)
                                    {
                                        // ✅ Y-wall: Wall runs along Y-axis, so width (along the wall) = Y range
                                        cornerWidth = wcsMaxY - wcsMinY;
                                        rotatedMinX = wcsMinX;
                                        rotatedMaxX = wcsMaxX;
                                        rotatedMinY = wcsMinY;
                                        rotatedMaxY = wcsMaxY;
                                    }
                                    else
                                    {
                                        // ✅ X-wall: Wall runs along X-axis, so width (along the wall) = X range
                                        cornerWidth = wcsMaxX - wcsMinX;
                                        rotatedMinX = wcsMinX;
                                        rotatedMaxX = wcsMaxX;
                                        rotatedMinY = wcsMinY;
                                        rotatedMaxY = wcsMaxY;
                                    }
                                    
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ STRAIGHT AXIS WALL: Using world-space corners directly (no RCS, no rotation). " +
                                        $"WallType={(isYWall ? "Y-wall" : "X-wall")}, " +
                                        $"minX={wcsMinX:F6}, maxX={wcsMaxX:F6}, minY={wcsMinY:F6}, maxY={wcsMaxY:F6}, " +
                                        $"Width={cornerWidth:F6}\n");
                                }
                                else
                                {
                                    // ✅ ROTATED AXIS: Transform corners to cluster's rotated coordinate system using rotation matrix
                                    // Per methodology (line 947-996): Use rotation matrix, NOT RCS
                                    double cosCluster = Math.Cos(rotationAngle);
                                    double sinCluster = Math.Sin(rotationAngle);
                                    
                                    // Use first sleeve center as origin (per methodology line 950)
                                    originX = wallOrigin.X;
                                    originY = wallOrigin.Y;
                                    
                                    // Transform all corners to cluster's rotated coordinate system
                                    var transformedCorners = new List<XYZ>();
                                    foreach (var corner in allCorners)
                                    {
                                        // Translate relative to origin
                                        double relX = corner.X - originX;
                                        double relY = corner.Y - originY;
                                        
                                        // Rotate to cluster's intended axis coordinate system (per methodology line 971-972)
                                        double clusterX = relX * cosCluster - relY * sinCluster;
                                        double clusterY = relX * sinCluster + relY * cosCluster;
                                        
                                        transformedCorners.Add(new XYZ(clusterX, clusterY, corner.Z));
                                    }
                                    
                                    // Find min/max extents in rotated coordinate system
                                    double minRotX = transformedCorners.Min(c => c.X);
                                    double maxRotX = transformedCorners.Max(c => c.X);
                                    double minRotY = transformedCorners.Min(c => c.Y);
                                    double maxRotY = transformedCorners.Max(c => c.Y);
                                    
                                    // ✅ CRITICAL FIX: For rotated axis walls, determine X-wall vs Y-wall to calculate width correctly
                                    // X-wall with 90° rotation: X-axis in world space → Y-axis in rotated space, so width = Y range
                                    // Y-wall with rotation (rare): Y-axis in world space → X-axis in rotated space, so width = X range
                                    bool isYWall = Math.Abs(wallDirection.Y) > Math.Abs(wallDirection.X);
                                    
                                    if (isYWall)
                                    {
                                        // ✅ Y-WALL WITH ROTATION (rare case): Y-axis in world space → X-axis in rotated space after rotation
                                        // Width (along wall) = X range in rotated space
                                        cornerWidth = maxRotX - minRotX;
                                    }
                                    else
                                    {
                                        // ✅ X-WALL WITH 90° ROTATION: X-axis in world space → Y-axis in rotated space after 90° rotation
                                        // Width (along wall) = Y range in rotated space, NOT X range
                                        cornerWidth = maxRotY - minRotY;
                                    }
                                    
                                    rotatedMinX = minRotX;
                                    rotatedMaxX = maxRotX;
                                    rotatedMinY = minRotY;
                                    rotatedMaxY = maxRotY;
                                    
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ ROTATED AXIS WALL: Transformed {transformedCorners.Count} corners to cluster's rotated coordinate system. " +
                                        $"Rotation={rotationAngle * 180 / Math.PI:F1}°, WallType={(isYWall ? "Y-wall" : "X-wall")}, " +
                                        $"minRotX={minRotX:F6}, maxRotX={maxRotX:F6}, minRotY={minRotY:F6}, maxRotY={maxRotY:F6}, " +
                                        $"Width={cornerWidth:F6} (from {(isYWall ? "X" : "Y")} range in rotated space)\n");
                                }
                            }
                            
                            if (!isWallOrFraming)
                            {
                                // ✅ FLOOR/OTHER: Calculate in WCS with rotation if needed
                                // Calculate cluster bounding box by rotating corners back to aligned axis
                                
                                // ✅ CRITICAL: Check for mixed X/Y orientations BEFORE rotation transformation
                                // When MEP elements have different orientations (X and Y), corner-based calculation may not work correctly
                                // because corners are in different coordinate systems. Use envelope method instead.
                                bool hasMixedOrientations = false;
                                var orientations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                
                                foreach (int sleeveId in sleeveIdsInCluster)
                                {
                                    var cz = GetCachedClashZone(sleeveId, xmlFilePath);
                                    if (cz != null)
                                    {
                                        string orientation = cz.MepElementOrientationDirection;
                                        
                                        // ✅ Fallback: Infer orientation from rotation angle if missing
                                        if (string.IsNullOrEmpty(orientation))
                                        {
                                            double angle = Math.Abs(cz.MepElementRotationAngle); // Radians
                                            // Normalize to 0-PI
                                            while (angle > Math.PI) angle -= Math.PI;
                                            
                                            if (angle > Math.PI / 4.0 && angle < 3.0 * Math.PI / 4.0)
                                            {
                                                orientation = "Y"; // ~90 degrees
                                            }
                                            else
                                            {
                                                orientation = "X"; // ~0 or ~180 degrees
                                            }
                                        }
                                        
                                        if (!string.IsNullOrEmpty(orientation))
                                        {
                                            orientations.Add(orientation);
                                        }
                                    }
                                }
                                
                                hasMixedOrientations = orientations.Contains("X") && orientations.Contains("Y");
                                
                                if (hasMixedOrientations)
                                {
                                    // ✅ MIXED X/Y ORIENTATIONS: Build envelope directly from stored corners
                                    // Corners are already collected in allCorners - use them directly!
                                    // This is the most accurate method since corners are the actual sleeve positions
                                    
                                    if (allCorners.Count >= 4)
                                    {
                                        // ✅ Use corners directly - they're already in world space and accurate
                                        double floorMinX = allCorners.Min(c => c.X);
                                        double floorMinY = allCorners.Min(c => c.Y);
                                        double floorMinZ = allCorners.Min(c => c.Z);
                                        double floorMaxX = allCorners.Max(c => c.X);
                                        double floorMaxY = allCorners.Max(c => c.Y);
                                        double floorMaxZ = allCorners.Max(c => c.Z);
                                        
                                        cornerWidth = floorMaxX - floorMinX;
                                        cornerHeight = floorMaxY - floorMinY;
                                        cornerDepth = floorMaxZ - floorMinZ;
                                        
                                        // ✅ Calculate new center point from corner extents
                                        // X and Y: Midpoint of corner extents
                                        // Z: Get from first sleeve (not from corner extents)
                                        double midX = (floorMinX + floorMaxX) / 2.0;
                                        double midY = (floorMinY + floorMaxY) / 2.0;
                                        double placementZ = 0.0;
                                        
                                        // Z from first sleeve in cluster
                                        if (sleeveIdsInCluster.Count > 0)
                                        {
                                            var firstCz = GetCachedClashZone(sleeveIdsInCluster[0], xmlFilePath);
                                            if (firstCz != null)
                                            {
                                                placementZ = firstCz.SleevePlacementPointActiveDocumentZ;
                                                if (placementZ == 0.0)
                                                    placementZ = firstCz.IntersectionPointZ;
                                                if (placementZ == 0.0 && firstCz.SleeveCorner1Z.HasValue)
                                                    placementZ = firstCz.SleeveCorner1Z.Value;
                                            }
                                        }
                                        
                                        placementPoint = new XYZ(midX, midY, placementZ);
                                        
                                        // Store bounds for extents
                                        rotatedMinX = floorMinX;
                                        rotatedMaxX = floorMaxX;
                                        rotatedMinY = floorMinY;
                                        rotatedMaxY = floorMaxY;
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            double wMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerWidth);
                                            double hMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerHeight);
                                            double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerDepth);
                                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                $"[{DateTime.Now:HH:mm:ss}] ✅ FLOOR MIXED X/Y ORIENTATIONS: Building envelope from STORED CORNERS: " +
                                                $"W={wMm:F1}mm, H={hMm:F1}mm, D={dMm:F1}mm, " +
                                                $"Orientations=[{string.Join(", ", orientations)}], Sleeves={sleeveIdsInCluster.Count}, " +
                                                $"Corners={allCorners.Count}, " +
                                                $"Min=({floorMinX:F6}, {floorMinY:F6}, {floorMinZ:F6}), " +
                                                $"Max=({floorMaxX:F6}, {floorMaxY:F6}, {floorMaxZ:F6}), " +
                                                $"PlacementPoint=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
                                        }
                                        
                                        // ✅ Skip rotation transformation for mixed orientations (already have world-space bounds)
                                        // Set depth from StructuralElementThickness if not already set
                                        if (cornerDepth <= 0)
                                        {
                                            var firstCzForDepth = GetCachedClashZone(sleeveIdsInCluster[0], xmlFilePath);
                                            if (firstCzForDepth != null && firstCzForDepth.StructuralElementThickness > 0)
                                            {
                                                cornerDepth = firstCzForDepth.StructuralElementThickness;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // ✅ FALLBACK: If not enough corners, continue with corner-based calculation
                                        hasMixedOrientations = false;
                                    }
                                }
                                
                                // ✅ Only do corner-based calculation if NOT mixed orientations
                                if (!hasMixedOrientations)
                                {
                                    // ✅ CRITICAL FIX: For straight axis (0°), use corner centroid directly as placement point
                                    // For rotated axis, use intersection point centroid (rotation transformation accounts for difference)
                                    bool isStraightAxis = Math.Abs(rotationAngle) < 1e-6;
                                    
                                    if (isStraightAxis)
                                    {
                                        // ✅ STRAIGHT AXIS FLOOR: 
                                        // - Z coordinate: Get from first sleeve in cluster
                                        // - X coordinate: Midpoint of X extents from corners
                                        // - Y coordinate: Midpoint of Y extents from corners
                                        
                                        // Calculate width/height directly from corner extents (no rotation)
                                        double wcsMinX = allCorners.Min(c => c.X);
                                        double wcsMaxX = allCorners.Max(c => c.X);
                                        double wcsMinY = allCorners.Min(c => c.Y);
                                        double wcsMaxY = allCorners.Max(c => c.Y);
                                        
                                        // X and Y midpoints from corner extents
                                        double midX = (wcsMinX + wcsMaxX) / 2.0;
                                        double midY = (wcsMinY + wcsMaxY) / 2.0;
                                        
                                        // Z from first sleeve in cluster
                                        double placementZ = 0.0;
                                        if (sleeveIdsInCluster.Count > 0)
                                        {
                                            int firstSleeveId = sleeveIdsInCluster[0];
                                            var firstCz = GetCachedClashZone(firstSleeveId, xmlFilePath);
                                            if (firstCz != null)
                                            {
                                                placementZ = firstCz.SleevePlacementPointActiveDocumentZ;
                                                if (placementZ == 0.0)
                                                    placementZ = firstCz.IntersectionPointZ;
                                                
                                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                    $"[{DateTime.Now:HH:mm:ss}]   ✅ STRAIGHT AXIS FLOOR: Got Z from first sleeve {firstSleeveId}: " +
                                                    $"SleevePlacementPointZ={firstCz.SleevePlacementPointActiveDocumentZ:F6}, " +
                                                    $"IntersectionPointZ={firstCz.IntersectionPointZ:F6}, " +
                                                    $"Final placementZ={placementZ:F6}\n");
                                            }
                                            else
                                            {
                                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                    $"[{DateTime.Now:HH:mm:ss}]   ⚠️ STRAIGHT AXIS FLOOR: First sleeve {firstSleeveId} ClashZone is NULL!\n");
                                            }
                                        }
                                        else
                                        {
                                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                $"[{DateTime.Now:HH:mm:ss}]   ⚠️ STRAIGHT AXIS FLOOR: sleeveIdsInCluster is EMPTY! Count={sleeveIdsInCluster.Count}\n");
                                        }
                                        
                                        // ✅ CRITICAL: Override placement point with X/Y midpoints and Z from first sleeve
                                        placementPoint = new XYZ(midX, midY, placementZ);
                                        
                                        originX = midX;
                                        originY = midY;
                                        
                                        cornerWidth = wcsMaxX - wcsMinX;
                                        cornerHeight = wcsMaxY - wcsMinY;
                                        
                                        // Store bounds for extents
                                        rotatedMinX = wcsMinX;
                                        rotatedMaxX = wcsMaxX;
                                        rotatedMinY = wcsMinY;
                                        rotatedMaxY = wcsMaxY;
                                        
                                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                            $"[{DateTime.Now:HH:mm:ss}]   ✅ STRAIGHT AXIS FLOOR: X/Y from corner extents midpoints, Z from first sleeve. " +
                                            $"PlacementPoint=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6}), " +
                                            $"X: mid={midX:F6} (min={wcsMinX:F6}, max={wcsMaxX:F6}), " +
                                            $"Y: mid={midY:F6} (min={wcsMinY:F6}, max={wcsMaxY:F6}), " +
                                            $"Z: from first sleeve={placementZ:F6}, " +
                                            $"Width={cornerWidth:F6}, Height={cornerHeight:F6}, allCorners.Count={allCorners.Count}\n");
                                    }
                                    else
                                    {
                                        // ✅ ROTATED AXIS FLOOR: Use rotation transformation
                                        double cosA = Math.Cos(-rotationAngle); // Negative for inverse rotation
                                        double sinA = Math.Sin(-rotationAngle);
                                        
                                        // ✅ Use corner centroid as rotation origin for accurate geometric center
                                        // Corner centroid works correctly when MEP elements have different orientations
                                        originX = allCorners.Average(c => c.X);
                                        originY = allCorners.Average(c => c.Y);
                                        
                                        // ✅ CRITICAL FIX: Update placement point to match rotation origin (centroid)
                                        // This ensures the cluster is placed exactly at the centroid of the corners used for sizing, preventing misalignment
                                        double placementZ = placementPoint.Z;
                                        
                                        // Consistency with Straight Axis logic: fetch Z from first sleeve to ensure alignment with host plane
                                        if (sleeveIdsInCluster.Count > 0)
                                        {
                                            var firstCz = GetCachedClashZone(sleeveIdsInCluster[0], xmlFilePath);
                                            if (firstCz != null)
                                            {
                                                double z = firstCz.SleevePlacementPointActiveDocumentZ;
                                                if (z == 0.0) z = firstCz.IntersectionPointZ;
                                                if (z != 0.0) placementZ = z;
                                            }
                                        }
                                        
                                        placementPoint = new XYZ(originX, originY, placementZ);
                                        
                                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                            $"[{DateTime.Now:HH:mm:ss}]   Origin (corner centroid): ({originX:F6}, {originY:F6}), PlacementPoint=({placementPoint.X:F6}, {placementPoint.Y:F6}), cosA={cosA:F6}, sinA={sinA:F6}\n");
                                        
                                        // Rotate each corner and find min/max in rotated space
                                        double minRotX = double.MaxValue, maxRotX = double.MinValue;
                                        double minRotY = double.MaxValue, maxRotY = double.MinValue;
                                        
                                        foreach (var corner in allCorners)
                                        {
                                            // Translate to origin
                                            double relX = corner.X - originX;
                                            double relY = corner.Y - originY;
                                            
                                            // Rotate
                                            double rotX = relX * cosA - relY * sinA;
                                            double rotY = relX * sinA + relY * cosA;
                                            
                                            // Update bounds
                                            minRotX = Math.Min(minRotX, rotX);
                                            maxRotX = Math.Max(maxRotX, rotX);
                                            minRotY = Math.Min(minRotY, rotY);
                                            maxRotY = Math.Max(maxRotY, rotY);
                                        }
                                        
                                        cornerWidth = maxRotX - minRotX;
                                        cornerHeight = maxRotY - minRotY;
                                        
                                        // Store rotated bounds for extents
                                        rotatedMinX = minRotX;
                                        rotatedMaxX = maxRotX;
                                        rotatedMinY = minRotY;
                                        rotatedMaxY = maxRotY;
                                        
                                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                            $"[{DateTime.Now:HH:mm:ss}]   Rotated bounds: minX={minRotX:F6}, maxX={maxRotX:F6}, minY={minRotY:F6}, maxY={maxRotY:F6}\n");
                                    }
                                }
                            }
                            
                            // ✅ HEIGHT: Calculate from Z range of bounding boxes (vertical dimension)
                            // For walls/framing: Height = Z range (vertical), NOT from corner Z (all corners have same Z for 2D opening)
                            // For floors/other: Height = Y range (vertical in rotated space)
                            var clashZonesForHeight = sleeveIdsInCluster
                                .Select(id => GetCachedClashZone(id, xmlFilePath))
                                .Where(cz => cz != null)
                                .ToList();
                            
                            double cornerMinZ = clashZonesForHeight.Count > 0 ? clashZonesForHeight.Min(cz => cz.SleeveBoundingBoxMinZ) : 0.0;
                            double cornerMaxZ = clashZonesForHeight.Count > 0 ? clashZonesForHeight.Max(cz => cz.SleeveBoundingBoxMaxZ) : 0.0;
                            double calculatedHeight = cornerMaxZ - cornerMinZ; // Height = Z range (vertical dimension)
                            
                            // ✅ CRITICAL FIX: For walls/framing, use Z range for height, not RCS Y
                            // For floors/other, cornerHeight is already calculated from Y range in rotated space
                            if (isWallOrFraming)
                            {
                                cornerHeight = calculatedHeight; // Override with Z range (vertical) for walls/framing
                            }
                            else
                            {
                                // For floors/other, cornerHeight is already correct from rotated Y range
                                // But also verify Z range matches (should be same for 2D opening)
                                if (Math.Abs(cornerHeight - calculatedHeight) > 0.001) // 0.001 feet = ~0.3mm tolerance
                                {
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ⚠️ Height mismatch: Y-range={cornerHeight:F6}ft, Z-range={calculatedHeight:F6}ft, using Y-range\n");
                                }
                            }
                            
                            // ✅ DEPTH: Calculate based on host type (only if not already set for mixed orientations)
                            // For walls/framing: depth = wall/framing thickness (overridden later in SetSizeParameters)
                            // For floors/other: depth = StructuralElementThickness (same as individual sleeves)
                            if (cornerDepth <= 0)
                            {
                                if (isWallOrFraming)
                                {
                                    // ✅ WALL/FRAMING: Depth will be set from wall thickness
                                    // Retrieve max thickness from constituent clash zones
                                    double maxWallThickness = 0.0;
                                    bool isFloorElement = false;
                                    
                                    // Check if any element is a floor (just in case isWallOrFraming flag is misleading)
                                    foreach (var cz in clashZonesForHeight)
                                    {
                                        if (cz == null) continue;
                                        if (!string.IsNullOrEmpty(cz.StructuralElementType) && 
                                            cz.StructuralElementType.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            isFloorElement = true;
                                            break;
                                        }
                                    }

                                    foreach (var cz in clashZonesForHeight)
                                    {
                                        if (cz == null) continue;
                                        
                                        double t = 0.0;
                                        if (isFloorElement)
                                        {
                                            // ✅ FIXED: For floors, use StructuralElementThickness
                                            t = cz.StructuralElementThickness;
                                        }
                                        else
                                        {
                                            // For walls, prefer WallThickness
                                            t = cz.WallThickness > 0 ? cz.WallThickness : cz.StructuralElementThickness;
                                        }
                                        
                                        if (t > maxWallThickness) maxWallThickness = t;
                                    }
                                    
                                    cornerDepth = maxWallThickness > 0 ? maxWallThickness : 0.0;
                                    
                                    if (!DeploymentConfiguration.DeploymentMode)
                                    {
                                        string depthType = isFloorElement ? "FLOOR (Structural)" : "WALL/FRAMING (WallThick)";
                                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                            $"[{DateTime.Now:HH:mm:ss}]   ✅ {depthType} DEPTH: Using max thickness={cornerDepth * 304.8:F1}mm (calculated from {clashZonesForHeight.Count} clash zones)\n");
                                    }
                                }
                                else
                                {
                                    // ✅ FLOOR/OTHER: Get depth from StructuralElementThickness (same as individual sleeves)
                                    // Use the first clash zone's StructuralElementThickness (all sleeves in cluster should have same host)
                                    var firstClashZoneForDepth = clashZonesForHeight.FirstOrDefault();
                                    if (firstClashZoneForDepth != null && firstClashZoneForDepth.StructuralElementThickness > 0)
                                    {
                                        cornerDepth = firstClashZoneForDepth.StructuralElementThickness;
                                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                            $"[{DateTime.Now:HH:mm:ss}]   ✅ FLOOR DEPTH: Using StructuralElementThickness={cornerDepth * 304.8:F1}mm from first sleeve\n");
                                    }
                                    else
                                    {
                                        // ⚠️ FALLBACK: If StructuralElementThickness is missing, log warning
                                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                            $"[{DateTime.Now:HH:mm:ss}]   ⚠️ FLOOR DEPTH: StructuralElementThickness not found or zero, depth will be 0.0mm\n");
                                    }
                                }
                            }
                            
                            double widthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerWidth);
                            double heightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerHeight);
                            double depthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerDepth);
                            
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ CORNER-BASED SUCCESS: W={widthMm:F1}mm, H={heightMm:F1}mm, D={depthMm:F1}mm (from {allCorners.Count} corners)\n");
                            
                            // ✅ Calculate rotated bounding box extents in world coordinates
                            // For straight axis floors: bounds are already in world space (no need to add origin)
                            // For rotated axis floors: bounds are relative to origin, need to add origin back
                            // For walls/framing: bounds are already in RCS coordinates (no need to add origin)
                            if (!isWallOrFraming)
                            {
                                bool isStraightAxis = Math.Abs(rotationAngle) < 1e-6;
                                if (!isStraightAxis)
                                {
                                    // ✅ ROTATED AXIS FLOOR: Add origin back to rotated bounds (they're relative to origin)
                                    rotatedMinX = rotatedMinX + originX;
                                    rotatedMinY = rotatedMinY + originY;
                                    rotatedMaxX = rotatedMaxX + originX;
                                    rotatedMaxY = rotatedMaxY + originY;
                                }
                                // ✅ STRAIGHT AXIS FLOOR: Bounds are already in world space (no transformation needed)
                            }
                            
                            (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ) cornerResult = 
                                (cornerWidth, cornerHeight, cornerDepth, placementPoint,
                                rotatedMinX, rotatedMinY, cornerMinZ,
                                rotatedMaxX, rotatedMaxY, cornerMaxZ);
                            
                            StoreInCache(cluster, rotationAngle, cornerResult, calcStopwatch);
                            return cornerResult;
                        }
                        else
                        {
                            // ⚠️ CRITICAL ERROR: Not enough corners - this should NEVER happen if corners are saved correctly
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ CRITICAL ERROR: Not enough corners collected: {allCorners.Count} (need at least 4)\n" +
                                $"  This indicates corners were NOT saved correctly for individual sleeves!\n" +
                                $"  Sleeve IDs in cluster: {string.Join(", ", sleeveIdsInCluster)}\n" +
                                $"  Check if MEP_ElementId is set on individual sleeves and if corners are saved to database.\n" +
                                $"  Check if SleevePersistenceService.PersistSleeveData was called after placement.\n");
                            
                            // ✅ FALLBACK FOR STRAIGHT-AXIS FLOOR CLUSTERS: Use individual sleeve placement points if corners are missing
                            // This ensures cluster sleeves are still placed correctly even if corner saving failed
                            bool isStraightAxis = Math.Abs(rotationAngle) < 1e-6;
                            bool isFloorHost = !isWallOrFraming;
                            
                            if (isStraightAxis && isFloorHost && placementPoint.IsZeroLength())
                            {
                                // ✅ FALLBACK: Calculate placement point from individual sleeve placement points
                                var placementPoints = new List<XYZ>();
                                foreach (int sleeveId in sleeveIdsInCluster)
                                {
                                    var cz = GetCachedClashZone(sleeveId, xmlFilePath);
                                    if (cz != null)
                                    {
                                        double px = cz.SleevePlacementPointActiveDocumentX;
                                        double py = cz.SleevePlacementPointActiveDocumentY;
                                        double pz = cz.SleevePlacementPointActiveDocumentZ;
                                        
                                        if (px != 0.0 || py != 0.0 || pz != 0.0)
                                        {
                                            placementPoints.Add(new XYZ(px, py, pz));
                                        }
                                    }
                                }
                                
                                if (placementPoints.Count > 0)
                                {
                                    placementPoint = new XYZ(
                                        placementPoints.Average(p => p.X),
                                        placementPoints.Average(p => p.Y),
                                        placementPoints.Average(p => p.Z)
                                    );
                                    
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ FALLBACK: Calculated placement point from {placementPoints.Count} individual sleeve placement points: " +
                                        $"({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
                                }
                            }
                            
                            // ❌ NO FALLBACK: Throw exception to force investigation - corners MUST be available
                            throw new InvalidOperationException(
                                $"CRITICAL: Corner-based calculation failed - only {allCorners.Count} corners found (need at least 4). " +
                                $"This indicates corners were not saved correctly for individual sleeves. " +
                                $"Sleeve IDs: {string.Join(", ", sleeveIdsInCluster)}. " +
                                $"Check MEP_ElementId parameter on sleeves and corner saving logic (SleevePersistenceService.PersistSleeveData).");
                        }
                    }
                    else
                    {
                        // ⚠️ CRITICAL ERROR: No corners found - this should NEVER happen
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ CRITICAL ERROR: No corners found in ANY sleeve!\n" +
                            $"  This indicates corners were NOT saved during individual sleeve placement!\n" +
                            $"  Sleeve IDs in cluster: {string.Join(", ", sleeveIdsInCluster)}\n" +
                            $"  Check if MEP_ElementId is set on individual sleeves and if corners are saved to database.\n" +
                            $"  Check if SleevePersistenceService.PersistSleeveData was called after placement.\n");
                        
                        // ❌ NO FALLBACK: Throw exception to force investigation - corners MUST be available
                        throw new InvalidOperationException(
                            $"CRITICAL: Corner-based calculation failed - no corners found in any sleeve. " +
                            $"This indicates corners were not saved during individual sleeve placement. " +
                            $"Sleeve IDs: {string.Join(", ", sleeveIdsInCluster)}. " +
                            $"Check MEP_ElementId parameter on sleeves and corner saving logic (SleevePersistenceService.PersistSleeveData).");
                    }
                }
                catch (Exception cornerEx)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ CRITICAL EXCEPTION in corner-based calculation: {cornerEx.Message}\n" +
                        $"StackTrace: {cornerEx.StackTrace}\n" +
                        $"Sleeve IDs in cluster: {string.Join(", ", sleeveIdsInCluster)}\n" +
                        $"This should NOT happen - corners MUST be available for all sleeves.\n");
                    
                    // ❌ NO FALLBACK: Re-throw to force investigation
                    throw;
                }
            }
            // Note: Corner-based calculation is now always attempted and MUST succeed (no fallback to RCS)

            // ✅ CRITICAL FIX: For rotated floor clusters, if corner-based calculation failed,
            // reconstruct world bounding boxes from placement points + dimensions (same as straight-axis logic)
            // This ensures correct sizing for rotated clusters when corner-based fails
            bool shouldReconstructFromPlacementPoints = false;
            if (rotatedBboxes.Count == 0 || (Math.Abs(rotationAngle) > 1e-6 && rotatedBboxes.Count < cluster.Count))
            {
                // Check if this is a floor cluster
                var checkFirstSleeve = cluster.FirstOrDefault();
                if (checkFirstSleeve != null)
                {
                    var checkCz = GetCachedClashZone(checkFirstSleeve.SleeveInstanceId, xmlFilePath);
                    if (checkCz != null)
                    {
                        bool isFloorHost = string.Equals(checkCz.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase) ||
                                          string.Equals(checkCz.StructuralElementType, "Floors", StringComparison.OrdinalIgnoreCase);
                        if (isFloorHost && Math.Abs(rotationAngle) > 1e-6)
                        {
                            // ✅ For rotated floor clusters, always reconstruct from placement points
                            // The stored rotated bounding boxes are in individual sleeve local coordinates,
                            // not cluster-aligned coordinates, so unioning them gives incorrect results
                            shouldReconstructFromPlacementPoints = true;
                        }
                    }
                }
            }

            if (rotatedBboxes.Count == 0 && !shouldReconstructFromPlacementPoints)
            {
                // ✅ CRITICAL: For walls/framing, check if we should use RCS or transform WCS to RCS
                var fallbackFirstSleeveData = cluster[0];
                if (fallbackFirstSleeveData != null)
                {
                    try
                    {
                        int firstSleeveInstanceId = fallbackFirstSleeveData.SleeveInstanceId;
                        if (firstSleeveInstanceId > 0)
                        {
                            var firstCz = _getClashZoneFunc(firstSleeveInstanceId, xmlFilePath) as ClashZone;
                            if (firstCz != null)
                            {
                                bool isWallHost = string.Equals(firstCz.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                                  string.Equals(firstCz.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                                bool isFramingHost = string.Equals(firstCz.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
                                
                                // ✅ For walls/framing with 0° rotation, transform WCS bounding boxes to RCS on-the-fly
                                if ((isWallHost || isFramingHost) && firstCz.WallDirection != null && !firstCz.WallDirection.IsZeroLength())
                                {
                                    // Transform WCS bounding boxes to RCS
                                    var wcsBboxes = new List<(XYZ min, XYZ max)>();
                                    foreach (var sleeveData in cluster)
                                    {
                                        var cz = GetCachedClashZone(sleeveData.SleeveInstanceId, xmlFilePath);
                                        if (cz != null && 
                                            (cz.SleeveBoundingBoxMinX != 0 || cz.SleeveBoundingBoxMaxX != 0 ||
                                             cz.SleeveBoundingBoxMinY != 0 || cz.SleeveBoundingBoxMaxY != 0 ||
                                             cz.SleeveBoundingBoxMinZ != 0 || cz.SleeveBoundingBoxMaxZ != 0))
                                        {
                                            var wcsBbox = new BoundingBoxXYZ
                                            {
                                                Min = new XYZ(cz.SleeveBoundingBoxMinX, cz.SleeveBoundingBoxMinY, cz.SleeveBoundingBoxMinZ),
                                                Max = new XYZ(cz.SleeveBoundingBoxMaxX, cz.SleeveBoundingBoxMaxY, cz.SleeveBoundingBoxMaxZ)
                                            };
                                            
                                            // Transform to RCS
                                            var rcsBbox = WallRcsTransformer.TransformToRcs(wcsBbox, firstCz.WallDirection);
                                            if (rcsBbox != null)
                                            {
                                                wcsBboxes.Add((rcsBbox.Min, rcsBbox.Max));
                                            }
                                        }
                                    }
                                    
                                    if (wcsBboxes.Count > 0)
                                    {
                                        // Union in RCS
                                        double rcsMinX = wcsBboxes.Min(b => b.min.X);
                                        double rcsMinY = wcsBboxes.Min(b => b.min.Y);
                                        double rcsMinZ = wcsBboxes.Min(b => b.min.Z);
                                        double rcsMaxX = wcsBboxes.Max(b => b.max.X);
                                        double rcsMaxY = wcsBboxes.Max(b => b.max.Y);
                                        double rcsMaxZ = wcsBboxes.Max(b => b.max.Z);
                                        
                                        double rcsWidth = rcsMaxX - rcsMinX;
                                        double rcsHeight = rcsMaxZ - rcsMinZ;
                                        double rcsDepth = rcsMaxY - rcsMinY;
                                        
                                        XYZ rcsMid = new XYZ((rcsMinX + rcsMaxX) / 2, (rcsMinY + rcsMaxY) / 2, (rcsMinZ + rcsMaxZ) / 2);
                                        XYZ wcsMid = WallRcsTransformer.TransformToWcs(rcsMid, firstCz.WallDirection,
                                            new XYZ(firstCz.SleevePlacementPointActiveDocumentX,
                                                   firstCz.SleevePlacementPointActiveDocumentY,
                                                   firstCz.SleevePlacementPointActiveDocumentZ)) ?? XYZ.Zero;
                                        
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            double wMm = RevitUnitConversionService.Instance.FromInternalMillimeters(rcsWidth);
                                            double hMm = RevitUnitConversionService.Instance.FromInternalMillimeters(rcsHeight);
                                            double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(rcsDepth);
                                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                $"[{DateTime.Now:HH:mm:ss}] ✅ RCS TRANSFORM: Converted WCS to RCS for wall cluster: W={wMm:F1}mm, H={hMm:F1}mm, D={dMm:F1}mm, PlacementPoint=({placementPoint.X:F1}, {placementPoint.Y:F1}, {placementPoint.Z:F1})\n");
                                        }
                                        
                                        // ✅ FIX: Use intersection point centroid for placement, not bounding box midpoint
                                        (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ) rcsResult = 
                                            (rcsWidth, rcsHeight, rcsDepth, placementPoint, null, null, null, null, null, null);
                                        StoreInCache(cluster, rotationAngle, rcsResult, calcStopwatch);
                                        return rcsResult;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception rcsEx)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error transforming to RCS: {rcsEx.Message}\n");
                    }
                }
                
                // ✅ CRITICAL FIX: For floor clusters, build envelope from placement points + dimensions
                // This works for both straight-axis (0°) and rotated-axis (135°, 315°, etc.) clusters
                // For rotated clusters, we reconstruct world bounding boxes from placement points + dimensions
                // because the stored bounding boxes are in individual sleeve local coordinates, not cluster-aligned coordinates
                bool isStraightAxis = Math.Abs(rotationAngle) < 1e-6;
                if (isStraightAxis || (Math.Abs(rotationAngle) > 1e-6 && rotatedBboxes.Count == 0)) // Straight axis OR rotated axis with no rotated bboxes
                {
                    try
                    {
                        var floorFirstSleeveData = cluster[0];
                        if (floorFirstSleeveData != null)
                        {
                            int firstSleeveInstanceId = floorFirstSleeveData.SleeveInstanceId;
                            if (firstSleeveInstanceId > 0)
                            {
                                var firstCz = GetCachedClashZone(firstSleeveInstanceId, xmlFilePath);
                                if (firstCz != null)
                                {
                                    bool isFloorHost = string.Equals(firstCz.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase) ||
                                                       string.Equals(firstCz.StructuralElementType, "Floors", StringComparison.OrdinalIgnoreCase);
                                    
                                    if (isFloorHost)
                                    {
                                        // Check if cluster has mixed X/Y orientations
                                        var orientations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                        foreach (var sleeveData in cluster)
                                        {
                                            var cz = GetCachedClashZone(sleeveData.SleeveInstanceId, xmlFilePath);
                                            if (cz != null)
                                            {
                                                string orientation = cz.MepElementOrientationDirection;
                                                
                                                // ✅ Fallback: Infer orientation from rotation angle if missing
                                                if (string.IsNullOrEmpty(orientation))
                                                {
                                                    double angle = Math.Abs(cz.MepElementRotationAngle); // Radians
                                                    // Normalize to 0-PI
                                                    while (angle > Math.PI) angle -= Math.PI;
                                                    
                                                    if (angle > Math.PI / 4.0 && angle < 3.0 * Math.PI / 4.0)
                                                    {
                                                        orientation = "Y"; // ~90 degrees
                                                    }
                                                    else
                                                    {
                                                        orientation = "X"; // ~0 or ~180 degrees
                                                    }
                                                }
                                                
                                                if (!string.IsNullOrEmpty(orientation))
                                                {
                                                    orientations.Add(orientation);
                                                }
                                            }
                                        }
                                        
                                        // ✅ DIAGNOSTIC: Log orientation detection
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                $"[{DateTime.Now:HH:mm:ss}] 🔍 FLOOR ORIENTATION CHECK: Found orientations=[{string.Join(", ", orientations)}], ClusterSize={cluster.Count}, HasX={orientations.Contains("X")}, HasY={orientations.Contains("Y")}\n");
                                        }
                                        
                                        // If we have both X and Y orientations, build envelope from calculated world bboxes
                                        if (orientations.Contains("X") && orientations.Contains("Y"))
                                        {
                                            var floorBboxes = new List<(XYZ min, XYZ max)>();
                                            foreach (var sleeveData in cluster)
                                            {
                                                var cz = GetCachedClashZone(sleeveData.SleeveInstanceId, xmlFilePath);
                                                if (cz != null)
                                                {
                                                    // ✅ Calculate World Bounding Box from Placement Point + Dimensions
                                                    double placementX = cz.SleevePlacementPointActiveDocumentX;
                                                    double placementY = cz.SleevePlacementPointActiveDocumentY;
                                                    double placementZ = cz.SleevePlacementPointActiveDocumentZ;
                                                    
                                                    // ✅ CRITICAL FIX: Get dimensions in feet (stored in feet, no conversion needed)
                                                    // SleeveWidth and SleeveHeight are stored in Revit internal units (feet), not millimeters
                                                    double sleeveWidth = cz.SleeveWidth;
                                                    double sleeveHeight = cz.SleeveHeight;
                                                    // Use StructuralElementThickness for depth if available, otherwise default to 1.0ft
                                                    // StructuralElementThickness is also in feet (Revit internal units)
                                                    double sleeveDepth = cz.StructuralElementThickness > 0 ? cz.StructuralElementThickness : 1.0; 

                                                    double halfWidth, halfHeight;

                                                    // Determine orientation (use explicit or inferred)
                                                    bool isY = string.Equals(cz.MepElementOrientationDirection, "Y", StringComparison.OrdinalIgnoreCase);
                                                    if (!isY && string.IsNullOrEmpty(cz.MepElementOrientationDirection))
                                                    {
                                                        // Fallback inference
                                                        double angle = Math.Abs(cz.MepElementRotationAngle);
                                                        while (angle > Math.PI) angle -= Math.PI;
                                                        if (angle > Math.PI / 4.0 && angle < 3.0 * Math.PI / 4.0)
                                                        {
                                                            isY = true;
                                                        }
                                                    }

                                                    // Apply rotation based on orientation
                                                    if (isY)
                                                    {
                                                        // Rotated 90 degrees: Width is along Y, Height is along X
                                                        halfWidth = sleeveHeight / 2.0;  // X-dimension
                                                        halfHeight = sleeveWidth / 2.0;  // Y-dimension
                                                    }
                                                    else
                                                    {
                                                        // Default X orientation: Width is along X, Height is along Y
                                                        halfWidth = sleeveWidth / 2.0;   // X-dimension
                                                        halfHeight = sleeveHeight / 2.0; // Y-dimension
                                                    }
                                                    double halfDepth = sleeveDepth / 2.0;

                                                    floorBboxes.Add((
                                                        new XYZ(placementX - halfWidth, placementY - halfHeight, placementZ - halfDepth),
                                                        new XYZ(placementX + halfWidth, placementY + halfHeight, placementZ + halfDepth)
                                                    ));
                                                }
                                            }
                                            
                                            if (floorBboxes.Count > 0)
                                            {
                                                // Union of calculated world sleeve bounding boxes
                                                double floorMinX = floorBboxes.Min(b => b.min.X);
                                                double floorMinY = floorBboxes.Min(b => b.min.Y);
                                                double floorMinZ = floorBboxes.Min(b => b.min.Z);
                                                double floorMaxX = floorBboxes.Max(b => b.max.X);
                                                double floorMaxY = floorBboxes.Max(b => b.max.Y);
                                                double floorMaxZ = floorBboxes.Max(b => b.max.Z);
                                                
                                                double floorWidth = floorMaxX - floorMinX;
                                                double floorHeight = floorMaxY - floorMinY;
                                                double floorDepth = floorMaxZ - floorMinZ;
                                                
                                                // ✅ Calculate new center point from the union envelope
                                                XYZ newPlacementPoint = new XYZ(
                                                    (floorMinX + floorMaxX) / 2.0,
                                                    (floorMinY + floorMaxY) / 2.0,
                                                    (floorMinZ + floorMaxZ) / 2.0
                                                );

                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    double wMm = RevitUnitConversionService.Instance.FromInternalMillimeters(floorWidth);
                                                    double hMm = RevitUnitConversionService.Instance.FromInternalMillimeters(floorHeight);
                                                    double dMm = RevitUnitConversionService.Instance.FromInternalMillimeters(floorDepth);
                                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                        $"[{DateTime.Now:HH:mm:ss}] ✅ FLOOR MIXED X/Y: Building envelope from CALCULATED WORLD bboxes: W={wMm:F1}mm, H={hMm:F1}mm, D={dMm:F1}mm, " +
                                                        $"Orientations=[{string.Join(", ", orientations)}], Sleeves={floorBboxes.Count}, " +
                                                        $"OldCenter=({placementPoint.X:F2},{placementPoint.Y:F2}), NewCenter=({newPlacementPoint.X:F2},{newPlacementPoint.Y:F2})\n");
                                                }
                                                
                                                (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ) floorResult = 
                                                    (floorWidth, floorHeight, floorDepth, newPlacementPoint, null, null, null, null, null, null);
                                                StoreInCache(cluster, rotationAngle, floorResult, calcStopwatch);
                                                return floorResult;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception floorEx)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error handling floor mixed X/Y cluster: {floorEx.Message}\n");
                    }
                }
                
                // ✅ CRITICAL FIX: For floor clusters, calculate envelope from placement points + sleeve dimensions
                // The stored bounding boxes are in sleeve local coordinates, not world coordinates!
                // We need to build the envelope from each sleeve's placement point + its actual dimensions
                // This applies to both straight-axis (0°) and rotated-axis (135°, 315°, etc.) floor clusters
                var dbBboxes = new List<(XYZ min, XYZ max)>();
                bool isFloorClusterProcessing = false;

                // Check if this is a floor cluster first to decide processing logic
                var firstSleeve = cluster.FirstOrDefault();
                if (firstSleeve != null)
                {
                    var firstCz = GetCachedClashZone(firstSleeve.SleeveInstanceId, xmlFilePath);
                    if (firstCz != null)
                    {
                        isFloorClusterProcessing = string.Equals(firstCz.StructuralElementType, "Floor", StringComparison.OrdinalIgnoreCase) ||
                                                 string.Equals(firstCz.StructuralElementType, "Floors", StringComparison.OrdinalIgnoreCase);
                    }
                }

                // ✅ CRITICAL: If corner-based failed for rotated floor cluster, force reconstruction from placement points
                if (shouldReconstructFromPlacementPoints)
                {
                    isFloorClusterProcessing = true;
                }

                foreach (var sleeveData in cluster)
                {
                    var clashZone = GetCachedClashZone(sleeveData.SleeveInstanceId, xmlFilePath);
                    if (clashZone == null) continue;
                    
                    if (isFloorClusterProcessing)
                    {
                        // ✅ For floor clusters: Use placement point + sleeve dimensions to build world-space bbox
                        double placementX = clashZone.SleevePlacementPointActiveDocumentX;
                        double placementY = clashZone.SleevePlacementPointActiveDocumentY;
                        double placementZ = clashZone.SleevePlacementPointActiveDocumentZ;
                        
                        // Get dimensions in feet (stored in feet, no conversion needed)
                        double sleeveWidth = clashZone.SleeveWidth;
                        double sleeveHeight = clashZone.SleeveHeight;
                        // Use StructuralElementThickness for depth if available, otherwise default to 1.0ft
                        double sleeveDepth = clashZone.StructuralElementThickness > 0 ? clashZone.StructuralElementThickness : 1.0;

                        double halfWidth, halfHeight;

                        // Apply rotation based on orientation
                        if (string.Equals(clashZone.MepElementOrientationDirection, "Y", StringComparison.OrdinalIgnoreCase))
                        {
                            // Rotated 90 degrees: Width is along Y, Height is along X
                            halfWidth = sleeveHeight / 2.0;  // X-dimension
                            halfHeight = sleeveWidth / 2.0;  // Y-dimension
                        }
                        else
                        {
                            // Default X orientation: Width is along X, Height is along Y
                            halfWidth = sleeveWidth / 2.0;   // X-dimension
                            halfHeight = sleeveHeight / 2.0; // Y-dimension
                        }
                        double halfDepth = sleeveDepth / 2.0;
                        
                        dbBboxes.Add((
                            new XYZ(placementX - halfWidth, placementY - halfHeight, placementZ - halfDepth),
                            new XYZ(placementX + halfWidth, placementY + halfHeight, placementZ + halfDepth)
                        ));
                    }
                    else
                    {
                        // For walls/framing, use stored bounding boxes (already in world coordinates)
                        if (clashZone.SleeveBoundingBoxMinX != 0 || clashZone.SleeveBoundingBoxMaxX != 0 ||
                            clashZone.SleeveBoundingBoxMinY != 0 || clashZone.SleeveBoundingBoxMaxY != 0 ||
                            clashZone.SleeveBoundingBoxMinZ != 0 || clashZone.SleeveBoundingBoxMaxZ != 0)
                        {
                            dbBboxes.Add((
                                new XYZ(clashZone.SleeveBoundingBoxMinX, clashZone.SleeveBoundingBoxMinY, clashZone.SleeveBoundingBoxMinZ),
                                new XYZ(clashZone.SleeveBoundingBoxMaxX, clashZone.SleeveBoundingBoxMaxY, clashZone.SleeveBoundingBoxMaxZ)
                            ));
                        }
                    }
                }
                
                // If database has bounding boxes, use them (faster, no Revit API calls)
                if (dbBboxes.Count > 0)
                {
                    double minXf = dbBboxes.Min(b=>b.min.X); double minYf = dbBboxes.Min(b=>b.min.Y); double minZf = dbBboxes.Min(b=>b.min.Z);
                    double maxXf = dbBboxes.Max(b=>b.max.X); double maxYf = dbBboxes.Max(b=>b.max.Y); double maxZf = dbBboxes.Max(b=>b.max.Z);
                    
                    double wf = maxXf - minXf;
                    double hf = maxYf - minYf;
                    double df = maxZf - minZf;
                    
                    // ✅ FIX: Use intersection point centroid for placement, not bounding box midpoint
                    // UNLESS it's a floor cluster, then use the calculated center
                    XYZ finalPlacementPoint = placementPoint;
                    if (isFloorClusterProcessing)
                    {
                         finalPlacementPoint = new XYZ(
                            (minXf + maxXf) / 2.0,
                            (minYf + maxYf) / 2.0,
                            (minZf + maxZf) / 2.0
                        );
                    }

                    // ✅ FIX: Convert to millimeters for logging
                    double wfMm = RevitUnitConversionService.Instance.FromInternalMillimeters(wf);
                    double hfMm = RevitUnitConversionService.Instance.FromInternalMillimeters(hf);
                    double dfMm = RevitUnitConversionService.Instance.FromInternalMillimeters(df);
                    
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ✅ FALLBACK (axis-aligned from DB): W={wfMm:F1}mm, H={hfMm:F1}mm, D={dfMm:F1}mm ({dbBboxes.Count} sleeves from database, isFloor={isFloorClusterProcessing})\n");
                    
                    return (wf,hf,df,finalPlacementPoint,null,null,null,null,null,null);
                }
                
                // Last resort: Fallback to Revit bounding boxes if database data is missing (only if actualSleeves provided)
                if (actualSleeves != null && actualSleeves.Count > 0)
                {
                    var revitBboxes = actualSleeves.Select(s => s.get_BoundingBox(null)).Where(b => b != null && b.Enabled).ToList();
                    if (revitBboxes.Count > 0)
                    {
                        double minXr = revitBboxes.Min(b=>b.Min.X); double minYr = revitBboxes.Min(b=>b.Min.Y); double minZr = revitBboxes.Min(b=>b.Min.Z);
                        double maxXr = revitBboxes.Max(b=>b.Max.X); double maxYr = revitBboxes.Max(b=>b.Max.Y); double maxZr = revitBboxes.Max(b=>b.Max.Z);
                        double wr = maxXr - minXr; double hr = maxYr - minYr; double dr = maxZr - minZr;
                        
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ FALLBACK (axis-aligned from Revit): W={wr:F1}mm, H={hr:F1}mm, D={dr:F1}mm (database data missing, using Revit API)\n");
                        
                        // ✅ FIX: Use intersection point centroid for placement, not bounding box midpoint
                        return (wr,hr,dr,placementPoint,null,null,null,null,null,null);
                    }
                }
                
                // ✅ CRITICAL: If no database data AND no Revit fallback, return invalid
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL: No database data AND no Revit fallback available - cannot calculate bounding box\n");
                return (0,0,0,XYZ.Zero,null,null,null,null,null,null);
            }

            // ✅ CRITICAL FIX: For rotated axis clusters, rotated bounding boxes are in different coordinate systems
            // (each sleeve has its own rotation), so we cannot simply union them.
            // Instead, reconstruct world bounding boxes from placement points + dimensions (same as floor clusters)
            bool isRotatedAxis = Math.Abs(rotationAngle) > 1e-6;
            if (isRotatedAxis)
            {
                // ✅ For rotated axis clusters, ALWAYS reconstruct world bounding boxes from placement points + dimensions
                // This ensures correct sizing when corner-based calculation fails
                // The stored rotated bounding boxes are in world coordinates but each sleeve has its own rotation,
                // so unioning them directly gives incorrect results (e.g., 741mm × 341mm instead of 707mm × 707mm)
                var reconstructedBboxes = new List<(XYZ min, XYZ max)>();
                foreach (var sleeveData in cluster)
                {
                    var cz = GetCachedClashZone(sleeveData.SleeveInstanceId, xmlFilePath);
                    if (cz == null) continue;
                    
                    // Get placement point and dimensions
                    double placementX = cz.SleevePlacementPointActiveDocumentX;
                    double placementY = cz.SleevePlacementPointActiveDocumentY;
                    double placementZ = cz.SleevePlacementPointActiveDocumentZ;
                    
                    // Get dimensions in feet (stored in feet, no conversion needed)
                    double sleeveWidth = cz.SleeveWidth;
                    double sleeveHeight = cz.SleeveHeight;
                    double sleeveDepth = cz.StructuralElementThickness > 0 ? cz.StructuralElementThickness : 1.0;
                    
                    // Calculate half-dimensions
                    double halfWidth = sleeveWidth / 2.0;
                    double halfHeight = sleeveHeight / 2.0;
                    double halfDepth = sleeveDepth / 2.0;
                    
                    // Reconstruct world bounding box (no rotation applied - use stored WCS bbox if available)
                    bool hasStoredBbox = cz.SleeveBoundingBoxMinX != 0 || cz.SleeveBoundingBoxMaxX != 0 ||
                                         cz.SleeveBoundingBoxMinY != 0 || cz.SleeveBoundingBoxMaxY != 0 ||
                                         cz.SleeveBoundingBoxMinZ != 0 || cz.SleeveBoundingBoxMaxZ != 0;
                    
                    if (hasStoredBbox)
                    {
                        // ✅ Use stored WCS bounding box (already in world coordinates)
                        reconstructedBboxes.Add((
                            new XYZ(cz.SleeveBoundingBoxMinX, cz.SleeveBoundingBoxMinY, cz.SleeveBoundingBoxMinZ),
                            new XYZ(cz.SleeveBoundingBoxMaxX, cz.SleeveBoundingBoxMaxY, cz.SleeveBoundingBoxMaxZ)
                        ));
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ ROTATED-AXIS: Using stored WCS bbox for sleeve {sleeveData.SleeveInstanceId}: Min=({cz.SleeveBoundingBoxMinX:F6}, {cz.SleeveBoundingBoxMinY:F6}, {cz.SleeveBoundingBoxMinZ:F6}), Max=({cz.SleeveBoundingBoxMaxX:F6}, {cz.SleeveBoundingBoxMaxY:F6}, {cz.SleeveBoundingBoxMaxZ:F6})\n");
                        }
                    }
                    else
                    {
                        // ⚠️ Fallback: Calculate from placement point + dimensions (stored bbox not available)
                        reconstructedBboxes.Add((
                            new XYZ(placementX - halfWidth, placementY - halfHeight, placementZ - halfDepth),
                            new XYZ(placementX + halfWidth, placementY + halfHeight, placementZ + halfDepth)
                        ));
                        
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ ROTATED-AXIS: Stored bbox MISSING for sleeve {sleeveData.SleeveInstanceId}, reconstructing from placement point ({placementX:F6}, {placementY:F6}, {placementZ:F6}) + dimensions (W={sleeveWidth * 304.8:F1}mm, H={sleeveHeight * 304.8:F1}mm)\n");
                        }
                    }
                }
                
                if (reconstructedBboxes.Count > 0)
                {
                    // Union of reconstructed world bounding boxes
                    double reconMinX = reconstructedBboxes.Min(b => b.min.X);
                    double reconMinY = reconstructedBboxes.Min(b => b.min.Y);
                    double reconMinZ = reconstructedBboxes.Min(b => b.min.Z);
                    double reconMaxX = reconstructedBboxes.Max(b => b.max.X);
                    double reconMaxY = reconstructedBboxes.Max(b => b.max.Y);
                    double reconMaxZ = reconstructedBboxes.Max(b => b.max.Z);
                    
                    double reconWidth = reconMaxX - reconMinX;
                    double reconHeight = reconMaxY - reconMinY;
                    double reconDepth = reconMaxZ - reconMinZ;
                    
                    // Convert to millimeters for logging
                    double reconWidthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(reconWidth);
                    double reconHeightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(reconHeight);
                    double reconDepthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(reconDepth);
                    
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ✅ ROTATED-AXIS (reconstructed from WCS): W={reconWidthMm:F1}mm, H={reconHeightMm:F1}mm, D={reconDepthMm:F1}mm, Rotation={rotationAngle * 180 / Math.PI:F1}°\n");
                    
                    // ✅ FIX: Use intersection point centroid for placement, not bounding box midpoint
                    var reconResult = (reconWidth, reconHeight, reconDepth, placementPoint, reconMinX, reconMinY, reconMinZ, reconMaxX, reconMaxY, reconMaxZ);
                    
                    // Store in cache and return
                    StoreInCache(cluster, rotationAngle, reconResult, calcStopwatch);
                    return reconResult;
                }
            }
            
            // ✅ FALLBACK: Simple union of rotated boxes (only for straight-axis clusters or if reconstruction failed)
            // This is less accurate but better than nothing for straight-axis clusters
            double minX = rotatedBboxes.Min(b=>b.min.X);
            double minY = rotatedBboxes.Min(b=>b.min.Y);
            double minZ = rotatedBboxes.Min(b=>b.min.Z);
            double maxX = rotatedBboxes.Max(b=>b.max.X);
            double maxY = rotatedBboxes.Max(b=>b.max.Y);
            double maxZ = rotatedBboxes.Max(b=>b.max.Z);

            double width = maxX - minX;
            double height = maxY - minY;
            double depth = maxZ - minZ;
            XYZ mid = new XYZ((minX+maxX)/2,(minY+maxY)/2,(minZ+maxZ)/2);
            
                // Convert to millimeters for logging
                double unionWidthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(width);
                double unionHeightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(height);
                double unionDepthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(depth);
                
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ UNION (fallback - rotated bboxes): W={unionWidthMm:F1}mm, H={unionHeightMm:F1}mm, D={unionDepthMm:F1}mm, Rotation={rotationAngle * 180 / Math.PI:F1}°\n");
            
            // ✅ FIX: Use intersection point centroid for placement, not bounding box midpoint
            var result = (width,height,depth,placementPoint,minX,minY,minZ,maxX,maxY,maxZ);
            
            // ✅ PERFORMANCE: Store result in cache for future use
            try
            {
                if (cluster != null && cluster.Count > 0)
                {
                    var sleeveIds = cluster.Select(s => s?.SleeveInstanceId ?? 0).Where(id => id > 0).OrderBy(id => id).ToList();
                    if (sleeveIds.Count == cluster.Count && _rotatedBboxCache.Value.Count < MAX_ROTATED_BBOX_CACHE_SIZE)
                    {
                        string cacheKey = $"RBB_{string.Join("_", sleeveIds)}_{rotationAngle:F6}";
                        _rotatedBboxCache.Value[cacheKey] = result;
                        
                        if (!DeploymentConfiguration.DeploymentMode && calcStopwatch.ElapsedMilliseconds > 10)
                        {
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ CACHE STORED: Rotated bounding box calculated in {calcStopwatch.ElapsedMilliseconds}ms for {cluster.Count} sleeves, rotation={rotationAngle * 180 / Math.PI:F1}°\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // If cache storage fails, continue (non-critical)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cache storage failed: {ex.Message}\n");
                }
            }
            
            return result;
        }

        /// <summary>
        /// Get stored rotation data for a cluster sleeve
        /// </summary>
        public (double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)? GetRotationData(int clusterInstanceId)
        {
            if (_clusterRotationData.Value.TryGetValue(clusterInstanceId, out var data))
            {
                return data;
            }
            return null;
        }

        /// <summary>
        /// Store rotation data for a cluster sleeve
        /// </summary>
        public void StoreRotationData(int clusterInstanceId, double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)
        {
            _clusterRotationData.Value[clusterInstanceId] = (rotationAngleDeg, isRotated, rotatedBboxMin, rotatedBboxMax, rotatedWidth, rotatedHeight, rotatedDepth);
        }

        /// <summary>
        /// Clear all stored rotation data
        /// </summary>
        public void ClearRotationData()
        {
            _clusterRotationData.Value.Clear();
        }

        /// <summary>
        /// ✅ OPTIMIZATION: Pre-load ClashZones into cache before parallel processing
        /// This eliminates database lookups during parallel execution (major performance improvement)
        /// </summary>
        public int PreloadClashZones(IEnumerable<int> sleeveIds, string? xmlFilePath = null)
        {
            if (sleeveIds == null)
                return 0;

            int preloadedCount = 0;
            var uniqueIds = sleeveIds.Where(id => id > 0).Distinct().ToList();
            
            // ✅ BATCH PRE-LOAD: Load all ClashZones in a single pass (before parallel processing)
            // This populates the cache so GetCachedClashZone hits cache instead of doing database lookups
            foreach (int sleeveId in uniqueIds)
            {
                // Skip if already in cache
                if (_clashZoneCache.Value.ContainsKey(sleeveId))
                    continue;

                // Load and cache
                try
                {
                    var cz = _getClashZoneFunc(sleeveId, xmlFilePath);
                    var clashZone = cz as ClashZone;
                    
                    if (clashZone != null && _clashZoneCache.Value.Count < MAX_CLASHZONE_CACHE_SIZE)
                    {
                        _clashZoneCache.Value[sleeveId] = clashZone;
                        preloadedCount++;
                    }
                }
                catch
                {
                    // Skip failed lookups
                }
            }

            return preloadedCount;
        }

        /// <summary>
        /// ✅ OPTIMIZATION: Pre-load ClashZones from dictionary (FAST - no database queries)
        /// Use this when ClashZones are already loaded in memory (e.g., from batch database query)
        /// </summary>
        public int PreloadClashZonesFromDictionary(Dictionary<int, ClashZone> clashZones)
        {
            if (clashZones == null || clashZones.Count == 0)
                return 0;

            int preloadedCount = 0;
            
            // ✅ FAST PATH: Directly populate cache from dictionary (no database queries)
            foreach (var kvp in clashZones)
            {
                int sleeveId = kvp.Key;
                ClashZone clashZone = kvp.Value;
                
                if (sleeveId > 0 && clashZone != null && _clashZoneCache.Value.Count < MAX_CLASHZONE_CACHE_SIZE)
                {
                    // Skip if already in cache
                    if (!_clashZoneCache.Value.ContainsKey(sleeveId))
                    {
                        _clashZoneCache.Value[sleeveId] = clashZone;
                        preloadedCount++;
                    }
                }
            }

            return preloadedCount;
        }
        
        /// <summary>
        /// ✅ PERFORMANCE: Helper method to get ClashZone with caching to avoid repeated database queries
        /// </summary>
        private ClashZone GetCachedClashZone(int sleeveInstanceId, string xmlFilePath)
        {
            if (sleeveInstanceId <= 0)
                return null;
            
            // Check cache first
            if (_clashZoneCache.Value.TryGetValue(sleeveInstanceId, out var cached))
            {
                return cached;
            }
            
            // Cache miss - load from function
            var cz = _getClashZoneFunc(sleeveInstanceId, xmlFilePath);
            var clashZone = cz as ClashZone;
            
            // Store in cache (if not null and cache not full)
            if (clashZone != null && _clashZoneCache.Value.Count < MAX_CLASHZONE_CACHE_SIZE)
            {
                _clashZoneCache.Value[sleeveInstanceId] = clashZone;
            }
            
            return clashZone;
        }
        
        /// <summary>
        /// ✅ PERFORMANCE: Helper method to store rotated bounding box result in cache
        /// </summary>
        private void StoreInCache(List<dynamic> cluster, double rotationAngle, 
            (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ) result,
            System.Diagnostics.Stopwatch calcStopwatch)
        {
            try
            {
                if (cluster != null && cluster.Count > 0 && _rotatedBboxCache.Value.Count < MAX_ROTATED_BBOX_CACHE_SIZE)
                {
                    var sleeveIds = cluster.Select(s => s?.SleeveInstanceId ?? 0).Where(id => id > 0).OrderBy(id => id).ToList();
                    if (sleeveIds.Count == cluster.Count)
                    {
                        string cacheKey = $"RBB_{string.Join("_", sleeveIds)}_{rotationAngle:F6}";
                        _rotatedBboxCache.Value[cacheKey] = result;
                        
                        if (!DeploymentConfiguration.DeploymentMode && calcStopwatch.ElapsedMilliseconds > 10)
                        {
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ CACHE STORED: Rotated bounding box calculated in {calcStopwatch.ElapsedMilliseconds}ms for {cluster.Count} sleeves, rotation={rotationAngle * 180 / Math.PI:F1}°\n");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // If cache storage fails, continue (non-critical)
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Cache storage failed: {ex.Message}\n");
                }
            }
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Calculate cluster placement point from intersection points (centroid)
        /// Individual sleeves are placed at intersection points, so cluster should be at average of intersection points
        /// This ensures the cluster sleeve is placed correctly relative to the individual intersection points
        /// </summary>


        private XYZ CalculatePlacementPointFromIntersections(List<dynamic> cluster, string? xmlFilePath)
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

                    var cz = GetCachedClashZone(sleeveInstanceId, xmlFilePath);
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
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Error getting intersection point for sleeve in cluster: {ex.Message}\n");
                }
            }

            if (validCount == 0)
            {
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ No valid intersection points found in cluster ({cluster.Count} sleeves)\n");
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
