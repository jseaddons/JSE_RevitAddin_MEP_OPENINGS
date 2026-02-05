using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections.Concurrent;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Geometry;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement; // ✅ SHARED: For SleeveRotationService reuse

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Rotation
{
    /// <summary>
    /// Service for handling cluster rotation calculations and rotated bounding box determinations
    /// Extracted from UniversalClusterService for better separation of concerns
    /// </summary>
    public class ClusterRotationService : IClusterRotationService
    {
        // Store rotation angle and rotated bounding box for each cluster sleeve
        // Key: ClusterInstanceId
        private readonly Dictionary<int, (double rotationAngleDeg, bool isRotated, XYZ rotatedBboxMin, XYZ rotatedBboxMax, double rotatedWidth, double rotatedHeight, double rotatedDepth)> _clusterRotationData;
        
        // ✅ THREAD-SAFE DEDUPLICATION: Track processed IDs to prevent concurrent threads from claiming the same ID
        private readonly ConcurrentDictionary<int, byte> _processedClusterIds = new ConcurrentDictionary<int, byte>();

        // ✅ PERFORMANCE: Cache rotated bounding box calculations by cluster signature + rotation angle
        // Key: Hash of (sorted sleeve IDs + rotation angle), Value: (width, height, depth, mid, rotatedMinX, ...)
        private readonly Dictionary<string, (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)> _rotatedBboxCache;
        private const int MAX_ROTATED_BBOX_CACHE_SIZE = 1000; // Limit cache size to prevent memory growth

        // ✅ PERFORMANCE: Cache individual ClashZone lookups to avoid repeated database queries for same sleeve
        // Key: sleeveInstanceId, Value: ClashZone (cached to avoid repeated _getClashZoneFunc calls)
        private readonly Dictionary<int, ClashZone> _clashZoneCache;
        private const int MAX_CLASHZONE_CACHE_SIZE = 5000; // Limit cache size to prevent memory growth

        // Delegate for getting ClashZone by sleeve instance ID (injected dependency)
        private readonly Func<int, string, ClashZone> _getClashZoneFunc;
        
        // Delegate for getting cluster placement by cluster ID (injected dependency)
        private readonly Func<int, XYZ> _getClusterPlacementFunc;

        // ✅ Same rotation logic as individual sleeves (Wall X/Y, Floor circular/rectangular, Framing)
        private readonly SleeveRotationService _sleeveRotationService;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="getClashZoneFunc">Function to retrieve ClashZone by sleeve instance ID</param>
        /// <param name="getClusterPlacementFunc">Optional function to retrieve existing cluster placement from DB</param>
        public ClusterRotationService(
            Func<int, string, ClashZone> getClashZoneFunc,
            Func<int, XYZ> getClusterPlacementFunc = null)
        {
            // ✅ VALIDATION: Check for null before assignment (better error message)
            if (getClashZoneFunc == null)
            {
                throw new ArgumentNullException(nameof(getClashZoneFunc), 
                    "getClashZoneFunc cannot be null. Provide a function that returns ClashZone (can return null if not found).");
            }
            
            _clusterRotationData = new Dictionary<int, (double, bool, XYZ, XYZ, double, double, double)>();
            _rotatedBboxCache = new Dictionary<string, (double, double, double, XYZ, double?, double?, double?, double?, double?, double?)>();
            _clashZoneCache = new Dictionary<int, ClashZone>();
            _getClashZoneFunc = getClashZoneFunc;
            _getClusterPlacementFunc = getClusterPlacementFunc;
            _sleeveRotationService = new SleeveRotationService();
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
            // ✅ PHASE 4 FIX: Get orientation from DATABASE, not from Revit API
            // This replaces all legacy heuristic logic (finding first zone, checking wall types manually, etc.)
            
            if (cluster != null && cluster.Count > 0)
            {
                var firstItem = cluster[0];
                ClashZone? firstClashZone = null;
                
                if (firstItem is ClashZone cz)
                    firstClashZone = cz;
                else if (firstItem?.ClashZone != null)
                    firstClashZone = firstItem.ClashZone as ClashZone;
                
                if (firstClashZone != null)
                {
                    // ✅ CRITICAL FIX: Circular elements (Pipes and Round Ducts) should NOT rotate in clusters
                    // This applies universally to all hosts (Walls, Floors, etc.)
                    // User Request: "for circular pipes clustering should not rotate"
                    bool allCircular = true;
                    foreach (var item in cluster)
                    {
                        ClashZone? itemCz = null;
                        if (item is ClashZone czItem) itemCz = czItem;
                        else if (item?.ClashZone != null) itemCz = item.ClashZone as ClashZone;

                        if (itemCz == null) { allCircular = false; break; }

                        // ✅ DETECTION BY MEP ELEMENT: Circular elements (Pipes/Round Ducts) never rotate
                        // User Req: "for circular element it should not rotate always even if it is rectangula sleeve"
                        bool isPipe = string.Equals(itemCz.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase);
                        bool isRoundDuct = (string.Equals(itemCz.MepElementCategory, "Ducts", StringComparison.OrdinalIgnoreCase) || 
                                           string.Equals(itemCz.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase)) &&
                                          (string.Equals(itemCz.DuctShape, "Round", StringComparison.OrdinalIgnoreCase) ||
                                           string.Equals(itemCz.DuctShape, "Circular", StringComparison.OrdinalIgnoreCase));

                        if (!isPipe && !isRoundDuct)
                        {
                            allCircular = false;
                            break;
                        }
                    }

                    if (cluster.Count > 0 && allCircular)
                    {
                        // For floors (and non wall/framing hosts) we keep the old rule:
                        // all-circular clusters get 0° rotation.
                        // For walls / structural framing we MUST still respect HostOrientation
                        // (X → 90°, Y → 0°) so that cluster orientation matches individual sleeves.
                        string hostType = firstClashZone.StructuralElementType ?? "";
                        bool isWallOrFraming =
                            hostType.StartsWith("Wall", StringComparison.OrdinalIgnoreCase) ||
                            hostType.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase);

                        if (!isWallOrFraming)
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ✅ ALL-CIRCULAR NON-WALL CLUSTER: Skipping rotation (0.0°)\n");
                            }
                            return 0.0;
                        }
                        // Wall / framing + all circular: fall through and use HostOrientation logic below
                    }

                    // Normalize HostOrientation string
                    string hostOrientation = (firstClashZone.HostOrientation ?? "").Trim();
                    string hostTypeForOrientation = firstClashZone.StructuralElementType ?? "";
                    bool isWallOrFramingForOrientation =
                        hostTypeForOrientation.StartsWith("Wall", StringComparison.OrdinalIgnoreCase) ||
                        hostTypeForOrientation.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase);
                    
                    // ✅ VALIDATION: For WALL/FRAMING only, all zones must have SAME HostOrientation (X vs Y)
                    // For FLOOR we do NOT require HostOrientation to match (often "" or "Floor") so rotated MEP gets MepElementRotationAngle
                    bool allSameOrientation = true;
                    if (isWallOrFramingForOrientation)
                    {
                        foreach (var item in cluster)
                        {
                            ClashZone? itemCz = null;
                            if (item is ClashZone clashZone)
                                itemCz = clashZone;
                            else if (item?.ClashZone != null)
                                itemCz = item.ClashZone as ClashZone;
                            
                            string itemOrientation = (itemCz?.HostOrientation ?? "").Trim();
                            if (!string.Equals(itemOrientation, hostOrientation, StringComparison.OrdinalIgnoreCase))
                            {
                                allSameOrientation = false;
                                SafeFileLogger.SafeAppendText("cluster_errors.log",
                                    $"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL: Cluster has MIXED orientations! " +
                                    $"Zone1 Orientation='{hostOrientation}', Zone2 Orientation='{itemOrientation}'\n");
                                break;
                            }
                        }
                        if (!allSameOrientation)
                        {
                            SafeFileLogger.SafeAppendText("cluster_errors.log",
                                $"[{DateTime.Now:HH:mm:ss}] ❌ SKIPPING cluster due to mixed orientations (Wall/Framing)\n");
                            return 0.0;
                        }
                    }
                    
                    // ✅ PHASE 4: Use database orientation directly (Simple Logic as requested)
                    // If HostOrientation is X/X-WALL -> Rotate 90 degrees
                    // Else -> 0 degrees
                    // ✅ USER RULE (2026-02-05): For Floors, NO host orientation needed. Strictly use MEP orientation for rotated elements.
                    // To avoid affecting walls/framing, we ONLY bypass this if host is explicitly a Floor.
                    bool isFloorForOrientation = (firstClashZone.StructuralElementType ?? "").IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (!isFloorForOrientation && (string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase) || 
                        hostOrientation.IndexOf("X-WALL", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        double rotationAngle = Math.PI / 2.0; // 90 degrees for X-walls
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ ORIENTATION (WALL): X-wall → 90° rotation\n");
                        }
                        return rotationAngle;
                    }
                    else if (!isFloorForOrientation && (string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase) || 
                             hostOrientation.IndexOf("Y-WALL", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                         double rotationAngle = 0.0; // 0 degrees for Y-walls (User Req: "y should remain at 0 degre")
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ ORIENTATION (WALL): Y-wall → 0° rotation\n");
                        }
                        return rotationAngle;
                    }
                    else
                    {
                        // For floors (or anything not explicitly X/Y wall), use MEP element rotation angle (from database)
                        // ✅ FIX (2026-02-05): "yes 0 degree ok ... because we already get extreme corners to shape the box no need to rotate 0 degree"
                        // Rule: If all angles are orthogonal (0, 90, 180, 270), return 0.
                        // Only return non-zero if we find a "truly" rotated angle (e.g. 45 degrees).
                        
                        double rotationAngle = 0.0;
                        bool foundNonOrthogonal = false;

                        foreach (var item in cluster)
                        {
                            ClashZone? itemCz = null;
                            if (item is ClashZone czItem) itemCz = czItem;
                            else if (item?.ClashZone != null) itemCz = item.ClashZone as ClashZone;
                            
                            if (itemCz != null)
                            {
                                double angle = itemCz.MepElementRotationAngle;
                                // Normalize to 0-2PI for comparison
                                while (angle < 0) angle += 2 * Math.PI;
                                while (angle >= 2 * Math.PI) angle -= 2 * Math.PI;

                                // Check if it's NOT a multiple of 90 degrees (1.570796 rad)
                                double remainder = Math.Abs(angle % (Math.PI / 2.0));
                                if (remainder > 1e-4 && Math.Abs(remainder - (Math.PI / 2.0)) > 1e-4)
                                {
                                    rotationAngle = angle;
                                    foundNonOrthogonal = true;
                                    break;
                                }
                            }
                        }

                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            string typeLog = string.IsNullOrEmpty(hostOrientation) ? "Unknown/Floor" : hostOrientation;
                            string pathLog = foundNonOrthogonal ? "NON-ORTHOGONAL" : "ORTHOGONAL-MIX (Default to 0)";
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ ORIENTATION (OTHER): {typeLog} → {rotationAngle * 180 / Math.PI:F1}° rotation ({pathLog})\n");
                        }
                        return rotationAngle;
                    }
                }
            }

            // ❌ If we reach here, database values were not available - ERROR
            SafeFileLogger.SafeAppendText("cluster_errors.log",
                $"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL: Could not get HostOrientation from database for cluster\n");
            return 0.0; // Fallback to no rotation (error case)
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
                        string cacheKey = $"RBB_FIXED_{string.Join("_", sleeveIds)}_{rotationAngle:F6}";
                        
                        if (_rotatedBboxCache.TryGetValue(cacheKey, out var cachedResult))
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
                                    $"[{DateTime.Now:HH:mm:ss}] 💾 CACHE MISS: Rotated bounding box for {cluster.Count} sleeves, rotation={rotationAngle * 180 / Math.PI:F1}°, cacheKey={cacheKey.Substring(0, Math.Min(50, cacheKey.Length))}, cacheSize={_rotatedBboxCache.Count}\n");
                                DebugLogger.Info($"[BBOX-CACHE] 💾 CACHE MISS: Rotated bounding box for {cluster.Count} sleeves, rotation={rotationAngle * 180 / Math.PI:F1}°, cacheSize={_rotatedBboxCache.Count}");
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
            XYZ placementPoint = null;
            
            // ✅ BATCH MODE FIX: Try to get existing placement from DB first (Single Source of Truth)
            // If the cluster already exists in DB (has ClusterInstanceId), we MUST use the stored placement
            // This bypasses recalculation errors where incorrect grouping in batch mode shifts the centroid
            if (_getClusterPlacementFunc != null && cluster[0] != null)
            {
                try
                {
                    // Check if first sleeve has a valid ClusterInstanceId
                    int clusterId = -1;
                    dynamic firstSleeve = cluster[0];
                    object clusterIdObj = null;
                    
                    // Try to get ClusterInstanceId from dynamic object (direct property)
                    try { clusterIdObj = firstSleeve.ClusterInstanceId; } catch { }
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] DEBUG: Sleeve Type={firstSleeve?.GetType().Name}, Direct ID access result={clusterIdObj ?? "null"}\n");

                    // ✅ CRITICAL FIX: Try to get from ClashZone property on dynamic object
                    if (clusterIdObj == null || (clusterIdObj is int cId1 && cId1 <= 0))
                    {
                        try 
                        {
                            var cz = firstSleeve.ClashZone;
                            if (cz != null) 
                            {
                                clusterIdObj = cz.ClusterInstanceId;
                                if (!DeploymentConfiguration.DeploymentMode)
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] DEBUG: Got ID {clusterIdObj} from ClashZone property\n");
                            }
                            else
                            {
                                 if (!DeploymentConfiguration.DeploymentMode)
                                    SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] DEBUG: firstSleeve.ClashZone IS NULL\n");
                            }
                        }
                        catch (Exception ex) 
                        { 
                             if (!DeploymentConfiguration.DeploymentMode)
                                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] DEBUG: Error accessing ClashZone property: {ex.Message}\n");
                        }
                    }

                    // Fallback: try to get from cached/DB ClashZone lookup
                    if (clusterIdObj == null || (clusterIdObj is int cId2 && cId2 <= 0))
                    {
                         var cz = GetCachedClashZone(firstSleeve.SleeveInstanceId, xmlFilePath);
                         if (cz != null) 
                         {
                            clusterId = cz.ClusterInstanceId;
                            if (!DeploymentConfiguration.DeploymentMode)
                                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] DEBUG: Got ID {clusterId} from GetCachedClashZone\n");
                         }
                         else
                         {
                            if (!DeploymentConfiguration.DeploymentMode)
                                SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] DEBUG: GetCachedClashZone returned NULL for {firstSleeve.SleeveInstanceId}\n");
                         }
                    }
                    else
                    {
                        clusterId = (int)clusterIdObj;
                    }
                    
                    if (clusterId > 0)
                    {
                        // ✅ THREAD-SAFE DEDUPLICATION: Try to claim this ID for this run.
                        // If TryAdd returns FALSE, it means another thread or previous call has already claimed this ID.
                        // We must treat this current cluster as a NEW cluster (ignore DB ID) to prevent stacking.
                        if (!_processedClusterIds.TryAdd(clusterId, 0))
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_debug.log", 
                                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ DEDUPLICATION (Thread-Safe): ClusterInstanceId {clusterId} already processed! " +
                                    $"Treating this duplicate group as a NEW cluster (ignoring DB ID) to prevent stacking.\n");
                            }
                            clusterId = -1; // Force new calculation (geometry based)
                        }

                         // Verify delegate
                         if (_getClusterPlacementFunc == null)
                             SafeFileLogger.SafeAppendText("cluster_debug.log", $"[{DateTime.Now:HH:mm:ss}] DEBUG: _getClusterPlacementFunc IS NULL!\n");
                    }
                    
                    if (clusterId > 0)
                    {
                        var dbPlacement = _getClusterPlacementFunc(clusterId);
                        if (dbPlacement != null)
                        {
                            placementPoint = dbPlacement;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                    $"[{DateTime.Now:HH:mm:ss}] 🎯 DB SOURCE OF TRUTH: Using stored placement for Cluster {clusterId}: ({placementPoint.X:F4}, {placementPoint.Y:F4}, {placementPoint.Z:F4})\n");
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
            
            // Fallback to calculation if DB lookup failed or returned null
            if (placementPoint == null)
            {
                 placementPoint = CalculatePlacementPointFromIntersections(cluster, xmlFilePath);
            }
            
            if (placementPoint.IsZeroLength() || double.IsNaN(placementPoint.X))
            {
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ Failed to calculate placement point from intersections, falling back to first sleeve intersection point\n");
                // Fallback: Use first sleeve's intersection point
                var fallbackCz = GetCachedClashZone(cluster[0].SleeveInstanceId, xmlFilePath);
                if (fallbackCz != null)
                {
                    placementPoint = new XYZ(fallbackCz.IntersectionPointX, fallbackCz.IntersectionPointY, fallbackCz.IntersectionPointZ);
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

                // ❌ OLD ROUND PIPE LOGIC REMOVED - Logic moved inside strictly separated Wall/Floor blocks below
                // This ensures we don't mix Wall vs Floor logic for round pipes
                // Previous logic was Lines 461-786 - now deleted/deferred
                bool deferRoundPipeCalculation = isRoundPipeCluster; 
                
                if (deferRoundPipeCalculation && !DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                         $"[{DateTime.Now:HH:mm:ss}] ➡️ DEFERRING Round Pipe Logic: Will be handled inside specific Host Type blocks (Wall vs Floor) to ensure correct orientation.\n");
                }


                // (Block Removed)
                // Logic already handled in previous code that detects HostType/WallDirection


                // (Block Removed)

                // (Block Removed)

                // (Block Removed) cache logic moved later

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
                        var hostCz = GetCachedClashZone(firstSleeveInstanceId, xmlFilePath);
                        if (hostCz != null)
                        {
                            bool isWallHost = string.Equals(hostCz.StructuralElementType, "Wall", StringComparison.OrdinalIgnoreCase) ||
                                              string.Equals(hostCz.StructuralElementType, "Walls", StringComparison.OrdinalIgnoreCase);
                            bool isFramingHost = string.Equals(hostCz.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
                            
                            if ((isWallHost || isFramingHost) && hostCz.WallDirection != null && !hostCz.WallDirection.IsZeroLength())
                            {
                                isWallOrFraming = true;
                                wallDirection = hostCz.WallDirection;
                                wallOrigin = new XYZ(hostCz.SleevePlacementPointX, 
                                                    hostCz.SleevePlacementPointY, 
                                                    hostCz.SleevePlacementPointZ);
                                
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
                
                // ✅ FIX: For batch clustering, individual sleeves may be deleted (SleeveInstanceId = -1)
                // but corners are still available in the ClashZone database record.
                // Use the ClashZone from the cluster data directly instead of looking it up by SleeveInstanceId.
                int sleeveInstanceId = 0;
                ClashZone clashZone = null;
                
                try
                {
                    sleeveInstanceId = sleeveData.SleeveInstanceId;
                    
                    // Try to get ClashZone from the sleeveData object itself (it should have a ClashZone property)
                    if (sleeveData is ClashZone)
                    {
                        clashZone = sleeveData as ClashZone;
                    }
                    else
                    {
                        // Try to access ClashZone property dynamically
                        try
                        {
                            clashZone = sleeveData.ClashZone;
                        }
                        catch
                        {
                            // If that fails and we have a valid SleeveInstanceId, try loading from cache
                            if (sleeveInstanceId > 0)
                            {
                                clashZone = GetCachedClashZone(sleeveInstanceId, xmlFilePath);
                            }
                        }
                    }
                    
                    if (clashZone == null)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Sleeve {sleeveInstanceId}: No ClashZone data available\n");
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ❌ Error getting ClashZone: {ex.Message}\n");
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
                
                // ✅ CRITICAL FIX: Store ClashZone objects directly instead of IDs
                // For batch clustering, individual sleeves may be deleted (SleeveInstanceId = -1)
                // but ClashZone objects with corners are still available in the cluster data
                var clashZonesInCluster = new List<ClashZone>();
                foreach (var sleeveData in cluster)
                {
                    try
                    {
                        ClashZone clashZone = null;
                        
                        // Try to get ClashZone from sleeveData (same logic as earlier in the method)
                        if (sleeveData is ClashZone)
                        {
                            clashZone = sleeveData as ClashZone;
                        }
                        else
                        {
                            try
                            {
                                clashZone = sleeveData.ClashZone;
                            }
                            catch
                            {
                                // If that fails, try to get SleeveInstanceId and load from cache
                                try
                                {
                                    int sleeveId = sleeveData.SleeveInstanceId;
                                    if (sleeveId > 0)
                                    {
                                        clashZone = GetCachedClashZone(sleeveId, xmlFilePath);
                                    }
                                }
                                catch { }
                            }
                        }
                        
                        if (clashZone != null)
                        {
                            clashZonesInCluster.Add(clashZone);
                        }
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
                    bool hasAnyCorners = false;
                    int sleevesWithCorners = 0;
                    
                    // Now check corners using the ClashZone objects directly (no lookup needed)
                    foreach (var clashZone in clashZonesInCluster)
                    {
                        if (clashZone != null)
                        {
                            int sleeveId = clashZone.SleeveInstanceId > 0 ? clashZone.SleeveInstanceId : -1;
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
                                        $"PlacementPointZ={clashZone.SleevePlacementPointZ:F6}, BBoxMinZ={clashZone.SleeveBoundingBoxMinZ:F6}\n");

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
                        
                        foreach (var clashZone in clashZonesInCluster)
                        {
                            if (clashZone != null)
                            {
                                int sleeveId = clashZone.SleeveInstanceId > 0 ? clashZone.SleeveInstanceId : -1;
                                
                                // ✅ CRITICAL FIX: Explicitly type nullable properties to avoid dynamic dispatch errors
                                // ✅ READ FROM DATABASE: These corners were batch saved after regeneration
                                // - No recalculation needed - corners are already in database
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
                                
                                if (corner1X.HasValue && corner1Y.HasValue &&
                                    corner2X.HasValue && corner2Y.HasValue &&
                                    corner3X.HasValue && corner3Y.HasValue &&
                                    corner4X.HasValue && corner4Y.HasValue)
                                {
                                    // ✅ CRITICAL FIX: Use stored corner Z coordinates if available, otherwise fall back to bounding box MinZ
                                    // The stored corner Z coordinates are the authoritative source - they were calculated from the actual sleeve placement
                                    double corner1ZValue = corner1Z.HasValue ? corner1Z.Value : clashZone.SleeveBoundingBoxMinZ;
                                    double corner2ZValue = corner2Z.HasValue ? corner2Z.Value : clashZone.SleeveBoundingBoxMinZ;
                                    double corner3ZValue = corner3Z.HasValue ? corner3Z.Value : clashZone.SleeveBoundingBoxMinZ;
                                    double corner4ZValue = corner4Z.HasValue ? corner4Z.Value : clashZone.SleeveBoundingBoxMinZ;
                                    
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
                                
                                // ✅ 1. ROUND PIPES/DUCTS IN WALLS
                                // Strictly use Edge-to-Edge logic along wall length
                                if (deferRoundPipeCalculation)
                                {
                                     // Get authoritative thickness first
                                     double maxThickness = clashZonesInCluster.Any() 
                                         ? clashZonesInCluster.Max(c => c.StructuralElementThickness) 
                                         : 0.2; // Default 200mm
                                     
                                     // Determine Wall Axis
                                     bool isYWall = Math.Abs(wallDirection.Y) > Math.Abs(wallDirection.X);
                                     
                                     // Gather extents
                                     double mnX = allCorners.Min(c => c.X);
                                     double mxX = allCorners.Max(c => c.X);
                                     double mnY = allCorners.Min(c => c.Y);
                                     double mxY = allCorners.Max(c => c.Y);
                                     
                                     if (isYWall)
                                     {
                                         // Wall along Y -> Width is Y-range
                                         cornerWidth = mxY - mnY;
                                         rotatedMinX = mnX; rotatedMaxX = mxX;
                                         rotatedMinY = mnY; rotatedMaxY = mxY;
                                     }
                                     else
                                     {
                                         // Wall along X -> Width is X-range
                                         cornerWidth = mxX - mnX;
                                         rotatedMinX = mnX; rotatedMaxX = mxX;
                                         rotatedMinY = mnY; rotatedMaxY = mxY;
                                     }
                                     
                                     // Depth = Thickness (Authoritative)
                                     cornerDepth = maxThickness;
                                     
                                     // Height = Z-range (or MaxDiameter logic if needed, but Z-range is safer for stacks)
                                     double mnZ = allCorners.Min(c => c.Z);
                                     double mxZ = allCorners.Max(c => c.Z);
                                     
                                     // If single pipe, height might just be diameter, but let's stick to extents for safety
                                     // unless it's way too small?
                                     // For now, Z-extent is standard.
                                     cornerHeight = mxZ - mnZ;
                                     
                                     // Placement Point: Geometric Center
                                     double midX = (mnX + mxX) / 2.0;
                                     double midY = (mnY + mxY) / 2.0;
                                     double midZ = (mnZ + mxZ) / 2.0;
                                     
                                     placementPoint = new XYZ(midX, midY, midZ);
                                     
                                     if (!DeploymentConfiguration.DeploymentMode)
                                         SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                             $"[{DateTime.Now:HH:mm:ss}] 🔵 ROUND WALL CLUSTER: W={cornerWidth*304.8:F1}, D={cornerDepth*304.8:F1}, Center=({midX:F2},{midY:F2},{midZ:F2})\n");
                                }
                                else
                                {
                                    // ✅ 2. RECTANGULAR ELEMENTS IN WALLS (Existing Logic)
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

                                     // ✅ OPTIONAL FIX: Update placement point to be the geometric center of the bounds
                                     if (OptimizationFlags.UseGeometricCenterForWalls)
                                     {
                                         double midX = (wcsMinX + wcsMaxX) / 2.0;
                                         double midY = (wcsMinY + wcsMaxY) / 2.0;
                                         
                                         // ✅ SEPARATED LOGIC: Calculate Z geometric center for Walls too
                                         double wcsMinZ = allCorners.Min(c => c.Z);
                                         double wcsMaxZ = allCorners.Max(c => c.Z);
                                         double midZ = (wcsMinZ + wcsMaxZ) / 2.0;
                                         
                                         placementPoint = new XYZ(midX, midY, midZ);

                                         SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                             $"[{DateTime.Now:HH:mm:ss}]   ✅ STRAIGHT WALL CENTER: [FLAG ENABLED] Shifted from Centroid to Geometric Center: ({placementPoint.X:F4}, {placementPoint.Y:F4}, {placementPoint.Z:F4})\n");
                                     }
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
                                    
                                    rotatedMinY = minRotY;
                                    rotatedMaxY = maxRotY;

                                     // ✅ CRITICAL FIX: Update placementPoint to be the GEOMETRIC CENTER for Rotated Walls
                                     // Gated by safety flag per user request.
                                     if (OptimizationFlags.UseGeometricCenterForWalls)
                                     {
                                         double midRotX = (minRotX + maxRotX) / 2.0;
                                         double midRotY = (minRotY + maxRotY) / 2.0;

                                         // Inverse Rotate (from Rotated Space -> Relative World Space)
                                         // Matrix Inverse (Transpose):
                                         // x = x' * cos(A) + y' * sin(A)
                                         // y = x' * (-sin(A)) + y' * cos(A)
                                         double relMidX = midRotX * cosCluster + midRotY * sinCluster;
                                         double relMidY = midRotX * (-sinCluster) + midRotY * cosCluster;

                                         // Add Origin (Relative World Space -> Absolute World Space)
                                         double worldMidX = originX + relMidX;
                                         double worldMidY = originY + relMidY;

                                         // ✅ SEPARATED LOGIC: Calculate Z geometric center for Walls too
                                         double worldMinZ = allCorners.Min(c => c.Z);
                                         double worldMaxZ = allCorners.Max(c => c.Z);
                                         double worldMidZ = (worldMinZ + worldMaxZ) / 2.0;

                                         // Update placementPoint
                                         placementPoint = new XYZ(worldMidX, worldMidY, worldMidZ);

                                         SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                             $"[{DateTime.Now:HH:mm:ss}]   ✅ ROTATED WALL CENTER: [FLAG ENABLED] Shifted from Centroid to Geometric Center: ({placementPoint.X:F4}, {placementPoint.Y:F4}, {placementPoint.Z:F4}). MidRot=({midRotX:F4},{midRotY:F4})\n");
                                     }
                                     else
                                     {
                                         SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                             $"[{DateTime.Now:HH:mm:ss}]   ⚠️ ROTATED WALL CENTER: [FLAG DISABLED] Staying at Centroid: ({placementPoint.X:F4}, {placementPoint.Y:F4})\n");
                                     }
                                    
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ ROTATED AXIS WALL: Transformed {transformedCorners.Count} corners to cluster's rotated coordinate system. " +
                                        $"Rotation={rotationAngle * 180 / Math.PI:F1}°, WallType={(isYWall ? "Y-wall" : "X-wall")}, " +
                                        $"minRotX={minRotX:F6}, maxRotX={maxRotX:F6}, minRotY={minRotY:F6}, maxRotY={maxRotY:F6}, " +
                                        $"Width={cornerWidth:F6} (from {(isYWall ? "X" : "Y")} range in rotated space)\n");
                                }
                                } // End of injected 'else' (Rectangular Wall)
                            }
                            
                            if (!isWallOrFraming)
                            {
                                // ✅ FLOOR/OTHER: Calculate in WCS with rotation if needed
                                // Calculate cluster bounding box by rotating corners back to aligned axis
                                
                                if (deferRoundPipeCalculation)
                                {
                                     // ✅ 1. ROUND PIPES/DUCTS IN FLOORS
                                     // Width/Height = World X/Y extents (aligned)
                                     // Depth = Structural Thickness
                                     
                                      double maxThickness = clashZonesInCluster.Any() 
                                         ? clashZonesInCluster.Max(c => c.StructuralElementThickness) 
                                         : 0.2; // Default 200mm
                                     
                                     double mnX = allCorners.Min(c => c.X);
                                     double mxX = allCorners.Max(c => c.X);
                                     double mnY = allCorners.Min(c => c.Y);
                                     double mxY = allCorners.Max(c => c.Y);
                                     double mnZ = allCorners.Min(c => c.Z);
                                     double mxZ = allCorners.Max(c => c.Z);
                                     
                                     // Check for rotation (e.g. if the "World Alignment" is actually rotated due to MEP Rotation)
                                     // But for "Round Floor", usually X/Y is sufficient unless strictly rotated?
                                     // If rotationAngle is significant, we should probably follow rotated logic.
                                     // Let's reuse the ROTATED logic from below but adapt it for Round.
                                     
                                     // Actually, simple bounding box is often enough for round floor clusters unless they are diagonal runs.
                                     // Let's trust the extents.
                                     
                                     cornerWidth = mxX - mnX;
                                     cornerHeight = mxY - mnY;
                                     cornerDepth = maxThickness;
                                     
                                     // Geometric Center
                                     placementPoint = new XYZ(
                                         (mnX + mxX) / 2.0,
                                         (mnY + mxY) / 2.0,
                                         (mnZ + mxZ) / 2.0
                                     );
                                     
                                     rotatedMinX = mnX; rotatedMaxX = mxX;
                                     rotatedMinY = mnY; rotatedMaxY = mxY;
                                     
                                      if (!DeploymentConfiguration.DeploymentMode)
                                         SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                             $"[{DateTime.Now:HH:mm:ss}] 🔵 ROUND FLOOR CLUSTER: W={cornerWidth*304.8:F1}, H={cornerHeight*304.8:F1}, Center=({placementPoint.X:F2},{placementPoint.Y:F2},{placementPoint.Z:F2})\n");
                                }
                                else
                                { 
                                    // ✅ 2. RECTANGULAR ELEMENTS IN FLOORS (Existing Logic)
                                
                                // ✅ CRITICAL: Check for mixed X/Y orientations BEFORE rotation transformation
                                // When MEP elements have different orientations (X and Y), corner-based calculation may not work correctly
                                // because corners are in different coordinate systems. Use envelope method instead.
                                bool hasMixedOrientations = false;
                                var orientations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                
                                 foreach (var cz in clashZonesInCluster)
                                 {
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
                                        cornerDepth = 0.0; // Strictly from structural thickness later
                                        
                                        // ✅ Calculate new center point from corner extents
                                        // X and Y: Midpoint of corner extents
                                        // Z: Get from first sleeve (not from corner extents)
                                        double midX = (floorMinX + floorMaxX) / 2.0;
                                        double midY = (floorMinY + floorMaxY) / 2.0;
                                        double mixedFloorZ = 0.0;
                                        
                                         // ✅ FIXED: Use Geometric Center Z (Midpoint of Z-extents) initially
                                         mixedFloorZ = (floorMinZ + floorMaxZ) / 2.0;

                                         // ✅ CRITICAL RE-VERIFICATION: Floor Cluster Z-Position Fix (Half-In/Half-Out)
                                         // User Requirement: "rotated cluster should be pitch z coordinates for floor same as its constiuents... just get any constiuent sleeve and use the z coordinate"
                                         // This overrides Geometric Center logic for Z-axis on floors
                                         if (clashZonesInCluster.Count > 0)
                                         {
                                             var mixedCz = clashZonesInCluster[0];
                                             if (mixedCz != null)
                                             {
                                                 double z = mixedCz.SleevePlacementPointZ;
                                                 if (z == 0.0) z = mixedCz.IntersectionPointZ;
                                                 if (z != 0.0) 
                                                 {
                                                     mixedFloorZ = z;
                                                      if (!DeploymentConfiguration.DeploymentMode)
                                                     {
                                                         SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                             $"[{DateTime.Now:HH:mm:ss}]   ✅ MIXED FLOOR Z-FIX: Using Constituent Z (Pitch Z). Z={mixedFloorZ:F6}\n");
                                                     }
                                                 }
                                             }
                                         }
                                        
                                        placementPoint = new XYZ(midX, midY, mixedFloorZ);
                                        
                                        // ✅ Z-AXIS DIAGNOSTIC: Log actual Z-coordinates to identify shift source
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            double theoreticalMidZ = (floorMinZ + floorMaxZ) / 2.0;
                                            SafeFileLogger.SafeAppendText("cluster_placement_debug.log", 
                                                $"[{DateTime.Now:HH:mm:ss}] 🎯 Z-AXIS DIAGNOSTIC (MIXED FLOOR): " +
                                                $"LowerCornerZ={floorMinZ:F6}, UpperCornerZ={floorMaxZ:F6}, " +
                                                $"TheoreticalMidZ={theoreticalMidZ:F6}, PlacedZ={mixedFloorZ:F6}, " +
                                                $"Shift={mixedFloorZ - theoreticalMidZ:F6}\n");
                                        }
                                        
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
                                                 $"Orientations=[{string.Join(", ", orientations)}], Sleeves={clashZonesInCluster.Count}, " +
                                                 $"GUIDs=[{string.Join(", ", clashZonesInCluster.Select(z => z.ClashZoneGuid))}], " +
                                                 $"Corners={allCorners.Count}, " +
                                                 $"Min=({floorMinX:F6}, {floorMinY:F6}, {floorMinZ:F6}), " +
                                                 $"Max=({floorMaxX:F6}, {floorMaxY:F6}, {floorMaxZ:F6}), " +
                                                 $"PlacementPoint=({placementPoint.X:F6}, {placementPoint.Y:F6}, {placementPoint.Z:F6})\n");
                                        }
                                        
                                        // ✅ Skip rotation transformation for mixed orientations (already have world-space bounds)
                                        // Set depth from StructuralElementThickness if not already set
                                        // Strictly use structural thickness for depth
                                         var firstCzForDepth = clashZonesInCluster.FirstOrDefault();
                                         cornerDepth = firstCzForDepth?.StructuralElementThickness ?? 0.0;
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
                                         
                                         SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                                             $"    🧮 MATH (STRAIGHT FLOOR): X=[{wcsMinX:F4} to {wcsMaxX:F4}] -> MidX={midX:F4}, Y=[{wcsMinY:F4} to {wcsMaxY:F4}] -> MidY={midY:F4}\n");
                                        
                                        // ✅ CRITICAL RE-VERIFICATION: Floor Cluster Z-Position Fix (Half-In/Half-Out)
                                        // User Requirement: "rotated cluster should be pitch z coordinates for floor same as its constiuents... just get any constiuent sleeve and use the z coordinate"
                                        // This overrides Geometric Center logic for Z-axis on floors
                                        
                                        double minAZ = allCorners.Min(c => c.Z);
                                        double maxAZ = allCorners.Max(c => c.Z);
                                        
                                        // Default to Geometric Center initially for safety
                                        double placementZ = (minAZ + maxAZ) / 2.0;

                                        // ✅ OVERRIDE with "Pitch Z" (Constituent Z) for Floors
                                        if (!isWallOrFraming && clashZonesInCluster.Count > 0)
                                        {
                                            var refCz = clashZonesInCluster[0];
                                            if (refCz != null)
                                            {
                                                double z = refCz.SleevePlacementPointZ;
                                                if (z == 0.0) z = refCz.IntersectionPointZ;
                                                if (z != 0.0) 
                                                {
                                                    placementZ = z;
                                                    if (!DeploymentConfiguration.DeploymentMode)
                                                    {
                                                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                            $"[{DateTime.Now:HH:mm:ss}]   ✅ FLOOR Z-FIX: Using Constituent Z (Pitch Z) instead of Geometric Center. Z={placementZ:F6} (from Sleeve {refCz.SleeveInstanceId})\n");
                                                    }
                                                }
                                            }
                                        }

                                        /* PREVIOUS LOGIC REMOVED
                                        // Z: From first sleeve in cluster (to maintain elevation? No, we want center)
                                        double placementZ = 0.0;
                                        if (clashZonesInCluster.Count > 0)
                                        {
                                             var straightCz = clashZonesInCluster[0];
                                             if (straightCz != null)
                                             {
                                                 int firstSleeveId = straightCz.SleeveInstanceId;
                                                 placementZ = straightCz.SleevePlacementPointZ;
                                                 
                                                 if (placementZ == 0.0) placementZ = straightCz.IntersectionPointZ;
                                                 if (placementZ == 0.0 && straightCz.SleeveCorner1Z.HasValue) placementZ = straightCz.SleeveCorner1Z.Value;

                                                 SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                     $"[{DateTime.Now:HH:mm:ss}]   ✅ STRAIGHT AXIS FLOOR: Got Z from first sleeve (ID={firstSleeveId}, GUID={straightCz.ClashZoneGuid}): " +
                                                     $"SleevePlacementPointZ={straightCz.SleevePlacementPointZ:F6}, " +
                                                     $"IntersectionPointZ={straightCz.IntersectionPointZ:F6}, " +
                                                     $"Final placementZ={placementZ:F6}\n");
                                             }
                                             else
                                             {
                                                 SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                     $"[{DateTime.Now:HH:mm:ss}]   ⚠️ STRAIGHT AXIS FLOOR: First sleeve ClashZone is NULL!\n");
                                             }
                                         }
                                        else
                                         {
                                             SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                 $"[{DateTime.Now:HH:mm:ss}]   ⚠️ STRAIGHT AXIS FLOOR: clashZonesInCluster is EMPTY! Count={clashZonesInCluster.Count}\n");
                                         }
                                        */
                                        
                                        // ✅ CRITICAL: Override placement point with X/Y midpoints and Z from first sleeve
                                        placementPoint = new XYZ(midX, midY, placementZ);
                                        
                                        // ✅ Z-AXIS DIAGNOSTIC: Log actual Z-coordinates to identify shift source
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            double lowerZ = allCorners.Min(c => c.Z);
                                            double upperZ = allCorners.Max(c => c.Z);
                                            double theoreticalMidZ = (lowerZ + upperZ) / 2.0;
                                            SafeFileLogger.SafeAppendText("cluster_placement_debug.log", 
                                                $"[{DateTime.Now:HH:mm:ss}] 🎯 Z-AXIS DIAGNOSTIC (STRAIGHT FLOOR): " +
                                                $"LowerCornerZ={lowerZ:F6}, UpperCornerZ={upperZ:F6}, " +
                                                $"TheoreticalMidZ={theoreticalMidZ:F6}, PlacedZ={placementZ:F6}, " +
                                                $"Shift={placementZ - theoreticalMidZ:F6} (Should be non-zero if Pitch Z used)\n");
                                        }
                                        
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
                                         
                                         // ✅ Use corner centroid as rotation origin OR geometric center via flag
                                         if (OptimizationFlags.UseGeometricCenterForFloors)
                                         {
                                             // 1. First pass: Use centroid as temporary origin to find bounds in rotated space
                                             double tempOriginX = allCorners.Average(c => c.X);
                                             double tempOriginY = allCorners.Average(c => c.Y);
                                             
                                             double gMinX = double.MaxValue, gMaxX = double.MinValue;
                                             double gMinY = double.MaxValue, gMaxY = double.MinValue;
                                             
                                             foreach (var c in allCorners)
                                             {
                                                 double rx = (c.X - tempOriginX) * cosA - (c.Y - tempOriginY) * sinA;
                                                 double ry = (c.X - tempOriginX) * sinA + (c.Y - tempOriginY) * cosA;
                                                 gMinX = Math.Min(gMinX, rx); gMaxX = Math.Max(gMaxX, rx);
                                                 gMinY = Math.Min(gMinY, ry); gMaxY = Math.Max(gMaxY, ry);
                                             }
                                             
                                             // 2. Calculate midpoint in rotated space and transform back to world space
                                             double midRotX = (gMinX + gMaxX) / 2.0;
                                             double midRotY = (gMinY + gMaxY) / 2.0;
                                             
                                             // Inverse transform: [cos sin; -sin cos]
                                             double relX = midRotX * cosA + midRotY * sinA;
                                             double relY = -midRotX * sinA + midRotY * cosA;
                                             
                                             originX = tempOriginX + relX;
                                             originY = tempOriginY + relY;
                                             
                                             SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                                                 $"    🧮 MATH (ROTATED FLOOR): [FLAG ENABLED] Geometric Center Origin = ({originX:F4}, {originY:F4})\n");
                                         }
                                         else
                                         {
                                             // LEGACY: Use corner centroid as rotation origin
                                             originX = allCorners.Average(c => c.X);
                                             originY = allCorners.Average(c => c.Y);
                                             
                                             SafeFileLogger.SafeAppendText("cluster_sizing.log", 
                                                 $"    🧮 MATH (ROTATED FLOOR): [FLAG DISABLED] Legacy Origin (Average) = ({originX:F4}, {originY:F4})\n");
                                         }
                                        
                                        // ✅ CRITICAL FIX: Update placement point to match rotation origin (centroid)
                                        // This ensures the cluster is placed exactly at the centroid of the corners used for sizing, preventing misalignment
                                        double placementZ = placementPoint.Z;
                                        
                                         // Consistency with Straight Axis logic: fetch Z from first sleeve to ensure alignment with host plane
                                         if (clashZonesInCluster.Count > 0)
                                         {
                                             var rotatedCz = clashZonesInCluster[0];
                                             if (rotatedCz != null)
                                             {
                                                 double z = rotatedCz.SleevePlacementPointZ;
                                                 if (z == 0.0) z = rotatedCz.IntersectionPointZ;
                                             if (z != 0.0) placementZ = z;
                                         }
                                     }
                                    
                                    // ✅ CRITICAL RE-VERIFICATION: Floor Cluster Z-Position Fix (Half-In/Half-Out)
                                    // Override Geometric Center if it was set
                                    if (!isWallOrFraming && clashZonesInCluster.Count > 0)
                                    {
                                        var refCz = clashZonesInCluster[0];
                                        if (refCz != null)
                                        {
                                            double z = refCz.SleevePlacementPointZ;
                                            if (z == 0.0) z = refCz.IntersectionPointZ;
                                            if (z != 0.0) 
                                            {
                                                placementZ = z;
                                                if (!DeploymentConfiguration.DeploymentMode)
                                                {
                                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ ROTATED FLOOR Z-FIX: Using Constituent Z (Pitch Z). Z={placementZ:F6}\n");
                                                }
                                            }
                                        }
                                    }

                                    placementPoint = new XYZ(originX, originY, placementZ);
                                        
                                        // ✅ Z-AXIS DIAGNOSTIC: Log actual Z-coordinates to identify shift source
                                        if (!DeploymentConfiguration.DeploymentMode)
                                        {
                                            // For rotated floors, we need to find Min/Max Z from allCorners
                                            double lowerZ = allCorners.Min(c => c.Z);
                                            double upperZ = allCorners.Max(c => c.Z);
                                            double theoreticalMidZ = (lowerZ + upperZ) / 2.0;
                                            SafeFileLogger.SafeAppendText("cluster_placement_debug.log", 
                                                $"[{DateTime.Now:HH:mm:ss}] 🎯 Z-AXIS DIAGNOSTIC (ROTATED FLOOR): " +
                                                $"LowerCornerZ={lowerZ:F6}, UpperCornerZ={upperZ:F6}, " +
                                                $"TheoreticalMidZ={theoreticalMidZ:F6}, PlacedZ={placementZ:F6}, " +
                                                $"Shift={placementZ - theoreticalMidZ:F6}\n");
                                        }
                                        
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
                                } // End of if (!hasMixedOrientations)
                                } // End of injected 'else' (Rectangular Floor)
                            } // End of if (!isWallOrFraming)
                            
                            // ✅ HEIGHT: Calculate from Z range of bounding boxes (vertical dimension)
                            // For walls/framing: Height = Z range (vertical), NOT from corner Z (all corners have same Z for 2D opening)
                            // For floors/other: Height = Y range (vertical in rotated space)
                             double cornerMinZ = clashZonesInCluster.Count > 0 ? clashZonesInCluster.Min(cz => cz.SleeveBoundingBoxMinZ) : 0.0;
                             double cornerMaxZ = clashZonesInCluster.Count > 0 ? clashZonesInCluster.Max(cz => cz.SleeveBoundingBoxMaxZ) : 0.0;
                             double calculatedHeight = cornerMaxZ - cornerMinZ; // Height = Z range (vertical dimension)
                            
                            // ✅ CRITICAL FIX: For walls/framing, use Z range for height, not RCS Y
                            // For floors/other, cornerHeight is already calculated from Y range in rotated space
                            if (isWallOrFraming)
                            {
                                cornerHeight = calculatedHeight; // Override with Z range (vertical) for walls/framing
                            }
                            else
                            {
                                // For floors/other, cornerHeight is already correct from rotated Y range.
                                // CalculatedHeight is the Z-range (thickness), which is correctly assigned to depth below.
                                // We don't compare them here as they represent different physical dimensions for floors.
                            }
                            
                            // ✅ DEPTH: Calculate based on host type (only if not already set for mixed orientations)
                            // For walls/framing: depth = wall/framing thickness (overridden later in SetSizeParameters)
                            // For floors/other: depth = StructuralElementThickness (same as individual sleeves)
                            // ✅ DEPTH: Authoritative Source StructuralElementThickness fullstop
                            var firstCz = clashZonesInCluster.FirstOrDefault();
                            if (firstCz != null)
                            {
                                cornerDepth = firstCz.StructuralElementThickness;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ DEPTH: Set strictly to StructuralElementThickness={cornerDepth * 304.8:F1}mm (HostType={firstCz.StructuralElementType})\n");
                                }
                            }
                            
                            // ✅ ROUNDING: Apply user settings (RoundAlwaysUp, RoundingValue) from UI
                            // Fixed the "Odd Value" issue by rounding cluster dimensions properly
                            cornerWidth = OpeningSettingsHelper.RoundDimensionForCluster(cornerWidth);
                            cornerHeight = OpeningSettingsHelper.RoundDimensionForCluster(cornerHeight);
                            
                            double widthMm = cornerWidth * 304.8;
                            double heightMm = cornerHeight * 304.8;
                            double depthMm = cornerDepth * 304.8;
                            
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ CORNER-BASED SUCCESS (ROUNDED): W={widthMm:F1}mm, H={heightMm:F1}mm, D={depthMm:F1}mm (from {allCorners.Count} corners)\n");
                            
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
                                 $"  ClashZone GUIDs in cluster: {string.Join(", ", clashZonesInCluster.Select(z => z.ClashZoneGuid))}\n" +
                                 $"  Check if corners are saved to database and if SleevePersistenceService.PersistSleeveData was called.\n");
                            
                            // ✅ FALLBACK FOR STRAIGHT-AXIS FLOOR CLUSTERS: Use individual sleeve placement points if corners are missing
                            // This ensures cluster sleeves are still placed correctly even if corner saving failed
                            bool isStraightAxis = Math.Abs(rotationAngle) < 1e-6;
                            bool isFloorHost = !isWallOrFraming;
                            
                            if (isStraightAxis && isFloorHost && placementPoint.IsZeroLength())
                            {
                                // ✅ FALLBACK: Calculate placement point from individual sleeve placement points
                                var placementPoints = new List<XYZ>();
                                 foreach (var cz in clashZonesInCluster)
                                 {
                                     if (cz != null)
                                     {
                                         double px = cz.SleevePlacementPointX;
                                         double py = cz.SleevePlacementPointY;
                                         double pz = cz.SleevePlacementPointZ;
                                         
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
                            
                            // ❌ NO FALLBACK EXCEPTION: Log and SKIP this cluster if corners are insufficient.
                            // Caller (batch clustering) will simply not persist this cluster; corners themselves
                            // are still persisted by the batch corner extractor after placement.
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ❌ SKIPPING CLUSTER: Corner-based calculation failed - only {allCorners.Count} corners found (need at least 4). " +
                                $"ClashZone GUIDs: {string.Join(", ", clashZonesInCluster.Select(z => z.ClashZoneGuid))}. " +
                                $"Check corner saving logic (SleevePersistenceService.PersistSleeveData).\n");
                            // Return a zero-sized bbox; caller should treat as 'no valid cluster bbox'
                            return (0, 0, 0, XYZ.Zero, null, null, null, null, null, null);
                        }
                    }
                    else
                    {
                        // ⚠️ CRITICAL ERROR: No corners found – log and SKIP this cluster (do not throw).
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ SKIPPING CLUSTER: No corners found in ANY sleeve.\n" +
                            $"  This indicates corners were NOT saved during individual sleeve placement!\n" +
                            $"  ClashZone GUIDs in cluster: {string.Join(", ", clashZonesInCluster.Select(z => z.ClashZoneGuid))}\n" +
                            $"  Check if corners are saved to database (SleevePersistenceService.PersistSleeveData).\n");
                        return (0, 0, 0, XYZ.Zero, null, null, null, null, null, null);
                    }
                }
                catch (Exception cornerEx)
                {
                    // Log and SKIP this cluster instead of crashing the whole batch.
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                         $"[{DateTime.Now:HH:mm:ss}] ❌❌❌ CORNER CALC EXCEPTION (cluster skipped): {cornerEx.Message}\n" +
                         $"StackTrace: {cornerEx.StackTrace}\n" +
                         $"ClashZone GUIDs in cluster: {string.Join(", ", clashZonesInCluster.Select(z => z.ClashZoneGuid))}\n");
                    return (0, 0, 0, XYZ.Zero, null, null, null, null, null, null);
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
                                            new XYZ(firstCz.SleevePlacementPointX,
                                                   firstCz.SleevePlacementPointY,
                                                   firstCz.SleevePlacementPointZ)) ?? XYZ.Zero;
                                        
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
                                                    double placementX = cz.SleevePlacementPointX;
                                                    double placementY = cz.SleevePlacementPointY;
                                                    double placementZ = cz.SleevePlacementPointZ;
                                                    
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
                        double placementX = clashZone.SleevePlacementPointX;
                        double placementY = clashZone.SleevePlacementPointY;
                        double placementZ = clashZone.SleevePlacementPointZ;
                        
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
                    double placementX = cz.SleevePlacementPointX;
                    double placementY = cz.SleevePlacementPointY;
                    double placementZ = cz.SleevePlacementPointZ;
                    
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
                    if (sleeveIds.Count == cluster.Count && _rotatedBboxCache.Count < MAX_ROTATED_BBOX_CACHE_SIZE)
                    {
                        string cacheKey = $"RBB_{string.Join("_", sleeveIds)}_{rotationAngle:F6}";
                        _rotatedBboxCache[cacheKey] = result;
                        
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
            if (_clusterRotationData.TryGetValue(clusterInstanceId, out var data))
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
            _clusterRotationData[clusterInstanceId] = (rotationAngleDeg, isRotated, rotatedBboxMin, rotatedBboxMax, rotatedWidth, rotatedHeight, rotatedDepth);
        }

        /// <summary>
        /// Clear all stored rotation data
        /// </summary>
        public void ClearRotationData()
        {
            _clusterRotationData.Clear();
            _processedClusterIds.Clear(); // ✅ Also clear the deduplication tracker
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
                if (_clashZoneCache.ContainsKey(sleeveId))
                    continue;

                // Load and cache
                try
                {
                    var cz = _getClashZoneFunc(sleeveId, xmlFilePath);
                    var clashZone = cz as ClashZone;
                    
                    if (clashZone != null && _clashZoneCache.Count < MAX_CLASHZONE_CACHE_SIZE)
                    {
                        _clashZoneCache[sleeveId] = clashZone;
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
                
                if (sleeveId > 0 && clashZone != null && _clashZoneCache.Count < MAX_CLASHZONE_CACHE_SIZE)
                {
                    // Skip if already in cache
                    if (!_clashZoneCache.ContainsKey(sleeveId))
                    {
                        _clashZoneCache[sleeveId] = clashZone;
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
            if (_clashZoneCache.TryGetValue(sleeveInstanceId, out var cached))
            {
                return cached;
            }
            
            // Cache miss - load from function
            var cz = _getClashZoneFunc(sleeveInstanceId, xmlFilePath);
            var clashZone = cz as ClashZone;
            
            // Store in cache (if not null and cache not full)
            if (clashZone != null && _clashZoneCache.Count < MAX_CLASHZONE_CACHE_SIZE)
            {
                _clashZoneCache[sleeveInstanceId] = clashZone;
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
                if (cluster != null && cluster.Count > 0 && _rotatedBboxCache.Count < MAX_ROTATED_BBOX_CACHE_SIZE)
                {
                    var sleeveIds = cluster.Select(s => s?.SleeveInstanceId ?? 0).Where(id => id > 0).OrderBy(id => id).ToList();
                    if (sleeveIds.Count == cluster.Count)
                    {
                        string cacheKey = $"RBB_{string.Join("_", sleeveIds)}_{rotationAngle:F6}";
                        _rotatedBboxCache[cacheKey] = result;
                        
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
        /// ✅ CRITICAL FIX: Calculate cluster placement point using HYBRID logic
        /// - CENTROID (Average): Used for Width (Along Wall) and Height (Z) axes.
        /// - FIRST SLEEVE: Used for Perpendicular (Wall Thickness) axis to ensure wall center alignment.
        /// - User Request: "placement point is at wall center same as first sleeve... get the centroid for width and height only"
        /// </summary>
        private XYZ CalculatePlacementPointFromIntersections(List<dynamic> cluster, string? xmlFilePath)
        {
            if (cluster == null || cluster.Count == 0)
                return XYZ.Zero;

            // ✅ Step 1: Calculate CENTROID (Average) first
            double sumX = 0.0, sumY = 0.0, sumZ = 0.0;
            int validCount = 0;
            ClashZone? firstValidCz = null;
            int firstSleeveId = 0;

            // Collect all points and find first valid sleeve for metadata
            foreach (var sleeveData in cluster)
            {
                if (sleeveData == null || sleeveData.SleeveInstanceId <= 0) continue;
                
                try
                {
                    var cz = GetCachedClashZone(sleeveData.SleeveInstanceId, xmlFilePath);
                    if (cz == null) continue;

                    // Capture first valid sleeve for Host/Orientation checks
                    if (firstValidCz == null)
                    {
                        firstValidCz = cz;
                        firstSleeveId = sleeveData.SleeveInstanceId;
                    }

                    double ipX = cz.IntersectionPointX;
                    double ipY = cz.IntersectionPointY;
                    double ipZ = cz.IntersectionPointZ;

                    if ((ipX != 0.0 || ipY != 0.0 || ipZ != 0.0) &&
                        !double.IsNaN(ipX) && !double.IsInfinity(ipX))
                    {
                        sumX += ipX;
                        sumY += ipY;
                        sumZ += ipZ;
                        validCount++;
                    }
                }
                catch { }
            }

            if (validCount == 0 || firstValidCz == null)
            {
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ No valid intersection points found in cluster ({cluster.Count} sleeves)\n");
                return XYZ.Zero;
            }

            // Initial Centroid (Average of all sleeves)
            XYZ centroid = new XYZ(sumX / validCount, sumY / validCount, sumZ / validCount);

            // ✅ Step 2: Determine if Hybrid Logic is needed (Wall/Framing)
            string hostType = firstValidCz.StructuralElementType ?? "";
            bool isWall = hostType.StartsWith("Wall", StringComparison.OrdinalIgnoreCase) || 
                          hostType.Equals("Structural Framing", StringComparison.OrdinalIgnoreCase);

            if (isWall)
            {
                // ✅ WALL LOGIC: Override Perpendicular Axis with First Sleeve's coordinate
                // "placement point is at wall center same as first sleeve"
                
                string orientation = (firstValidCz.HostOrientation ?? "").Trim().ToUpper();
                double finalX = centroid.X;
                double finalY = centroid.Y;
                double finalZ = centroid.Z; // Height always uses Centroid (Average Z)

                if (orientation.Contains("X"))
                {
                    // X-WALL: Wall runs along X-axis.
                    // - Parallel (Width): X (Keep Centroid)
                    // - Perpendicular (Thickness): Y (Override with First Sleeve)
                    finalY = firstValidCz.IntersectionPointY;
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log", $"[{DateTime.Now:HH:mm:ss}] 📐 HYBRID X-WALL: Override Y (Thk)={finalY:F6}, Keep X (Len)={finalX:F6}\n");
                }
                else if (orientation.Contains("Y"))
                {
                    // Y-WALL: Wall runs along Y-axis.
                    // - Parallel (Width): Y (Keep Centroid)
                    // - Perpendicular (Thickness): X (Override with First Sleeve)
                    finalX = firstValidCz.IntersectionPointX;

                    if (!DeploymentConfiguration.DeploymentMode)
                        SafeFileLogger.SafeAppendText("cluster_sizing.log", $"[{DateTime.Now:HH:mm:ss}] 📐 HYBRID Y-WALL: Override X (Thk)={finalX:F6}, Keep Y (Len)={finalY:F6}\n");
                }
                else
                {
                    // Fallback (Unknown Orientation): Stick to First Sleeve complete override for safety
                    finalX = firstValidCz.IntersectionPointX;
                    finalY = firstValidCz.IntersectionPointY;
                    // Keep Centroid Z for height centering
                }

                XYZ hybridPoint = new XYZ(finalX, finalY, finalZ);
                return hybridPoint;
            }

            // ✅ FLOOR LOGIC: Pure Centroid
            if (!DeploymentConfiguration.DeploymentMode)
            {
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ✅ PLACEMENT (FLOOR): Centroid ({centroid.X:F6}, {centroid.Y:F6}, {centroid.Z:F6})\n");
            }
            return centroid;
        }
    }
}
