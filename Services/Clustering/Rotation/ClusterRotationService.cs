using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.BoundingBox;
using JSE_RevitAddin_MEP_OPENINGS.Services;

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
            
            _clusterRotationData = new Dictionary<int, (double, bool, XYZ, XYZ, double, double, double)>();
            _getClashZoneFunc = getClashZoneFunc;
        }

        /// <summary>
        /// Determine the dominant rotation angle for a cluster of sleeves
        /// Returns 0 for straight axis-aligned clusters (aligned to WCS: 0°, 90°, 180°, 270°)
        /// Returns rotation angle for rotated axis-aligned clusters (non-straight: 45°, 225°, etc.)
        /// 
        /// ✅ CRITICAL: Pipes (circular elements) always return 0.0 - no rotation needed, place straight to WCS
        /// </summary>
        public double DetermineRotationAngle(List<dynamic> cluster, string? xmlFilePath = null)
        {
            try
            {
                if (cluster == null || cluster.Count == 0)
                    return 0.0;

                // ✅ CRITICAL: Check if cluster contains pipes or round ducts (circular elements)
                // Circular elements (pipes and round ducts) should always be placed straight to WCS (axis-aligned), no rotation needed
                bool isCircularElementCluster = false;
                string circularElementType = "";
                foreach (var sleeveData in cluster)
                {
                    if (sleeveData?.ClashZone == null)
                        continue;

                    var clashZone = sleeveData.ClashZone as ClashZone;
                    if (clashZone == null)
                        continue;

                    // Check if this is a pipe (circular element)
                    if (string.Equals(clashZone.MepElementCategory, "Pipes", StringComparison.OrdinalIgnoreCase))
                    {
                        isCircularElementCluster = true;
                        circularElementType = "PIPE";
                        break;
                    }
                    
                    // Check if this is a round duct (circular element)
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

                // ✅ CIRCULAR ELEMENT FIX: Pipes and round ducts are circular elements - always place straight to WCS (no rotation)
                if (isCircularElementCluster)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CLUSTER-ANGLE] {circularElementType} cluster detected: Returning 0.0° (straight axis-aligned to WCS, no rotation needed for circular elements)");
                    }
                    return 0.0;
                }

                // ✅ WALL ORIENTATION LOGIC (PRIMARY): For walls/structural framing, rotation is based on wall orientation
                // This is the MAIN logic for wall-hosted clusters, not a fallback
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
                        // ✅ WALL ROTATION LOGIC (PRIMARY): Use WallDirectionType to determine X-wall or Y-wall
                        // CRITICAL: Only check WALL orientation (X-wall or Y-wall), NOT MEP orientation!
                        // For wall clustering, rotation is determined ONLY by the wall's direction type
                        // Individual sleeves: LEFT view family
                        //   - X-walls: +90° rotation (LEFT family needs rotation for X-walls)
                        //   - Y-walls: 0° rotation (LEFT family works naturally for Y-walls)
                        // Cluster sleeves should use the SAME rotation as individual sleeves to maintain correct orientation
                        
                        // ✅ DIAGNOSTIC: Log all wall orientation data for debugging
                        string wallDirectionType = firstClashZone.WallDirectionType ?? "";
                        string hostOrientation = firstClashZone.HostOrientation ?? "";
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE-DEBUG] WALL HOST DETECTED: StructuralType='{firstClashZone.StructuralElementType}', WallDirectionType='{wallDirectionType}', HostOrientation='{hostOrientation}'\n");
                        }
                        
                        // ✅ CRITICAL: Check WallDirectionType directly (contains "X-WALL" or "Y-WALL")
                        // DO NOT use MepElementOrientationDirection - that can contain MEP orientation for floors
                        bool isXWall = wallDirectionType.Contains("X-WALL", StringComparison.OrdinalIgnoreCase);
                        bool isYWall = wallDirectionType.Contains("Y-WALL", StringComparison.OrdinalIgnoreCase);
                        
                        if (isXWall)
                        {
                            // X-wall: +90° rotation required (same as individual sleeves)
                            // LEFT view family extrudes along Y-axis, needs +90° rotation for X-walls
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-ANGLE] WALL ROTATION: {firstClashZone.StructuralElementType} - X-WALL detected (WallDirectionType='{wallDirectionType}') → Returning 90.0° (π/2 radians - matches individual sleeve rotation)");
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE] ✅ X-WALL DETECTED via WallDirectionType → 90.0°\n");
                            }
                            return Math.PI / 2.0; // 90° rotation for X-oriented walls/framing (same as individual sleeves)
                        }
                        else if (isYWall)
                        {
                            // Y-wall: 0° rotation (same as individual sleeves)
                            // LEFT view family works naturally for Y-walls, no rotation needed
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[CLUSTER-ANGLE] WALL ROTATION: {firstClashZone.StructuralElementType} - Y-WALL detected (WallDirectionType='{wallDirectionType}') → Returning 0.0° (no rotation - matches individual sleeve rotation)");
                                SafeFileLogger.SafeAppendText("cluster_debug.log",
                                    $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE] ✅ Y-WALL DETECTED via WallDirectionType → 0.0°\n");
                            }
                            return 0.0; // No rotation for Y-oriented walls/framing (same as individual sleeves)
                        }
                        else
                        {
                            // ✅ FALLBACK: If WallDirectionType doesn't contain X-WALL or Y-WALL, check HostOrientation
                            // HostOrientation should contain "X" or "Y" for walls/framing (this is what groupKey uses)
                            if (string.Equals(hostOrientation, "X", StringComparison.OrdinalIgnoreCase))
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[CLUSTER-ANGLE] WALL ROTATION: {firstClashZone.StructuralElementType} - HostOrientation='X' (fallback from WallDirectionType='{wallDirectionType}') → Returning 90.0°");
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE] ✅ X-WALL DETECTED via HostOrientation fallback → 90.0°\n");
                                }
                                return Math.PI / 2.0;
                            }
                            else if (string.Equals(hostOrientation, "Y", StringComparison.OrdinalIgnoreCase))
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Info($"[CLUSTER-ANGLE] WALL ROTATION: {firstClashZone.StructuralElementType} - HostOrientation='Y' (fallback from WallDirectionType='{wallDirectionType}') → Returning 0.0°");
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE] ✅ Y-WALL DETECTED via HostOrientation fallback → 0.0°\n");
                                }
                                return 0.0;
                            }
                            else
                            {
                                // ✅ DIAGNOSTIC: Log when wall direction cannot be determined
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                                        $"[{DateTime.Now:HH:mm:ss}] [CLUSTER-ANGLE] ⚠️ WALL HOST but cannot determine X/Y: WallDirectionType='{wallDirectionType}', HostOrientation='{hostOrientation}' → Continuing to MEP rotation calculation\n");
                                }
                            }
                        }
                        // If wall direction cannot be determined, continue to calculate from MEP rotation angles below
                    }
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
        CalculateRotatedBoundingBox(List<dynamic> cluster, List<FamilyInstance> actualSleeves, double rotationAngle, string? xmlFilePath = null)
        {
            // Simplified rotated bounding box calculation (union + optional corner refinement) without external ambiguous loggers.
            if (cluster == null || cluster.Count == 0 || actualSleeves == null || actualSleeves.Count == 0)
                return (0,0,0,XYZ.Zero,null,null,null,null,null,null);

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
                
                var cz = _getClashZoneFunc(sleeveInstanceId, xmlFilePath);
                if (cz == null)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ⚠️ Sleeve {sleeveInstanceId}: ClashZone is NULL\n");
                    continue;
                }
                
                // ✅ CRASH-SAFE: Cast to ClashZone to avoid dynamic binding issues
                var clashZone = cz as ClashZone;
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

            // ✅ CRITICAL FIX: For rotated axis-aligned clusters (non-straight axis-aligned to WCS), use corner-based watertight algorithm
            // Simple union of rotated bounding boxes gives oversized results for diagonal arrangements
            // ⚠️ IMPORTANT: Corner-based algorithm only needs corners + rotation, NOT rotated bounding boxes
            // Rotated bounding boxes are only saved for rotated axis-aligned sleeves (45°, 225°, etc.)
            // But individual sleeves might be straight axis-aligned (0°, 90°) so rotatedBboxes.Count=0
            // However, corners are ALWAYS saved for all sleeves, so we can use corner-based calculation
            if (Math.Abs(rotationAngle) > 1e-6)
            {
                // ✅ WATERTIGHT ALGORITHM: Use corner-based calculation for accurate sizing
                // This works even if rotated bounding boxes aren't in database, as long as corners are available
                try
                {
                    // ✅ DIAGNOSTIC: Log before attempting corner-based calculation
                    double rotationDeg = rotationAngle * 180.0 / Math.PI;
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] 🔍 ATTEMPTING corner-based calculation: rotation={rotationDeg:F1}° (rotated axis-aligned), clusterSize={cluster.Count}, rotatedBboxesCount={rotatedBboxes.Count}\n");
                    
                    // ✅ PRE-CHECK: Verify at least one sleeve has corners before attempting calculation
                    // Corners are ALWAYS saved for all sleeves during individual placement
                    bool hasAnyCorners = false;
                    int sleevesWithCorners = 0;
                    int sleevesWithoutCorners = 0;
                    foreach (var sleeveData in cluster)
                    {
                        var cz = _getClashZoneFunc(sleeveData.SleeveInstanceId, xmlFilePath);
                        if (cz != null)
                        {
                            var clashZone = cz as Models.ClashZone;
                            if (clashZone != null)
                            {
                                // ✅ DIAGNOSTIC: Check individual sleeve rotation angle
                                double sleeveRotationDeg = clashZone.MepElementRotationAngle * 180.0 / Math.PI;
                                
                                bool hasCorners = clashZone.SleeveCorner1X.HasValue && clashZone.SleeveCorner1Y.HasValue &&
                                                  clashZone.SleeveCorner2X.HasValue && clashZone.SleeveCorner2Y.HasValue &&
                                                  clashZone.SleeveCorner3X.HasValue && clashZone.SleeveCorner3Y.HasValue &&
                                                  clashZone.SleeveCorner4X.HasValue && clashZone.SleeveCorner4Y.HasValue;
                                
                                if (hasCorners)
                                {
                                    sleevesWithCorners++;
                                    hasAnyCorners = true;
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ✅ Sleeve {sleeveData.SleeveInstanceId}: Has corners, Rotation={sleeveRotationDeg:F1}°, " +
                                        $"ActiveCoords=({clashZone.SleevePlacementPointActiveDocumentX:F6}, {clashZone.SleevePlacementPointActiveDocumentY:F6}, {clashZone.SleevePlacementPointActiveDocumentZ:F6})\n");
                                }
                                else
                                {
                                    sleevesWithoutCorners++;
                                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                        $"[{DateTime.Now:HH:mm:ss}]   ⚠️ Sleeve {sleeveData.SleeveInstanceId}: Missing corners, Rotation={sleeveRotationDeg:F1}°, " +
                                        $"ActiveCoords=({clashZone.SleevePlacementPointActiveDocumentX:F6}, {clashZone.SleevePlacementPointActiveDocumentY:F6}, {clashZone.SleevePlacementPointActiveDocumentZ:F6}), " +
                                        $"HasCorner1X={clashZone.SleeveCorner1X.HasValue}, HasCorner1Y={clashZone.SleeveCorner1Y.HasValue}\n");
                                }
                            }
                        }
                    }
                    
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] 📊 Corner check: {sleevesWithCorners}/{cluster.Count} sleeves have corners, cluster rotation={rotationAngle * 180 / Math.PI:F1}°\n");
                    
                    if (!hasAnyCorners)
                    {
                        SafeFileLogger.SafeAppendText("cluster_sizing.log",
                            $"[{DateTime.Now:HH:mm:ss}] ⚠️ Skipping corner-based: rotation={rotationAngle * 180 / Math.PI:F1}°, no corners found in database, rotatedBboxesCount={rotatedBboxes.Count}\n");
                    }
                    else
                    {
                        // ✅ WRAP: Convert Func<int, string, ClashZone> to Func<int, string, dynamic> for corner calculator
                        Func<int, string, dynamic> getClashZoneDynamic = (sleeveId, path) => _getClashZoneFunc(sleeveId, path);
                        
                        var cornerResult = CornerBasedBoundingBoxCalculator.CalculateFromCorners(
                            cluster,
                            rotationAngle,
                            out XYZ origin,
                            getClashZoneDynamic,
                            xmlFilePath);

                        if (cornerResult.HasValue)
                        {
                            // Calculate depth from Z coordinates (use union of Z extents from actual sleeves, not rotatedBboxes which might be empty)
                            var clashZonesForDepth = cluster.Select(s => _getClashZoneFunc(s.SleeveInstanceId, xmlFilePath) as Models.ClashZone)
                                .Where(cz => cz != null)
                                .ToList();
                            double cornerMinZ = clashZonesForDepth.Count > 0 ? clashZonesForDepth.Min(cz => cz.SleeveBoundingBoxMinZ) : 0.0;
                            double cornerMaxZ = clashZonesForDepth.Count > 0 ? clashZonesForDepth.Max(cz => cz.SleeveBoundingBoxMaxZ) : 0.0;
                            double cornerDepth = cornerMaxZ - cornerMinZ;
                        
                        // ✅ MIDPOINT: Calculate in rotated coordinate space, then transform back to world coordinates
                        // The midpoint is calculated as center of bounding box in rotated space: ((minX+maxX)/2, (minY+maxY)/2)
                        // To transform back to world coordinates, we need the INVERSE rotation matrix (transpose)
                        // Forward: clusterX = relX * cos(θ) - relY * sin(θ), clusterY = relX * sin(θ) + relY * cos(θ)
                        // Inverse: worldX = clusterX * cos(θ) + clusterY * sin(θ), worldY = -clusterX * sin(θ) + clusterY * cos(θ)
                        double midX = (cornerResult.Value.minX + cornerResult.Value.maxX) / 2.0;
                        double midY = (cornerResult.Value.minY + cornerResult.Value.maxY) / 2.0;
                        double midZ = (cornerMinZ + cornerMaxZ) / 2.0;
                        
                        // ✅ INVERSE TRANSFORMATION: Transpose of rotation matrix (inverse for rotation matrices)
                        double cosA = Math.Cos(rotationAngle);  // Use same angle (inverse rotation matrix is transpose)
                        double sinA = Math.Sin(rotationAngle);
                        double midRotatedBackX = midX * cosA + midY * sinA;  // Note: + instead of -
                        double midRotatedBackY = -midX * sinA + midY * cosA;  // Note: -sinA instead of sinA
                        
                        // Translate back (add origin)
                        XYZ cornerMid = new XYZ(
                            origin.X + midRotatedBackX,
                            origin.Y + midRotatedBackY,
                            midZ  // Z stays the same
                        );
                        
                        // Convert to millimeters for logging
                        double widthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerResult.Value.width);
                        double heightMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerResult.Value.height);
                        double depthMm = RevitUnitConversionService.Instance.FromInternalMillimeters(cornerDepth);
                        
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ✅ CORNER-BASED: Cluster rotation={rotationAngle * 180 / Math.PI:F1}°, W={widthMm:F1}mm, H={heightMm:F1}mm, D={depthMm:F1}mm\n");
                            
                            return (cornerResult.Value.width, cornerResult.Value.height, cornerDepth, cornerMid,
                                cornerResult.Value.minX, cornerResult.Value.minY, cornerMinZ,
                                cornerResult.Value.maxX, cornerResult.Value.maxY, cornerMaxZ);
                        }
                        else
                        {
                            SafeFileLogger.SafeAppendText("cluster_sizing.log",
                                $"[{DateTime.Now:HH:mm:ss}] ⚠️ Corner-based calculation returned NULL, falling back to union for rotation={rotationAngle * 180 / Math.PI:F1}°\n");
                        }
                    }
                }
                catch (Exception cornerEx)
                {
                    SafeFileLogger.SafeAppendText("cluster_sizing.log",
                        $"[{DateTime.Now:HH:mm:ss}] ❌ Corner-based calculation EXCEPTION: {cornerEx.Message}\nStackTrace: {cornerEx.StackTrace}\nFalling back to union.\n");
                }
            }
            else
            {
                double rotationDeg = rotationAngle * 180.0 / Math.PI;
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ Skipping corner-based: rotation={rotationDeg:F1}° (straight axis-aligned to WCS), isNearZero={Math.Abs(rotationAngle) <= 1e-6}\n");
            }

            if (rotatedBboxes.Count == 0)
            {
                // Fallback: axis-aligned union from Revit bounding boxes
                var revitBboxes = actualSleeves.Select(s => s.get_BoundingBox(null)).Where(b => b != null && b.Enabled).ToList();
                if (revitBboxes.Count == 0) return (0,0,0,XYZ.Zero,null,null,null,null,null,null);
                double minXf = revitBboxes.Min(b=>b.Min.X); double minYf = revitBboxes.Min(b=>b.Min.Y); double minZf = revitBboxes.Min(b=>b.Min.Z);
                double maxXf = revitBboxes.Max(b=>b.Max.X); double maxYf = revitBboxes.Max(b=>b.Max.Y); double maxZf = revitBboxes.Max(b=>b.Max.Z);
                double wf = maxXf - minXf; double hf = maxYf - minYf; double df = maxZf - minZf; XYZ midF = new XYZ((minXf+maxXf)/2,(minYf+maxYf)/2,(minZf+maxZf)/2);
                
                SafeFileLogger.SafeAppendText("cluster_sizing.log",
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ FALLBACK (axis-aligned): W={wf:F1}mm, H={hf:F1}mm, D={df:F1}mm (no rotated bboxes)\n");
                
                return (wf,hf,df,midF,null,null,null,null,null,null);
            }

            // ✅ FALLBACK: Simple union of rotated boxes (only if corner-based failed)
            // This is less accurate but better than nothing
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
                    $"[{DateTime.Now:HH:mm:ss}] ⚠️ UNION (fallback): W={unionWidthMm:F1}mm, H={unionHeightMm:F1}mm, D={unionDepthMm:F1}mm, Rotation={rotationAngle * 180 / Math.PI:F1}°\n");
            
            return (width,height,depth,mid,minX,minY,minZ,maxX,maxY,maxZ);
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
        }
    }
}
