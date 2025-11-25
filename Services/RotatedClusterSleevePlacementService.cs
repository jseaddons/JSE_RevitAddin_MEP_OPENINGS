using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service class for placing non-axis-aligned rotated cluster sleeves.
    /// Handles rotation angle determination, bounding box calculation in rotated coordinate system,
    /// and placement of cluster sleeves with proper rotation.
    /// </summary>
    public class RotatedClusterSleevePlacementService
    {
        private readonly Func<int, string, ClashZone> _getClashZoneBySleeveInstanceId;
        private readonly Func<List<dynamic>, string, double> _determineDominantRotationAngle;
        private readonly Func<Document, List<dynamic>, List<FamilyInstance>, double, string, (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)> _getClusterBoundingBoxWithRotatedCoordinates;
        private readonly Func<Document, FamilyInstance, List<dynamic>, SleeveGroupKey, double, double, double, bool, bool> _setClusterSizeParameters;
        private readonly Func<Document, dynamic, Level> _getReferenceLevelFromXml;

        /// <summary>
        /// Helper struct for grouping key
        /// </summary>
        public struct SleeveGroupKey
        {
            public string hostType;
            public string systemType;
            public string orientation;

            public SleeveGroupKey(string hostType, string systemType, string orientation)
            {
                this.hostType = hostType;
                this.systemType = systemType;
                this.orientation = orientation;
            }
        }

        /// <summary>
        /// Constructor with dependency injection for required functions
        /// </summary>
        public RotatedClusterSleevePlacementService(
            Func<int, string, ClashZone> getClashZoneBySleeveInstanceId,
            Func<List<dynamic>, string, double> determineDominantRotationAngle,
            Func<Document, List<dynamic>, List<FamilyInstance>, double, string, (double width, double height, double depth, XYZ mid, double? rotatedMinX, double? rotatedMinY, double? rotatedMinZ, double? rotatedMaxX, double? rotatedMaxY, double? rotatedMaxZ)> getClusterBoundingBoxWithRotatedCoordinates,
            Func<Document, FamilyInstance, List<dynamic>, SleeveGroupKey, double, double, double, bool, bool> setClusterSizeParameters,
            Func<Document, dynamic, Level> getReferenceLevelFromXml)
        {
            _getClashZoneBySleeveInstanceId = getClashZoneBySleeveInstanceId;
            _determineDominantRotationAngle = determineDominantRotationAngle;
            _getClusterBoundingBoxWithRotatedCoordinates = getClusterBoundingBoxWithRotatedCoordinates;
            _setClusterSizeParameters = setClusterSizeParameters;
            _getReferenceLevelFromXml = getReferenceLevelFromXml;
        }

        /// <summary>
        /// Place a non-axis-aligned rotated cluster sleeve
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="cluster">List of sleeve data from XML</param>
        /// <param name="actualSleeves">List of actual Revit FamilyInstance sleeves</param>
        /// <param name="groupKey">Group key for the cluster</param>
        /// <param name="familySymbol">Family symbol to use for cluster sleeve</param>
        /// <param name="xmlFilePath">Optional XML file path for data loading</param>
        /// <returns>Placed FamilyInstance cluster sleeve, or null if placement failed</returns>
        public FamilyInstance PlaceRotatedClusterSleeve(
            Document doc,
            List<dynamic> cluster,
            List<FamilyInstance> actualSleeves,
            SleeveGroupKey groupKey,
            FamilySymbol familySymbol,
            string xmlFilePath = null)
        {
            try
            {
                if (cluster == null || cluster.Count == 0 || actualSleeves == null || actualSleeves.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[RotatedClusterSleevePlacementService] Invalid input: cluster or actualSleeves is null or empty");
                    return null;
                }

                // Step 1: Determine rotation angle
                double rotationAngle = _determineDominantRotationAngle(cluster, xmlFilePath);

                // Step 2: Apply fallback logic for walls/framing if rotation is 0
                if (Math.Abs(rotationAngle) < 1e-6)
                {
                    if (groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing")
                    {
                        string xmlOrientation = groupKey.orientation ?? "Unknown";
                        if (xmlOrientation.Equals("X", StringComparison.OrdinalIgnoreCase))
                        {
                            rotationAngle = Math.PI / 2;  // 90° rotation for X-oriented walls/framing
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[RotatedClusterSleevePlacementService] ⚠️ FALLBACK: {groupKey.hostType} with X orientation → Setting rotation to 90°");
                            }
                        }
                        else if (xmlOrientation.Equals("Y", StringComparison.OrdinalIgnoreCase))
                        {
                            rotationAngle = 0.0;  // No rotation for Y-oriented walls/framing
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Info($"[RotatedClusterSleevePlacementService] ⚠️ FALLBACK: {groupKey.hostType} with Y orientation → Keeping rotation at 0°");
                            }
                        }
                    }
                }

                if (!DeploymentConfiguration.DeploymentMode && Math.Abs(rotationAngle) > 1e-6)
                {
                    DebugLogger.Info($"[RotatedClusterSleevePlacementService] Using rotation angle: {rotationAngle * 180 / Math.PI:F1}° for cluster bounding box calculation");
                }

                // Step 3: Get cluster bounding box with rotated coordinates
                var (width, height, depth, mid, rotatedMinX, rotatedMinY, rotatedMinZ, rotatedMaxX, rotatedMaxY, rotatedMaxZ) = 
                    _getClusterBoundingBoxWithRotatedCoordinates(doc, cluster, actualSleeves, rotationAngle, xmlFilePath);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Log($"[RotatedClusterSleevePlacementService] ✅ ROTATED COORDINATE SYSTEM: Cluster bounding box - Width={width:F3}, Height={height:F3}, Depth={depth:F3}, Angle={rotationAngle * 180 / Math.PI:F1}°");
                }

                // Step 4: Get reference level
                Level? refLevel = _getReferenceLevelFromXml(doc, cluster[0]);
                if (refLevel == null)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Log($"Reference level not found for cluster sleeve. Skipping cluster.");
                    return null;
                }

                // Step 5: Place cluster sleeve
                FamilyInstance inst = doc.Create.NewFamilyInstance(mid, familySymbol, refLevel!, StructuralType.NonStructural);

                // ✅ CRITICAL: DO NOT rotate the cluster sleeve by rotationAngle
                // The rotationAngle is the cluster's "intended rotated axis" used ONLY for bounding box calculation
                // The cluster sleeve itself should be axis-aligned (0° rotation) because:
                // 1. The bounding box calculation already accounts for the rotation of individual sleeves
                // 2. The cluster sleeve dimensions (width x height) already represent the rotated bounding box size
                // 3. The cluster sleeve is a new element that should be placed axis-aligned
                // 
                // The rotationAngle is used to:
                // - Calculate bounding box in rotated coordinate system (GetClusterBoundingBoxWithRotatedCoordinates)
                // - Transform corners to rotated space for min/max calculation
                // - Transform the midpoint back to world coordinates for placement
                // 
                // But the cluster sleeve element itself should NOT be rotated - it's axis-aligned
                // Step 6: Skip rotation (cluster sleeve is axis-aligned, dimensions already account for rotation)
                if (!DeploymentConfiguration.DeploymentMode && Math.Abs(rotationAngle) > 1e-6)
                {
                    double angleDegrees = rotationAngle * 180 / Math.PI;
                    DebugLogger.Info($"[RotatedClusterSleevePlacementService] ⚠️ Cluster bounding box calculated with rotation angle {angleDegrees:F1}° (for coordinate system only), but cluster sleeve is placed axis-aligned (0°) - dimensions already account for rotation");
                }

                // Step 7: Calculate rotated bounding box for storage
                bool isRotated = Math.Abs(rotationAngle) > 1e-6;
                XYZ rotatedBboxMin = XYZ.Zero;
                XYZ rotatedBboxMax = XYZ.Zero;
                double rotatedWidth = width;
                double rotatedHeight = height;
                double rotatedDepth = depth;

                if (isRotated)
                {
                    // Calculate rotated bounding box in rotated coordinate system
                    double halfWidth = width / 2.0;
                    double halfHeight = height / 2.0;
                    double halfDepth = depth / 2.0;

                    rotatedBboxMin = new XYZ(-halfWidth, -halfHeight, -halfDepth);
                    rotatedBboxMax = new XYZ(halfWidth, halfHeight, halfDepth);

                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[RotatedClusterSleevePlacementService] Rotated bounding box in rotated coordinate system: Min=({rotatedBboxMin.X:F3}, {rotatedBboxMin.Y:F3}, {rotatedBboxMin.Z:F3}), Max=({rotatedBboxMax.X:F3}, {rotatedBboxMax.Y:F3}, {rotatedBboxMax.Z:F3})");
                    }
                }
                else
                {
                    // For axis-aligned, use the actual bounding box from sleeves
                    var allBboxes = actualSleeves.Select(s => s.get_BoundingBox(null)).Where(b => b != null).ToList();
                    if (allBboxes.Count > 0)
                    {
                        rotatedBboxMin = new XYZ(allBboxes.Min(b => b.Min.X), allBboxes.Min(b => b.Min.Y), allBboxes.Min(b => b.Min.Z));
                        rotatedBboxMax = new XYZ(allBboxes.Max(b => b.Max.X), allBboxes.Max(b => b.Max.Y), allBboxes.Max(b => b.Max.Z));
                    }
                }

                // Step 8: Set size parameters
                bool shouldSwapDimensions = ((groupKey.hostType == "Wall" || groupKey.hostType == "Structural Framing") && rotationAngle != 0.0);
                _setClusterSizeParameters(doc, inst, cluster, groupKey, width, height, depth, shouldSwapDimensions);

                // Step 9: Log size application
                LogClusterSizeApplication(inst, width, height, depth, rotationAngle, groupKey);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[RotatedClusterSleevePlacementService] ✅ Successfully placed rotated cluster sleeve {inst.Id.IntegerValue}: W={width * 304.8:F1}mm × H={height * 304.8:F1}mm, Angle={rotationAngle * 180 / Math.PI:F1}°");
                }

                // Step 10: Database persistence
                // Persistence of cluster sleeve bbox/rotation is centralized in RefactoredClusterService
                // via BatchSaveClusterDataToDatabase/SaveClusterDataToDatabase after placement.
                // Do not save here to avoid duplicate writes and inconsistent data paths.
                return inst;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[RotatedClusterSleevePlacementService] Error placing rotated cluster sleeve: {ex.Message}\n{ex.StackTrace}");
                }
                return null;
            }
        }

        /// <summary>
        /// Apply rotation to cluster sleeve if angle is non-axis-aligned
        /// </summary>
        /// <param name="doc">Revit document</param>
        /// <param name="inst">Cluster sleeve instance</param>
        /// <param name="placementPoint">Placement point (midpoint)</param>
        /// <param name="rotationAngle">Rotation angle in radians</param>
        public void ApplyRotationToClusterSleeve(Document doc, FamilyInstance inst, XYZ placementPoint, double rotationAngle)
        {
            if (Math.Abs(rotationAngle) < 1e-6)
                return;

            // Double-check: if angle is close to 0°, 90°, 180°, or 270°, don't rotate
            double angleDegrees = rotationAngle * 180 / Math.PI;
            // Normalize to 0-360 range
            while (angleDegrees < 0) angleDegrees += 360;
            while (angleDegrees >= 360) angleDegrees -= 360;

            double distTo0 = Math.Min(angleDegrees, 360 - angleDegrees);
            double distTo90 = Math.Abs(angleDegrees - 90);
            double distTo180 = Math.Abs(angleDegrees - 180);
            double distTo270 = Math.Abs(angleDegrees - 270);

            double thresholdDegrees = 2.0; // 2 degree tolerance
            bool isAxisAligned = distTo0 < thresholdDegrees || distTo90 < thresholdDegrees ||
                                distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;

            if (isAxisAligned)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[RotatedClusterSleevePlacementService] ⚠️ Skipping rotation for cluster sleeve {inst.Id}: Angle {angleDegrees:F1}° is axis-aligned (distTo0={distTo0:F1}°, distTo90={distTo90:F1}°, distTo180={distTo180:F1}°, distTo270={distTo270:F1}°)");
                }
            }
            else
            {
                XYZ axisOrigin = placementPoint;
                XYZ axisDirection = XYZ.BasisZ;
                Line rotationAxis = Line.CreateBound(axisOrigin, axisOrigin + axisDirection);
                ElementTransformUtils.RotateElement(doc, inst.Id, rotationAxis, rotationAngle);

                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[RotatedClusterSleevePlacementService] ✅ Applied rotation {angleDegrees:F1}° to cluster sleeve {inst.Id}");
                }
            }
        }

        /// <summary>
        /// Check if an angle is axis-aligned (0°, 90°, 180°, 270°)
        /// </summary>
        /// <param name="angleRad">Angle in radians</param>
        /// <returns>True if angle is axis-aligned</returns>
        public static bool IsAxisAlignedAngle(double angleRad)
        {
            double angleDeg = angleRad * 180.0 / Math.PI;
            // Normalize to 0-360 range
            while (angleDeg < 0) angleDeg += 360;
            while (angleDeg >= 360) angleDeg -= 360;

            double thresholdDegrees = 2.0; // 2 degree tolerance
            double distTo0 = Math.Min(angleDeg, 360 - angleDeg);
            double distTo90 = Math.Abs(angleDeg - 90);
            double distTo180 = Math.Abs(angleDeg - 180);
            double distTo270 = Math.Abs(angleDeg - 270);

            return distTo0 < thresholdDegrees || distTo90 < thresholdDegrees ||
                   distTo180 < thresholdDegrees || distTo270 < thresholdDegrees;
        }

        /// <summary>
        /// Check if two rotated sleeves should be clustered based on proximity
        /// </summary>
        /// <param name="sleeve1">First sleeve data</param>
        /// <param name="sleeve2">Second sleeve data</param>
        /// <param name="rotationAngle">Shared rotation angle in radians</param>
        /// <param name="toleranceDist">Tolerance distance for clustering</param>
        /// <param name="getClashZoneBySleeveInstanceId">Function to get ClashZone by sleeve instance ID</param>
        /// <param name="boundingBoxesOverlapFromXml">Fallback function for axis-aligned check</param>
        /// <returns>True if sleeves should be clustered</returns>
        public static bool CheckRotatedSleeveProximity(
            dynamic sleeve1,
            dynamic sleeve2,
            double rotationAngle,
            double toleranceDist,
            Func<int, string, ClashZone> getClashZoneBySleeveInstanceId,
            Func<dynamic, dynamic, double, bool> boundingBoxesOverlapFromXml)
        {
            try
            {
                string clusterDebugLogPath = SafeFileLogger.GetLogFilePath("cluster_debug.log");
                bool shouldLog = !DeploymentConfiguration.DeploymentMode;

                if (shouldLog)
                {
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== CHECK ROTATED SLEEVE PROXIMITY ==========\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {sleeve1.SleeveInstanceId} vs Sleeve {sleeve2.SleeveInstanceId}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotation Angle: {rotationAngle * 180.0 / Math.PI:F2}°\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Tolerance: {toleranceDist * 304.8:F1}mm\n");
                }

                if (sleeve1?.ClashZone == null || sleeve2?.ClashZone == null)
                {
                    if (shouldLog)
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ❌ ClashZone is null - returning false\n\n");
                    }
                    return false;
                }

                var cz1 = sleeve1.ClashZone as ClashZone;
                var cz2 = sleeve2.ClashZone as ClashZone;

                if (cz1 == null || cz2 == null)
                {
                    if (shouldLog)
                    {
                        System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ❌ ClashZone cast failed - returning false\n\n");
                    }
                    return false;
                }

                // Get sleeve centers from database (placement points)
                XYZ center1 = cz1.SleevePlacementPoint ?? new XYZ(
                    (cz1.SleeveBoundingBoxMinX + cz1.SleeveBoundingBoxMaxX) / 2.0,
                    (cz1.SleeveBoundingBoxMinY + cz1.SleeveBoundingBoxMaxY) / 2.0,
                    (cz1.SleeveBoundingBoxMinZ + cz1.SleeveBoundingBoxMaxZ) / 2.0);

                XYZ center2 = cz2.SleevePlacementPoint ?? new XYZ(
                    (cz2.SleeveBoundingBoxMinX + cz2.SleeveBoundingBoxMaxX) / 2.0,
                    (cz2.SleeveBoundingBoxMinY + cz2.SleeveBoundingBoxMaxY) / 2.0,
                    (cz2.SleeveBoundingBoxMinZ + cz2.SleeveBoundingBoxMaxZ) / 2.0);

                if (shouldLog)
                {
                    bool hasPlacementPoint1 = cz1.SleevePlacementPoint != null;
                    bool hasPlacementPoint2 = cz2.SleevePlacementPoint != null;
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {sleeve1.SleeveInstanceId} Center: ({center1.X:F6}, {center1.Y:F6}, {center1.Z:F6}) [{(hasPlacementPoint1 ? "from PlacementPoint" : "from BBox center")}]\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve {sleeve2.SleeveInstanceId} Center: ({center2.X:F6}, {center2.Y:F6}, {center2.Z:F6}) [{(hasPlacementPoint2 ? "from PlacementPoint" : "from BBox center")}]\n");
                }

                // Calculate shared rotated axis direction (unit vector along rotated X-axis)
                XYZ rotatedAxisDirection = new XYZ(Math.Cos(rotationAngle), Math.Sin(rotationAngle), 0).Normalize();

                // Calculate vector from center1 to center2 in world-space
                XYZ worldVector = center2 - center1;

                // Project distance along the shared rotated axis using dot product
                double distanceAlongAxis = worldVector.DotProduct(rotatedAxisDirection);

                // Calculate perpendicular distance from axis using cross product magnitude
                XYZ crossProduct = worldVector.CrossProduct(rotatedAxisDirection);
                double perpendicularDistance = crossProduct.GetLength();

                // Get sleeve dimensions from database (rotated bounding boxes if available, else axis-aligned)
                double sleeve1Width = 0.0;
                double sleeve1Height = 0.0;
                double sleeve2Width = 0.0;
                double sleeve2Height = 0.0;

                if (cz1.RotatedBoundingBoxMinX.HasValue && cz1.RotatedBoundingBoxMaxX.HasValue &&
                    cz1.RotatedBoundingBoxMinY.HasValue && cz1.RotatedBoundingBoxMaxY.HasValue)
                {
                    sleeve1Width = cz1.RotatedBoundingBoxMaxX.Value - cz1.RotatedBoundingBoxMinX.Value;
                    sleeve1Height = cz1.RotatedBoundingBoxMaxY.Value - cz1.RotatedBoundingBoxMinY.Value;
                }
                else
                {
                    sleeve1Width = cz1.SleeveBoundingBoxMaxX - cz1.SleeveBoundingBoxMinX;
                    sleeve1Height = cz1.SleeveBoundingBoxMaxY - cz1.SleeveBoundingBoxMinY;
                }

                if (cz2.RotatedBoundingBoxMinX.HasValue && cz2.RotatedBoundingBoxMaxX.HasValue &&
                    cz2.RotatedBoundingBoxMinY.HasValue && cz2.RotatedBoundingBoxMaxY.HasValue)
                {
                    sleeve2Width = cz2.RotatedBoundingBoxMaxX.Value - cz2.RotatedBoundingBoxMinX.Value;
                    sleeve2Height = cz2.RotatedBoundingBoxMaxY.Value - cz2.RotatedBoundingBoxMinY.Value;
                }
                else
                {
                    sleeve2Width = cz2.SleeveBoundingBoxMaxX - cz2.SleeveBoundingBoxMinX;
                    sleeve2Height = cz2.SleeveBoundingBoxMaxY - cz2.SleeveBoundingBoxMinY;
                }

                // Calculate half-dimensions for overlap check
                double halfWidth1 = sleeve1Width / 2.0;
                double halfWidth2 = sleeve2Width / 2.0;
                double halfHeight1 = sleeve1Height / 2.0;
                double halfHeight2 = sleeve2Height / 2.0;

                if (shouldLog)
                {
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Rotated Axis Direction: ({rotatedAxisDirection.X:F6}, {rotatedAxisDirection.Y:F6}, {rotatedAxisDirection.Z:F6})\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] World Vector (center1 to center2): ({worldVector.X:F6}, {worldVector.Y:F6}, {worldVector.Z:F6})\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve 1 Size: W={sleeve1Width * 304.8:F1}mm, H={sleeve1Height * 304.8:F1}mm\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Sleeve 2 Size: W={sleeve2Width * 304.8:F1}mm, H={sleeve2Height * 304.8:F1}mm\n");
                }

                // Check if sleeves are close enough along the rotated axis
                double maxDistanceAlongAxis = halfWidth1 + halfWidth2 + toleranceDist;
                bool closeAlongAxis = Math.Abs(distanceAlongAxis) <= maxDistanceAlongAxis;

                // Check if perpendicular distance is small enough
                double maxPerpendicularDistance = halfHeight1 + halfHeight2 + toleranceDist;
                bool closePerpendicular = perpendicularDistance <= maxPerpendicularDistance;

                // Cluster if both conditions are met
                bool shouldCluster = closeAlongAxis && closePerpendicular;

                if (shouldLog)
                {
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Distance Along Axis: {distanceAlongAxis * 304.8:F1}mm (max allowed: {maxDistanceAlongAxis * 304.8:F1}mm) - {(closeAlongAxis ? "✅ PASS" : "❌ FAIL")}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Perpendicular Distance: {perpendicularDistance * 304.8:F1}mm (max allowed: {maxPerpendicularDistance * 304.8:F1}mm) - {(closePerpendicular ? "✅ PASS" : "❌ FAIL")}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] Final Result: {(shouldCluster ? "✅ CLUSTER" : "❌ NO CLUSTER")}\n");
                    System.IO.File.AppendAllText(clusterDebugLogPath, $"[{DateTime.Now:HH:mm:ss.fff}] ========== END CHECK ROTATED SLEEVE PROXIMITY ==========\n\n");
                }

                if (!DeploymentConfiguration.DeploymentMode && shouldCluster)
                {
                    DebugLogger.Info($"[RotatedClusterSleevePlacementService] Sleeves {sleeve1.SleeveInstanceId} and {sleeve2.SleeveInstanceId}: " +
                        $"Rotation={rotationAngle * 180.0 / Math.PI:F1}°, " +
                        $"DistanceAlongAxis={distanceAlongAxis * 304.8:F1}mm (max={maxDistanceAlongAxis * 304.8:F1}mm), " +
                        $"PerpendicularDistance={perpendicularDistance * 304.8:F1}mm (max={maxPerpendicularDistance * 304.8:F1}mm)");
                }

                return shouldCluster;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[RotatedClusterSleevePlacementService] Error checking rotated sleeve proximity: {ex.Message}");
                // Fallback to regular bounding box check
                return boundingBoxesOverlapFromXml(sleeve1, sleeve2, toleranceDist);
            }
        }

        /// <summary>
        /// Log cluster size application details
        /// </summary>
        private void LogClusterSizeApplication(FamilyInstance inst, double width, double height, double depth, double rotationAngle, SleeveGroupKey groupKey)
        {
            try
            {
                var clusterSizeLogPath = SafeFileLogger.GetLogFilePath("cluster_size_application.log");
                var sizeLogBuilder = new System.Text.StringBuilder();
                sizeLogBuilder.AppendLine($"[{DateTime.Now:HH:mm:ss.fff}] ========== CLUSTER SIZE APPLICATION ==========");
                sizeLogBuilder.AppendLine($"Cluster Sleeve ID: {inst.Id.IntegerValue}");
                sizeLogBuilder.AppendLine($"Calculated dimensions:");
                sizeLogBuilder.AppendLine($"  Width (internal): {width:F6} = {width * 304.8:F1}mm");
                sizeLogBuilder.AppendLine($"  Height (internal): {height:F6} = {height * 304.8:F1}mm");
                sizeLogBuilder.AppendLine($"  Depth (internal): {depth:F6} = {depth * 304.8:F1}mm");
                sizeLogBuilder.AppendLine($"Rotation Angle: {rotationAngle * 180.0 / Math.PI:F2}°");
                sizeLogBuilder.AppendLine($"Host Type: {groupKey.hostType}");
                sizeLogBuilder.AppendLine($"Orientation: {groupKey.orientation}");

                var widthParam = inst.LookupParameter("Width");
                var heightParam = inst.LookupParameter("Height");
                var depthParam = inst.LookupParameter("Depth");

                if (widthParam != null)
                {
                    double setWidth = widthParam.AsDouble();
                    sizeLogBuilder.AppendLine($"AFTER SETTING PARAMETERS:");
                    sizeLogBuilder.AppendLine($"  Width parameter: {setWidth:F6} = {setWidth * 304.8:F1}mm (expected: {width * 304.8:F1}mm)");
                }
                if (heightParam != null)
                {
                    double setHeight = heightParam.AsDouble();
                    sizeLogBuilder.AppendLine($"  Height parameter: {setHeight:F6} = {setHeight * 304.8:F1}mm (expected: {height * 304.8:F1}mm)");
                }
                if (depthParam != null)
                {
                    double setDepth = depthParam.AsDouble();
                    sizeLogBuilder.AppendLine($"  Depth parameter: {setDepth:F6} = {setDepth * 304.8:F1}mm (expected: {depth * 304.8:F1}mm)");
                }

                sizeLogBuilder.AppendLine($"  ========== END CLUSTER SIZE APPLICATION ==========");
                sizeLogBuilder.AppendLine();

                System.IO.File.AppendAllText(clusterSizeLogPath, sizeLogBuilder.ToString());
            }
            catch { }
        }
    }
}

