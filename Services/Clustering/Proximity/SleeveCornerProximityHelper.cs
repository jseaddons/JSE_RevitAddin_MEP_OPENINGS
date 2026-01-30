using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Clustering.Proximity
{
    /// <summary>
    /// Helper for proximity calculations using SleeveCorner data from ClashZone database.
    /// Uses multithreading and batch processing for performance.
    /// 
    /// ✅ USE THIS FOR ALL RECTANGULAR SLEEVES (All categories):
    /// - Ducts (Rectangular)
    /// - Pipes (when placed as rectangular)
    /// - Cable Trays
    /// - Conduits
    /// - Dampers
    /// - Any sleeve with rectangular geometry
    /// 
    /// Sleeve corners provide accurate placement data when Revit bounding boxes are erratic.
    /// </summary>
    public class SleeveCornerProximityHelper
    {
        private const int DEFAULT_BATCH_SIZE = 100;
        private readonly int _batchSize;

        public SleeveCornerProximityHelper(int batchSize = DEFAULT_BATCH_SIZE)
        {
            _batchSize = batchSize;
        }
        
        /// <summary>
        /// Determine if a ClashZone represents a rectangular sleeve.
        /// Returns true for all rectangular sleeves regardless of category.
        /// </summary>
        public bool IsRectangularSleeve(ClashZone cz)
        {
            if (cz == null) return false;
            
            // ✅ Check 1: Has valid sleeve corners (primary indicator)
            if (HasValidSleeveCorners(cz))
                return true;
            
            // ✅ Check 2: Explicit rectangular duct shapes
            if (!string.IsNullOrEmpty(cz.DuctShape))
            {
                string shape = cz.DuctShape.ToLower();
                if (shape.Contains("rectangular") || shape.Contains("rect"))
                    return true;
                
                // Round/Circular are NOT rectangular
                if (shape.Contains("round") || shape.Contains("circular") || shape.Contains("oval"))
                    return false;
            }
            
            // ✅ Check 3: Category-based heuristics
            string category = cz.MepElementCategory?.ToLower() ?? "";
            
            // Dampers are always rectangular
            if (category.Contains("damper"))
                return true;
            
            // Cable trays are always rectangular
            if (category.Contains("cable tray"))
                return true;
            
            // Conduits can be rectangular
            if (category.Contains("conduit"))
            {
                // If it has valid corners, it's rectangular
                return HasValidSleeveCorners(cz);
            }
            
            // Pipes - check if rectangular (unusual but possible)
            if (category.Contains("pipe"))
            {
                // Pipes with valid corners are rectangular sleeves
                return HasValidSleeveCorners(cz);
            }
            
            // Ducts - check shape or corners
            if (category.Contains("duct"))
            {
                // If we have corners, it's rectangular
                return HasValidSleeveCorners(cz);
            }
            
            // Default: If has valid corners, treat as rectangular
            return HasValidSleeveCorners(cz);
        }
        
        /// <summary>
        /// Determine if a ClashZone represents a circular/round sleeve.
        /// </summary>
        public bool IsCircularSleeve(ClashZone cz)
        {
            if (cz == null) return false;
            
            // ✅ Check explicit round shapes
            if (!string.IsNullOrEmpty(cz.DuctShape))
            {
                string shape = cz.DuctShape.ToLower();
                if (shape.Contains("round") || shape.Contains("circular") || shape.Contains("oval"))
                    return true;
            }
            
            // ✅ Check if has diameter (circular indicator)
            if (cz.SleeveDiameter > 0 && !HasValidSleeveCorners(cz))
                return true;
            
            // ✅ Pipes without corners are circular
            string category = cz.MepElementCategory?.ToLower() ?? "";
            if (category.Contains("pipe") && cz.SleeveDiameter > 0)
                return true;
            
            return false;
        }

        /// <summary>
        /// Calculate bounding box from sleeve corners.
        /// </summary>
        public (double minX, double maxX, double minY, double maxY, double minZ, double maxZ) GetBoundingBoxFromCorners(ClashZone cz)
        {
            if (!HasValidSleeveCorners(cz))
            {
                throw new ArgumentException($"ClashZone {cz.Id} does not have valid sleeve corners");
            }

            var minX = Math.Min(Math.Min(cz.SleeveCorner1X ?? 0, cz.SleeveCorner2X ?? 0),
                               Math.Min(cz.SleeveCorner3X ?? 0, cz.SleeveCorner4X ?? 0));
            var maxX = Math.Max(Math.Max(cz.SleeveCorner1X ?? 0, cz.SleeveCorner2X ?? 0),
                               Math.Max(cz.SleeveCorner3X ?? 0, cz.SleeveCorner4X ?? 0));

            var minY = Math.Min(Math.Min(cz.SleeveCorner1Y ?? 0, cz.SleeveCorner2Y ?? 0),
                               Math.Min(cz.SleeveCorner3Y ?? 0, cz.SleeveCorner4Y ?? 0));
            var maxY = Math.Max(Math.Max(cz.SleeveCorner1Y ?? 0, cz.SleeveCorner2Y ?? 0),
                               Math.Max(cz.SleeveCorner3Y ?? 0, cz.SleeveCorner4Y ?? 0));

            var minZ = Math.Min(Math.Min(cz.SleeveCorner1Z ?? 0, cz.SleeveCorner2Z ?? 0),
                               Math.Min(cz.SleeveCorner3Z ?? 0, cz.SleeveCorner4Z ?? 0));
            var maxZ = Math.Max(Math.Max(cz.SleeveCorner1Z ?? 0, cz.SleeveCorner2Z ?? 0),
                               Math.Max(cz.SleeveCorner3Z ?? 0, cz.SleeveCorner4Z ?? 0));

            // ✅ SAFETY CHECK: Ensure minimum thickness (e.g. 1mm ~ 0.0033 ft) to prevent 2D boxes
            // This ensures overlaps are detected even if two faces are mathematically co-planar
            const double minThickness = 0.0033; // ~1mm
            if (maxX - minX < minThickness) { minX -= minThickness/2; maxX += minThickness/2; }
            if (maxY - minY < minThickness) { minY -= minThickness/2; maxY += minThickness/2; }
            if (maxZ - minZ < minThickness) { minZ -= minThickness/2; maxZ += minThickness/2; }

            return (minX, maxX, minY, maxY, minZ, maxZ);
        }

        /// <summary>
        /// Check if two ClashZones are within proximity using their sleeve corners.
        /// For wall-hosted sleeves, ignores the wall depth axis:
        /// - X-wall: Check Y and Z overlap (ignore X)
        /// - Y-wall: Check X and Z overlap (ignore Y)
        /// </summary>
        public bool AreWithinProximity(ClashZone cz1, ClashZone cz2, double toleranceDist)
        {
            if (!HasValidSleeveCorners(cz1) || !HasValidSleeveCorners(cz2))
            {
                return false;
            }

            try
            {
                var bbox1 = GetBoundingBoxFromCorners(cz1);
                var bbox2 = GetBoundingBoxFromCorners(cz2);

                // ✅ FIX: Use ceiling on tolerance to be more forgiving at boundary values
                // If user sets 100mm tolerance, distances up to 100.999mm will be treated as 100mm
                // This handles both floating-point precision AND slight measurement variations
                // Example: 100.9mm distance with 100mm tolerance → Math.Ceiling(100.9) = 101mm tolerance → clusters ✅
                var expandedTolerance = Math.Ceiling(toleranceDist * 304.8) / 304.8; // Convert to mm, ceiling, back to feet

                // Expand bbox1 by tolerance
                var min1X = bbox1.minX - expandedTolerance;
                var max1X = bbox1.maxX + expandedTolerance;
                var min1Y = bbox1.minY - expandedTolerance;
                var max1Y = bbox1.maxY + expandedTolerance;
                var min1Z = bbox1.minZ - expandedTolerance;
                var max1Z = bbox1.maxZ + expandedTolerance;

                // Check overlap
                bool overlapX = bbox2.maxX >= min1X && bbox2.minX <= max1X;
                bool overlapY = bbox2.maxY >= min1Y && bbox2.minY <= max1Y;
                bool overlapZ = bbox2.maxZ >= min1Z && bbox2.minZ <= max1Z;

                // ✅ DIAGNOSTIC LOGGING: Show why proximity check fails
                string category = cz1.MepElementCategory ?? "";
                if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                {
                    double toleranceMM = toleranceDist * 304.8;
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"[ProximityCheck] 🔍 DETAILED CHECK for {category}:\n" +
                        $"  Tolerance: {toleranceMM:F1}mm ({toleranceDist:F6}ft)\n" +
                        $"  Sleeve1 BBox: X=[{bbox1.minX:F3}, {bbox1.maxX:F3}], Y=[{bbox1.minY:F3}, {bbox1.maxY:F3}], Z=[{bbox1.minZ:F3}, {bbox1.maxZ:F3}]\n" +
                        $"  Sleeve2 BBox: X=[{bbox2.minX:F3}, {bbox2.maxX:F3}], Y=[{bbox2.minY:F3}, {bbox2.maxY:F3}], Z=[{bbox2.minZ:F3}, {bbox2.maxZ:F3}]\n" +
                        $"  Expanded1:    X=[{min1X:F3}, {max1X:F3}], Y=[{min1Y:F3}, {max1Y:F3}], Z=[{min1Z:F3}, {max1Z:F3}]\n" +
                        $"  Overlap: X={overlapX}, Y={overlapY}, Z={overlapZ}\n");
                }

                // ✅ FIX: For wall-hosted sleeves, ignore the wall depth axis
                // X-wall (wall running along X): ignore Y (depth), check X and Z
                // Y-wall (wall running along Y): ignore X (depth), check Y and Z
                string hostOrientation1 = cz1.HostOrientation ?? "";
                string hostOrientation2 = cz2.HostOrientation ?? "";
                
                // Both must be on same wall orientation to cluster
                if (hostOrientation1 == hostOrientation2 && !string.IsNullOrEmpty(hostOrientation1))
                {
                    if (hostOrientation1.Equals("X", StringComparison.OrdinalIgnoreCase))
                    {
                        // X-wall: ignore Y, check X and Z overlap
                        bool result = overlapX && overlapZ;
                        if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"  HostOrientation: X-wall → Check X && Z: {overlapX} && {overlapZ} = {result}\n");
                        }
                        return result;
                    }
                    else if (hostOrientation1.Equals("Y", StringComparison.OrdinalIgnoreCase))
                    {
                        // Y-wall: ignore X, check Y and Z overlap
                        bool result = overlapY && overlapZ;
                        if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                        {
                            SafeFileLogger.SafeAppendText("cluster_debug.log",
                                $"  HostOrientation: Y-wall → Check Y && Z: {overlapY} && {overlapZ} = {result}\n");
                        }
                        return result;
                    }
                }
                
                // Default: require all 3 axes to overlap
                bool finalResult = overlapX && overlapY && overlapZ;
                if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                {
                    SafeFileLogger.SafeAppendText("cluster_debug.log",
                        $"  Default (all axes): {overlapX} && {overlapY} && {overlapZ} = {finalResult}\n");
                }
                return finalResult;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Check if ClashZone has valid sleeve corners (not all zeros).
        /// </summary>
        public bool HasValidSleeveCorners(ClashZone cz)
        {
            if (cz == null) return false;

            // Check if at least one corner has non-zero values
            return (cz.SleeveCorner1X != 0 || cz.SleeveCorner1Y != 0 || cz.SleeveCorner1Z != 0) ||
                   (cz.SleeveCorner2X != 0 || cz.SleeveCorner2Y != 0 || cz.SleeveCorner2Z != 0) ||
                   (cz.SleeveCorner3X != 0 || cz.SleeveCorner3Y != 0 || cz.SleeveCorner3Z != 0) ||
                   (cz.SleeveCorner4X != 0 || cz.SleeveCorner4Y != 0 || cz.SleeveCorner4Z != 0);
        }

        /// <summary>
        /// Batch process proximity checks for a list of ClashZones with multithreading.
        /// Returns a dictionary: ClashZone.Id -> List of nearby ClashZone IDs
        /// </summary>
        public Dictionary<Guid, List<Guid>> FindNearbyClashZonesBatch(
            List<ClashZone> clashZones,
            double toleranceDist,
            bool enableParallel = true)
        {
            var results = new ConcurrentDictionary<Guid, List<Guid>>();

            if (clashZones == null || clashZones.Count == 0)
                return results.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            // Filter to only zones with valid corners
            var validZones = clashZones.Where(cz => HasValidSleeveCorners(cz)).ToList();

            if (validZones.Count == 0)
                return results.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            // Split into batches
            var batches = SplitIntoBatches(validZones, _batchSize);

            if (enableParallel)
            {
                System.Threading.Tasks.Parallel.ForEach(batches, batch =>
                {
                    ProcessBatch(batch, validZones, toleranceDist, results);
                });
            }
            else
            {
                foreach (var batch in batches)
                {
                    ProcessBatch(batch, validZones, toleranceDist, results);
                }
            }

            return results.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        /// <summary>
        /// Find all ClashZones within proximity for a specific ClashZone.
        /// </summary>
        public List<Guid> FindNearbyClashZones(
            ClashZone targetZone,
            List<ClashZone> candidateZones,
            double toleranceDist)
        {
            var nearby = new List<Guid>();

            if (!HasValidSleeveCorners(targetZone))
                return nearby;

            foreach (var candidate in candidateZones)
            {
                if (candidate.Id == targetZone.Id)
                    continue;

                if (AreWithinProximity(targetZone, candidate, toleranceDist))
                {
                    nearby.Add(candidate.Id);
                }
            }

            return nearby;
        }

        /// <summary>
        /// Calculate center point from sleeve corners.
        /// </summary>
        public (double x, double y, double z) GetCenterFromCorners(ClashZone cz)
        {
            if (!HasValidSleeveCorners(cz))
            {
                throw new ArgumentException($"ClashZone {cz.Id} does not have valid sleeve corners");
            }

            var bbox = GetBoundingBoxFromCorners(cz);
            return (
                (bbox.minX + bbox.maxX) / 2.0,
                (bbox.minY + bbox.maxY) / 2.0,
                (bbox.minZ + bbox.maxZ) / 2.0
            );
        }

        /// <summary>
        /// Calculate distance between two ClashZones using their corners.
        /// </summary>
        public double GetDistanceBetweenCorners(ClashZone cz1, ClashZone cz2)
        {
            if (!HasValidSleeveCorners(cz1) || !HasValidSleeveCorners(cz2))
            {
                return double.MaxValue;
            }

            var center1 = GetCenterFromCorners(cz1);
            var center2 = GetCenterFromCorners(cz2);

            var dx = center2.x - center1.x;
            var dy = center2.y - center1.y;
            var dz = center2.z - center1.z;

            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        /// <summary>
        /// Get dimensions (width, height, depth) from sleeve corners.
        /// </summary>
        public (double width, double height, double depth) GetDimensionsFromCorners(ClashZone cz)
        {
            if (!HasValidSleeveCorners(cz))
            {
                throw new ArgumentException($"ClashZone {cz.Id} does not have valid sleeve corners");
            }

            var bbox = GetBoundingBoxFromCorners(cz);
            return (
                bbox.maxX - bbox.minX,
                bbox.maxY - bbox.minY,
                bbox.maxZ - bbox.minZ
            );
        }

        #region Private Helper Methods

        private void ProcessBatch(
            List<ClashZone> batch,
            List<ClashZone> allZones,
            double toleranceDist,
            ConcurrentDictionary<Guid, List<Guid>> results)
        {
            foreach (var zone in batch)
            {
                var nearby = FindNearbyClashZones(zone, allZones, toleranceDist);
                results[zone.Id] = nearby;
            }
        }

        private List<List<ClashZone>> SplitIntoBatches(List<ClashZone> zones, int batchSize)
        {
            var batches = new List<List<ClashZone>>();
            for (int i = 0; i < zones.Count; i += batchSize)
            {
                var batch = zones.Skip(i).Take(batchSize).ToList();
                batches.Add(batch);
            }
            return batches;
        }

        #endregion
    }

    /// <summary>
    /// Result class for proximity analysis.
    /// </summary>
    public class ProximityAnalysisResult
    {
        public Guid ClashZoneId { get; set; }
        public List<Guid> NearbyClashZoneIds { get; set; } = new List<Guid>();
        public int NearbyCount => NearbyClashZoneIds.Count;
        public bool HasNearby => NearbyCount > 0;
    }

    /// <summary>
    /// Builder for fluent API usage.
    /// </summary>
    public class SleeveCornerProximityBuilder
    {
        private List<ClashZone> _clashZones = new List<ClashZone>();
        private double _toleranceDist = 0.0;
        private int _batchSize = 100;
        private bool _enableParallel = true;

        public SleeveCornerProximityBuilder WithClashZones(List<ClashZone> zones)
        {
            _clashZones = zones;
            return this;
        }

        public SleeveCornerProximityBuilder WithTolerance(double tolerance)
        {
            _toleranceDist = tolerance;
            return this;
        }

        public SleeveCornerProximityBuilder WithBatchSize(int batchSize)
        {
            _batchSize = batchSize;
            return this;
        }

        public SleeveCornerProximityBuilder WithParallel(bool enable)
        {
            _enableParallel = enable;
            return this;
        }

        public Dictionary<Guid, List<Guid>> Build()
        {
            var helper = new SleeveCornerProximityHelper(_batchSize);
            return helper.FindNearbyClashZonesBatch(_clashZones, _toleranceDist, _enableParallel);
        }

        public List<ProximityAnalysisResult> BuildDetailed()
        {
            var helper = new SleeveCornerProximityHelper(_batchSize);
            var results = helper.FindNearbyClashZonesBatch(_clashZones, _toleranceDist, _enableParallel);

            return results.Select(kvp => new ProximityAnalysisResult
            {
                ClashZoneId = kvp.Key,
                NearbyClashZoneIds = kvp.Value
            }).ToList();
        }
    }
}
