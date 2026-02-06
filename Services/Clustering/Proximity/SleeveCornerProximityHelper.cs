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
        /// 
        /// OLD BEHAVIOUR (axis overlap):
        /// - Expanded one bbox by tolerance and required axis-wise overlap (e.g. Y && Z for Y‑wall).
        /// - This meant a sleeve could pass if |ΔY| &lt; tol and |ΔZ| &lt; tol even when the true
        ///   diagonal distance &gt; tol.
        /// 
        /// NEW BEHAVIOUR (user-requested):
        /// - Use the **actual Euclidean edge‑to‑edge distance** between the 2D rectangles:
        ///   - X‑wall: distance in X–Z plane
        ///   - Y‑wall: distance in Y–Z plane
        /// - Compare that distance directly with the tolerance.
        /// - This prevents “diagonal chaining” where a visually far sleeve is clustered only
        ///   because each axis delta is just under the threshold.
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

                // ✅ FIX: Use ceiling on tolerance to be slightly forgiving at boundary values.
                // We still apply this, but now against the *Euclidean* distance rather than
                // axis-wise overlaps.
                var effectiveTolerance = Math.Ceiling(toleranceDist * 304.8) / 304.8; // feet

                // Helper to compute 1D gap between two intervals (0 if overlapping)
                double Gap(double min1, double max1, double min2, double max2)
                {
                    if (max1 < min2) return min2 - max1;
                    if (max2 < min1) return min1 - max2;
                    return 0.0;
                }

                // ✅ PERFORMANCE: Disabled diagnostic logging (causes O(N²) file I/O during clustering)
                // string category = cz1.MepElementCategory ?? "";
                // if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                // {
                //     double toleranceMM = toleranceDist * 304.8;
                //     SafeFileLogger.SafeAppendText("cluster_debug.log",
                //         $"[ProximityCheck] 🔍 DETAILED CHECK for {category}:\n" +
                //         $"  Tolerance: {toleranceMM:F1}mm ({toleranceDist:F6}ft)\n" +
                //         $"  Sleeve1 BBox: X=[{bbox1.minX:F3}, {bbox1.maxX:F3}], Y=[{bbox1.minY:F3}, {bbox1.maxY:F3}], Z=[{bbox1.minZ:F3}, {bbox1.maxZ:F3}]\n" +
                //         $"  Sleeve2 BBox: X=[{bbox2.minX:F3}, {bbox2.maxX:F3}], Y=[{bbox2.minY:F3}, {bbox2.maxY:F3}], Z=[{bbox2.minZ:F3}, {bbox2.maxZ:F3}]\n");
                // }
                string category = cz1.MepElementCategory ?? "";

                // ✅ FIX: For wall-hosted sleeves, compute TRUE 2D edge‑to‑edge distance in the wall plane.
                // X‑wall (wall running along X): distance in X–Z, ignore Y (depth)
                // Y‑wall (wall running along Y): distance in Y–Z, ignore X (depth)
                string hostOrientation1 = cz1.HostOrientation ?? "";
                string hostOrientation2 = cz2.HostOrientation ?? "";
                
                // Both must be on same wall orientation to cluster
                if (hostOrientation1 == hostOrientation2 && !string.IsNullOrEmpty(hostOrientation1))
                {
                    if (hostOrientation1.Equals("X", StringComparison.OrdinalIgnoreCase))
                    {
                        // X‑wall: distance in X–Z plane
                        double gapX = Gap(bbox1.minX, bbox1.maxX, bbox2.minX, bbox2.maxX);
                        double gapZ = Gap(bbox1.minZ, bbox1.maxZ, bbox2.minZ, bbox2.maxZ);
                        double dist = Math.Sqrt(gapX * gapX + gapZ * gapZ);
                        bool result = dist <= effectiveTolerance;

                        // if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                        // {
                        //     SafeFileLogger.SafeAppendText("cluster_debug.log",
                        //         $"  HostOrientation: X-wall → dXZ={dist * 304.8:F1}mm (tol={effectiveTolerance * 304.8:F1}mm) => {result}\n");
                        // }
                        return result;
                    }
                    else if (hostOrientation1.Equals("Y", StringComparison.OrdinalIgnoreCase))
                    {
                        // Y‑wall: distance in Y–Z plane
                        double gapY = Gap(bbox1.minY, bbox1.maxY, bbox2.minY, bbox2.maxY);
                        double gapZ = Gap(bbox1.minZ, bbox1.maxZ, bbox2.minZ, bbox2.maxZ);
                        double dist = Math.Sqrt(gapY * gapY + gapZ * gapZ);
                        bool result = dist <= effectiveTolerance;

                        // if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                        // {
                        //     SafeFileLogger.SafeAppendText("cluster_debug.log",
                        //         $"  HostOrientation: Y-wall → dYZ={dist * 304.8:F1}mm (tol={effectiveTolerance * 304.8:F1}mm) => {result}\n");
                        // }
                        return result;
                    }
                }

                // ✅ FIX: For floors, compute 2D distance in X-Y plane (ignore Z gap)
                // This fix addresses the issue where sleeves in floors are separated vertically (Z) but should cluster based on plan view proximity.
                string hostType1 = cz1.StructuralElementType ?? "";
                string hostType2 = cz2.StructuralElementType ?? "";
                if (hostType1.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    hostType2.IndexOf("Floor", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    double gapX2D = Gap(bbox1.minX, bbox1.maxX, bbox2.minX, bbox2.maxX);
                    double gapY2D = Gap(bbox1.minY, bbox1.maxY, bbox2.minY, bbox2.maxY);
                    double dist2D = Math.Sqrt(gapX2D * gapX2D + gapY2D * gapY2D);
                    bool floorResult = dist2D <= effectiveTolerance;

                    // if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                    // {
                    //     SafeFileLogger.SafeAppendText("cluster_debug.log",
                    //         $"  HostType: Floor → dXY(2D)={dist2D * 304.8:F1}mm (tol={effectiveTolerance * 304.8:F1}mm) => {floorResult}\n");
                    // }
                    return floorResult;
                }
                
                // Default: compute full 3D distance between the two corner bounding boxes
                double gapX3 = Gap(bbox1.minX, bbox1.maxX, bbox2.minX, bbox2.maxX);
                double gapY3 = Gap(bbox1.minY, bbox1.maxY, bbox2.minY, bbox2.maxY);
                double gapZ3 = Gap(bbox1.minZ, bbox1.maxZ, bbox2.minZ, bbox2.maxZ);
                double dist3D = Math.Sqrt(gapX3 * gapX3 + gapY3 * gapY3 + gapZ3 * gapZ3);
                bool finalResult = dist3D <= effectiveTolerance;

                // if (category.Contains("Duct") || category.Contains("Damper") || category.Contains("Tray"))
                // {
                //     SafeFileLogger.SafeAppendText("cluster_debug.log",
                //         $"  Default 3D distance={dist3D * 304.8:F1}mm (tol={effectiveTolerance * 304.8:F1}mm) => {finalResult}\n");
                // }
                return finalResult;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Sleeve shape only (no MEP element shape): rectangular opening = has corner geometry from Revit.
        /// Use this for proximity checker selection.
        /// </summary>
        public bool HasRectangularSleeveShape(ClashZone cz)
        {
            if (cz == null) return false;

            // If this sleeve is circular (has a diameter), it should NOT be treated as rectangular
            // even if corners exist in the DB (from extraction). Circular sleeves must go through
            // edge-to-edge logic only.
            bool isCircular = (cz.SleeveDiameter > 0) || (cz.CalculatedSleeveDiameter > 0);
            if (isCircular) return false;

            // Non-circular + has corners → rectangular opening
            return HasValidSleeveCorners(cz);
        }

        /// <summary>
        /// Sleeve shape only (no MEP element shape): circular opening = has diameter and no corner geometry.
        /// Use this for proximity checker selection (edge-to-edge).
        /// </summary>
        public bool HasCircularSleeveShape(ClashZone cz)
        {
            if (cz == null) return false;

            // NEW RULE: Circular detection is based purely on diameter.
            // If it has a diameter, it is circular and MUST go to edge-to-edge,
            // regardless of whether corners were extracted.
            return (cz.SleeveDiameter > 0) || (cz.CalculatedSleeveDiameter > 0);
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
