using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Data.Repositories;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Models
{
    /// <summary>
    /// Represents a group of sleeves that are in proximity to each other
    /// and should be combined into a single combined sleeve.
    /// 
    /// A proximity group contains sleeves from different categories
    /// (e.g., Pipes + Ducts) that are within the proximity threshold.
    /// </summary>
    public class ProximityGroup
    {
        /// <summary>
        /// List of sleeves in this proximity group
        /// </summary>
        public List<UnifiedSleeve> Sleeves { get; set; } = new List<UnifiedSleeve>();
        
        /// <summary>
        /// Gets the distinct categories represented in this group
        /// </summary>
        public List<string> Categories => Sleeves
            .Select(s => s.Category)
            .Distinct()
            .OrderBy(c => c)
            .ToList();
        
        /// <summary>
        /// Gets the count of sleeves in this group
        /// </summary>
        public int Count => Sleeves.Count;
        
        /// <summary>
        /// Adds a sleeve to this proximity group
        /// </summary>
        public void Add(UnifiedSleeve sleeve)
        {
            if (sleeve == null)
                throw new ArgumentNullException(nameof(sleeve));
            
            Sleeves.Add(sleeve);
        }
        
        /// <summary>
        /// Adds multiple sleeves to this proximity group
        /// </summary>
        public void AddRange(IEnumerable<UnifiedSleeve> sleeves)
        {
            if (sleeves == null)
                throw new ArgumentNullException(nameof(sleeves));
            
            Sleeves.AddRange(sleeves);
        }
        
        /// <summary>
        /// Calculates the combined bounding box that encompasses all sleeves in the group
        /// </summary>
        public BoundingBoxXYZ CalculateCombinedBoundingBox()
        {
            if (Sleeves.Count == 0)
                throw new InvalidOperationException("Cannot calculate bounding box for empty proximity group");
            
            // Collect all corner points from all sleeves
            var allCorners = Sleeves
                .Where(s => s.Corners != null && s.Corners.Count > 0)
                .SelectMany(s => s.Corners)
                .ToList();
            
            if (allCorners.Count == 0)
            {
                // Fallback: use bounding boxes if corners not available
                allCorners = Sleeves
                    .Where(s => s.BoundingBox != null)
                    .SelectMany(s => new[] { s.BoundingBox.Min, s.BoundingBox.Max })
                    .ToList();
            }
            
            if (allCorners.Count == 0)
                throw new InvalidOperationException("No geometry data available for bounding box calculation");
            
            var minX = allCorners.Min(c => c.X);
            var minY = allCorners.Min(c => c.Y);
            var minZ = allCorners.Min(c => c.Z);
            var maxX = allCorners.Max(c => c.X);
            var maxY = allCorners.Max(c => c.Y);
            var maxZ = allCorners.Max(c => c.Z);
            
            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ),
                Max = new XYZ(maxX, maxY, maxZ)
            };
        }
        
        /// <summary>
        /// Calculates the combined dimensions (width, height, depth) for the group
        /// </summary>
        public (double width, double height, double depth) CalculateCombinedDimensions()
        {
            // ✅ LIGHTWEIGHT OBB LOGIC (Use Rotation Vector Projection)
            // Projects corners onto the wall's local axes defined by RotationAngle.
            // Matches Cluster Sleeve logic: Efficiently handles rotated walls (45 deg) without 'Too Big' AABB.

            var referenceSleeve = Sleeves.First();
            double rotationRad = referenceSleeve.RotationAngleDeg * (Math.PI / 180.0);
            string orientation = GetHostOrientation();
            string hostType = GetHostType();

            // 1. Define Projection Vectors (Rotated Axes)
            XYZ vecU = new XYZ(Math.Cos(rotationRad), Math.Sin(rotationRad), 0); // "Rotated X" / Along Rotated X-Wall
            XYZ vecV = new XYZ(-Math.Sin(rotationRad), Math.Cos(rotationRad), 0); // "Rotated Y" / Normal to Rotated X-Wall (or Along Rotated Y-Wall)
            XYZ vecZ = XYZ.BasisZ;

            // 2. Project all corners
            var allCorners = Sleeves
                .Where(s => s.Corners != null && s.Corners.Count > 0)
                .SelectMany(s => s.Corners)
                .ToList();

            if (allCorners.Count == 0) // Fallback
            {
                 allCorners = Sleeves.Where(s => s.BoundingBox != null)
                     .SelectMany(s => new[] { 
                         s.BoundingBox.Min, 
                         s.BoundingBox.Max,
                         new XYZ(s.BoundingBox.Max.X, s.BoundingBox.Min.Y, s.BoundingBox.Min.Z),
                         new XYZ(s.BoundingBox.Min.X, s.BoundingBox.Max.Y, s.BoundingBox.Min.Z)
                     }).ToList();
            }

            double minU = double.MaxValue, maxU = double.MinValue;
            double minV = double.MaxValue, maxV = double.MinValue;
            double minZ = double.MaxValue, maxZ = double.MinValue;

            foreach (var pt in allCorners)
            {
                double u = pt.DotProduct(vecU);
                double v = pt.DotProduct(vecV);
                double z = pt.DotProduct(vecZ);

                if (u < minU) minU = u; if (u > maxU) maxU = u;
                if (v < minV) minV = v; if (v > maxV) maxV = v;
                if (z < minZ) minZ = z; if (z > maxZ) maxZ = z;
            }

            // 3. Determine Dimensions based on Host Type & Orientation
            double rangeU = maxU - minU; // Rotated X-Range
            double rangeV = maxV - minV; // Rotated Y-Range
            double rangeZ = maxZ - minZ; // Z-Range (Vertical)

            double calculatedWidth, calculatedHeight, calculatedDepth;

            // Database Thickness Check (Source of Truth)
            double? dbThickness = null;
            foreach (var s in Sleeves)
            {
                if (s.SourceData is ClashZone cz)
                {
                    if (cz.WallThickness > 0.001) dbThickness = cz.WallThickness;
                    else if (cz.FramingThickness > 0.001) dbThickness = cz.FramingThickness;
                    if (dbThickness.HasValue) break;
                }
                else if (s.SourceData is ClusterSleeveData csd && csd.ClusterDepth > 0.001)
                {
                    dbThickness = csd.ClusterDepth;
                    break;
                }
            }

            if (string.Equals(hostType, "Floor", StringComparison.OrdinalIgnoreCase))
            {
                // Floor: Planar dimensions are U/V
                // Width -> Range U (Rotated X)
                // Height -> Range V (Rotated Y) - Note: Height param usually maps to "Depth" in plan
                // Depth -> Range Z (Thickness)
                calculatedWidth = rangeU;
                calculatedHeight = rangeV; 
                calculatedDepth = dbThickness ?? rangeZ;
            }
            else
            {
                // Wall: Use Orientation to swap U/V for Width vs Depth
                // Z is always Height
                calculatedHeight = rangeZ;

                if (string.Equals(orientation, "Y", StringComparison.OrdinalIgnoreCase))
                {
                    // Y-Wall (North-South): Width is along Y-axis (vecV)
                    calculatedWidth = rangeV;
                    calculatedDepth = dbThickness ?? rangeU; // Depth is Thickness (X-axis / vecU)
                }
                else
                {
                    // X-Wall (East-West) or Default: Width is along X-axis (vecU)
                    calculatedWidth = rangeU;
                    calculatedDepth = dbThickness ?? rangeV; // Depth is Thickness (Y-axis / vecV)
                }
            }

            return (calculatedWidth, calculatedHeight, calculatedDepth);
        }
        
        /// <summary>
        /// Calculates the rotation angle in radians for the combined sleeve
        /// </summary>
        public double CalculateCombinedRotation()
        {
            var orientation = GetHostOrientation();
            
            // USER FEEDBACK: Family is 'Left Oriented' / Y-Aligned by default.
            // Y-Wall (running North-South): Rotation 0 (Native alignment)
            // X-Wall (running East-West): Rotation 90 deg (PI/2) to align Y-Width to X-Wall
            
            if (string.Equals(orientation, "Y", StringComparison.OrdinalIgnoreCase))
            {
                return 0.0; // Native alignment for Y-wall
            }
            
            // For Floors or slanted walls, we might want to respect the sleeve's rotation
            if (string.Equals(GetHostType(), "Floor", StringComparison.OrdinalIgnoreCase))
            {
               // ✅ USER REQUEST (2026-02-05): "for floor no rotation needed"
               // Even if constituent sleeves are rotated (e.g. 45 deg duct), the combined opening should be axis-aligned (0 deg).
               // The bounding box calculation already accounts for the rotated geometry (AABB), so the hole matches the full extent.
               return 0.0;
            }
            
            // Default / X-Wall -> Rotate 90 degrees
            return Math.PI / 2.0;
        }

        /// <summary>
        /// Calculates the combined placement point (center of combined bounding box)
        /// </summary>
        public XYZ CalculateCombinedPlacementPoint()
        {
            var bbox = CalculateCombinedBoundingBox();
            return (bbox.Min + bbox.Max) / 2.0;
        }
        
        /// <summary>
        /// Gets the host type from the first sleeve in the group
        /// (assumes all sleeves in proximity have the same host)
        /// </summary>
        public string GetHostType()
        {
            return Sleeves.FirstOrDefault()?.HostType;
        }
        
        /// <summary>
        /// Gets the host orientation from the first sleeve in the group
        /// (assumes all sleeves in proximity have the same host orientation)
        /// </summary>
        public string GetHostOrientation()
        {
            return Sleeves.FirstOrDefault()?.HostOrientation;
        }
        
        /// <summary>
        /// Checks if this group contains sleeves from multiple categories (cross-category)
        /// </summary>
        public bool IsCrossCategory()
        {
            return Categories.Count > 1;
        }
        
        /// <summary>
        /// Gets a summary string describing this proximity group
        /// </summary>
        public string GetSummary()
        {
            var categorySummary = string.Join("+", Categories);
            var typeSummary = Sleeves.GroupBy(s => s.Type)
                .Select(g => $"{g.Count()} {g.Key}")
                .ToList();
            
            return $"{categorySummary} ({string.Join(", ", typeSummary)})";
        }
        
        /// <summary>
        /// Validates that this proximity group is valid for combined sleeve creation
        /// </summary>
        public bool IsValid()
        {
            // Must have at least 2 sleeves
            if (Sleeves.Count < 2)
                return false;
            
            // Must be cross-category (different categories)
            if (!IsCrossCategory())
                return false;
            
            // All sleeves must have geometry data
            if (Sleeves.Any(s => s.BoundingBox == null && (s.Corners == null || s.Corners.Count == 0)))
                return false;
            
            return true;
        }
    }
}
