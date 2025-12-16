using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

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
            var bbox = CalculateCombinedBoundingBox();
            
            var width = bbox.Max.X - bbox.Min.X;
            var height = bbox.Max.Y - bbox.Min.Y;
            var depth = bbox.Max.Z - bbox.Min.Z;
            
            return (width, height, depth);
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
