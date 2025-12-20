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
            
            var xRange = bbox.Max.X - bbox.Min.X;
            var yRange = bbox.Max.Y - bbox.Min.Y;
            var zRange = bbox.Max.Z - bbox.Min.Z;
            
            var orientation = GetHostOrientation();
            
            // LOGIC: Use Union Range for Width/Height (Face Dimensions).
            // FIX: Prioritize explicit DB WallThickness/FramingThickness/ClusterDepth from ANY sleeve.
            // Search for the first valid thickness (> 0) to use as the Source of Truth.
            
            double? dbThickness = null;
            
            // 1. Try to find valid thickness in Individual Sleeves first (most accurate)
            foreach (var s in Sleeves.Where(x => x.Type == SleeveType.Individual))
            {
                if (s.SourceData is JSE_RevitAddin_MEP_OPENINGS.Models.ClashZone cz)
                {
                    if (cz.WallThickness > 0.001) dbThickness = cz.WallThickness;
                    else if (cz.FramingThickness > 0.001) dbThickness = cz.FramingThickness;
                    else if (cz.StructuralElementThickness > 0.001) dbThickness = cz.StructuralElementThickness;
                    
                    if (dbThickness.HasValue) break;
                }
            }
            
            // 2. If no valid thickness found, check Cluster Sleeves
            if (!dbThickness.HasValue)
            {
                foreach (var s in Sleeves.Where(x => x.Type == SleeveType.Cluster))
                {
                    if (s.SourceData is JSE_RevitAddin_MEP_OPENINGS.Data.Repositories.ClusterSleeveData csd && csd.ClusterDepth > 0.001)
                    {
                        dbThickness = csd.ClusterDepth;
                        break;
                    }
                }
            }

            var referenceSleeve = Sleeves.FirstOrDefault(s => s.Type == SleeveType.Individual) ?? Sleeves.First();
            var refBbox = referenceSleeve.BoundingBox;

            double width, height, depth;
            
            if (string.Equals(orientation, "Y", StringComparison.OrdinalIgnoreCase))
            {
                // Y-Wall
                width = yRange;  // Length along wall (Union)
                height = zRange; // Vertical height (Union)
                depth = dbThickness ?? (refBbox.Max.X - refBbox.Min.X); // DB Thickness or BBox X-dim
            }
            else if (string.Equals(orientation, "Z", StringComparison.OrdinalIgnoreCase) || 
                     string.Equals(GetHostType(), "Floor", StringComparison.OrdinalIgnoreCase))
            {
                // Floor
                width = xRange;
                height = yRange;
                depth = dbThickness ?? (refBbox.Max.Z - refBbox.Min.Z); // DB Thickness or BBox Z-dim
            }
            else
            {
                // Default / X-Wall
                width = xRange;  // Length along wall
                height = zRange; // Vertical height
                depth = dbThickness ?? (refBbox.Max.Y - refBbox.Min.Y); // DB Thickness or BBox Y-dim
            }
            
            return (width, height, depth);
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
               // Use rotation of first sleeve if available
               var first = Sleeves.FirstOrDefault();
               if (first != null && Math.Abs(first.RotationAngleDeg) > 0.1)
               {
                   return first.RotationAngleDeg * (Math.PI / 180.0);
               }
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
