using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Combined.Spatial
{
    /// <summary>
    /// Simplified R-Tree spatial index for fast proximity queries.
    /// Provides O(log n) query time vs O(n²) brute force comparison.
    /// 
    /// This is a simplified implementation optimized for the combined sleeve use case.
    /// For production, consider using a full R-Tree library like RBush or STRtree.
    /// </summary>
    public class SimplifiedSpatialIndex
    {
        private class SpatialEntry
        {
            public string Id { get; set; }
            public XYZ Min { get; set; }
            public XYZ Max { get; set; }
            public XYZ Center { get; set; }
        }
        
        private readonly List<SpatialEntry> _entries = new List<SpatialEntry>();
        private readonly double _gridSize = 10.0; // feet - grid cell size for spatial hashing
        private Dictionary<string, List<SpatialEntry>> _grid;
        
        /// <summary>
        /// Inserts a bounding box into the spatial index
        /// </summary>
        public void Insert(string id, XYZ min, XYZ max)
        {
            if (string.IsNullOrEmpty(id))
                throw new ArgumentNullException(nameof(id));
            if (min == null)
                throw new ArgumentNullException(nameof(min));
            if (max == null)
                throw new ArgumentNullException(nameof(max));
            
            var entry = new SpatialEntry
            {
                Id = id,
                Min = min,
                Max = max,
                Center = (min + max) / 2.0
            };
            
            _entries.Add(entry);
        }
        
        /// <summary>
        /// Builds the spatial index (call after all inserts)
        /// Uses spatial hashing for fast lookups
        /// </summary>
        public void Build()
        {
            _grid = new Dictionary<string, List<SpatialEntry>>();
            
            foreach (var entry in _entries)
            {
                // Get grid cells that this entry overlaps
                var gridCells = GetOverlappingGridCells(entry.Min, entry.Max);
                
                foreach (var cellKey in gridCells)
                {
                    if (!_grid.ContainsKey(cellKey))
                        _grid[cellKey] = new List<SpatialEntry>();
                    
                    _grid[cellKey].Add(entry);
                }
            }
        }
        
        /// <summary>
        /// Queries the spatial index for entries within a search box
        /// </summary>
        public List<string> Query(XYZ searchMin, XYZ searchMax)
        {
            if (searchMin == null)
                throw new ArgumentNullException(nameof(searchMin));
            if (searchMax == null)
                throw new ArgumentNullException(nameof(searchMax));
            
            if (_grid == null)
                Build(); // Auto-build if not built yet
            
            var results = new HashSet<string>();
            
            // Get grid cells that the search box overlaps
            var gridCells = GetOverlappingGridCells(searchMin, searchMax);
            
            foreach (var cellKey in gridCells)
            {
                if (_grid.TryGetValue(cellKey, out var entries))
                {
                    foreach (var entry in entries)
                    {
                        // Check if entry's bounding box intersects search box
                        if (BoundingBoxesIntersect(entry.Min, entry.Max, searchMin, searchMax))
                        {
                            results.Add(entry.Id);
                        }
                    }
                }
            }
            
            return results.ToList();
        }
        
        /// <summary>
        /// Queries for entries within a radius of a center point
        /// </summary>
        public List<string> QueryRadius(XYZ center, double radius)
        {
            if (center == null)
                throw new ArgumentNullException(nameof(center));
            
            var searchMin = center - new XYZ(radius, radius, radius);
            var searchMax = center + new XYZ(radius, radius, radius);
            
            var candidates = Query(searchMin, searchMax);
            
            // Filter by actual distance
            var results = new List<string>();
            foreach (var id in candidates)
            {
                var entry = _entries.FirstOrDefault(e => e.Id == id);
                if (entry != null)
                {
                    var distance = center.DistanceTo(entry.Center);
                    if (distance <= radius)
                    {
                        results.Add(id);
                    }
                }
            }
            
            return results;
        }
        
        /// <summary>
        /// Gets all grid cell keys that a bounding box overlaps
        /// </summary>
        private List<string> GetOverlappingGridCells(XYZ min, XYZ max)
        {
            var cells = new List<string>();
            
            var minCellX = (int)Math.Floor(min.X / _gridSize);
            var minCellY = (int)Math.Floor(min.Y / _gridSize);
            var minCellZ = (int)Math.Floor(min.Z / _gridSize);
            
            var maxCellX = (int)Math.Floor(max.X / _gridSize);
            var maxCellY = (int)Math.Floor(max.Y / _gridSize);
            var maxCellZ = (int)Math.Floor(max.Z / _gridSize);
            
            for (int x = minCellX; x <= maxCellX; x++)
            {
                for (int y = minCellY; y <= maxCellY; y++)
                {
                    for (int z = minCellZ; z <= maxCellZ; z++)
                    {
                        cells.Add($"{x},{y},{z}");
                    }
                }
            }
            
            return cells;
        }
        
        /// <summary>
        /// Checks if two bounding boxes intersect
        /// </summary>
        private bool BoundingBoxesIntersect(XYZ min1, XYZ max1, XYZ min2, XYZ max2)
        {
            return min1.X <= max2.X && max1.X >= min2.X &&
                   min1.Y <= max2.Y && max1.Y >= min2.Y &&
                   min1.Z <= max2.Z && max1.Z >= min2.Z;
        }
        
        /// <summary>
        /// Gets the total number of entries in the index
        /// </summary>
        public int Count => _entries.Count;
        
        /// <summary>
        /// Clears all entries from the index
        /// </summary>
        public void Clear()
        {
            _entries.Clear();
            _grid?.Clear();
            _grid = null;
        }
    }
}
