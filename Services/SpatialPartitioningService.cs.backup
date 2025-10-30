using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Spatial partitioning service using 3D hash grid for efficient collision detection
    /// PHASE 2 OPTIMIZATION: Reduces O(n×m) brute-force intersection to O(n×log m)
    /// </summary>
    public class SpatialPartitioningService
    {
        private readonly double _gridSize;
        private readonly Dictionary<(int, int, int), List<(Element element, Transform? transform, BoundingBoxXYZ bbox)>> _grid;
        
        public SpatialPartitioningService(double gridSize = 1.0) // 1ft grid
        {
            _gridSize = gridSize;
            _grid = new Dictionary<(int, int, int), List<(Element, Transform?, BoundingBoxXYZ)>>();
        }
        
        /// <summary>
        /// Build spatial grid with structural elements
        /// </summary>
        public void BuildGrid(List<(Element element, Transform? transform, BoundingBoxXYZ bbox, Solid? solid)> structuralElements)
        {
            _grid.Clear();
            
            foreach (var (element, transform, bbox, solid) in structuralElements)
            {
                // Get all grid cells this element overlaps
                var cells = GetOverlappingCells(bbox);
                
                foreach (var cell in cells)
                {
                    if (!_grid.ContainsKey(cell))
                    {
                        _grid[cell] = new List<(Element, Transform?, BoundingBoxXYZ)>();
                    }
                    _grid[cell].Add((element, transform, bbox));
                }
            }
        }
        
        /// <summary>
        /// Get all grid cells overlapping with bounding box
        /// </summary>
        private List<(int, int, int)> GetOverlappingCells(BoundingBoxXYZ bbox)
        {
            var cells = new List<(int, int, int)>();
            
            int minX = (int)Math.Floor(bbox.Min.X / _gridSize);
            int maxX = (int)Math.Ceiling(bbox.Max.X / _gridSize);
            int minY = (int)Math.Floor(bbox.Min.Y / _gridSize);
            int maxY = (int)Math.Ceiling(bbox.Max.Y / _gridSize);
            int minZ = (int)Math.Floor(bbox.Min.Z / _gridSize);
            int maxZ = (int)Math.Ceiling(bbox.Max.Z / _gridSize);
            
            for (int x = minX; x <= maxX; x++)
            {
                for (int y = minY; y <= maxY; y++)
                {
                    for (int z = minZ; z <= maxZ; z++)
                    {
                        cells.Add((x, y, z));
                    }
                }
            }
            
            return cells;
        }
        
        /// <summary>
        /// Get nearby elements for a given bounding box
        /// </summary>
        public List<(Element element, Transform? transform, BoundingBoxXYZ bbox)> GetNearbyElements(BoundingBoxXYZ bbox)
        {
            var nearbyElements = new HashSet<(Element, Transform?, BoundingBoxXYZ)>();
            
            var cells = GetOverlappingCells(bbox);
            foreach (var cell in cells)
            {
                if (_grid.ContainsKey(cell))
                {
                    foreach (var elementData in _grid[cell])
                    {
                        nearbyElements.Add(elementData);
                    }
                }
            }
            
            return nearbyElements.ToList();
        }
        
        /// <summary>
        /// Get statistics about grid usage
        /// </summary>
        public (int totalCells, int usedCells, double avgElementsPerCell) GetStatistics()
        {
            var usedCells = _grid.Count;
            var totalElements = _grid.Values.Sum(cell => cell.Count);
            var avgElementsPerCell = usedCells > 0 ? (double)totalElements / usedCells : 0;
            
            return (_grid.Count, usedCells, avgElementsPerCell);
        }
        
        /// <summary>
        /// Clear the spatial grid
        /// </summary>
        public void Clear()
        {
            _grid.Clear();
        }
    }
}
