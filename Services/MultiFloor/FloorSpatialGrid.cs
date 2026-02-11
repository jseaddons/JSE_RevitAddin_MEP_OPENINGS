using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    /// <summary>
    /// Spatial grid for O(1) proximity searches instead of O(N²)
    /// </summary>
    public class FloorSpatialGrid
    {
        private readonly Dictionary<(int x, int y, int z), List<Element>> _grid;
        private readonly double _cellSize;
        
        public FloorSpatialGrid(List<Element> elements, List<Element> hosts, double cellSize = 5.0)
        {
            _cellSize = cellSize;
            _grid = new Dictionary<(int x, int y, int z), List<Element>>();
            
            // Build grid
            foreach (var elem in elements.Concat(hosts))
            {
                var bbox = elem.get_BoundingBox(null);
                if (bbox == null) continue;
                
                var cells = GetCellsForBoundingBox(bbox);
                foreach (var cell in cells)
                {
                    if (!_grid.ContainsKey(cell))
                        _grid[cell] = new List<Element>();
                    
                    _grid[cell].Add(elem);
                }
            }
        }
        
        public List<Element> GetCandidates(Element element)
        {
            var bbox = element.get_BoundingBox(null);
            if (bbox == null) return new List<Element>();
            
            var candidates = new HashSet<Element>();
            var cells = GetCellsForBoundingBox(bbox);
            
            foreach (var cell in cells)
            {
                if (_grid.TryGetValue(cell, out var cellElements))
                {
                    foreach (var e in cellElements)
                    {
                        if (e.Id != element.Id) // Don't include self
                            candidates.Add(e);
                    }
                }
            }
            
            return candidates.ToList();
        }
        
        private List<(int x, int y, int z)> GetCellsForBoundingBox(BoundingBoxXYZ bbox)
        {
            var cells = new List<(int x, int y, int z)>();
            
            int minX = (int)Math.Floor(bbox.Min.X / _cellSize);
            int maxX = (int)Math.Floor(bbox.Max.X / _cellSize);
            int minY = (int)Math.Floor(bbox.Min.Y / _cellSize);
            int maxY = (int)Math.Floor(bbox.Max.Y / _cellSize);
            int minZ = (int)Math.Floor(bbox.Min.Z / _cellSize);
            int maxZ = (int)Math.Floor(bbox.Max.Z / _cellSize);
            
            for (int x = minX; x <= maxX; x++)
                for (int y = minY; y <= maxY; y++)
                    for (int z = minZ; z <= maxZ; z++)
                        cells.Add((x, y, z));
            
            return cells;
        }
    }
}
