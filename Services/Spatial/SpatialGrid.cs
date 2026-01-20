using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Spatial
{
    /// <summary>
    /// Generic spatial grid implementation for efficient proximity queries.
    /// Divides 3D space into uniform grid cells for O(1) insertion and O(k) proximity queries.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    /// <typeparam name="T">Type of items to index spatially</typeparam>
    public class SpatialGrid<T> : ISpatialGrid<T>
    {
        private readonly Dictionary<int, List<T>> _grid;
        private BoundingBoxXYZ _modelBounds;
        private readonly double _gridSize;
        private int _gridDivisionsX;
        private int _gridDivisionsY;
        private int _gridDivisionsZ;
        private Func<T, XYZ> _locationExtractor;
        
        public bool IsBuilt { get; private set; }
        public int Count { get; private set; }
        public double GridSize => _gridSize;

        /// <summary>
        /// Constructor with configurable grid size.
        /// </summary>
        /// <param name="gridSize">Grid cell size in Revit internal units (feet). Default: 10 feet.</param>
        public SpatialGrid(double gridSize = 10.0)
        {
            _grid = new Dictionary<int, List<T>>();
            _gridSize = gridSize;
            IsBuilt = false;
            Count = 0;
        }

        /// <summary>
        /// Build the spatial grid from a collection of items.
        /// </summary>
        public void Build(IEnumerable<T> items, Func<T, XYZ> locationExtractor)
        {
            if (items == null || !items.Any())
            {
                _modelBounds = new BoundingBoxXYZ();
                IsBuilt = false;
                Count = 0;
                return;
            }

            _locationExtractor = locationExtractor ?? throw new ArgumentNullException(nameof(locationExtractor));
            
            // Clear existing grid
            _grid.Clear();
            
            // Calculate model bounds
            _modelBounds = CalculateModelBounds(items);
            
            // Calculate grid divisions
            _gridDivisionsX = (int)Math.Ceiling((_modelBounds.Max.X - _modelBounds.Min.X) / _gridSize);
            _gridDivisionsY = (int)Math.Ceiling((_modelBounds.Max.Y - _modelBounds.Min.Y) / _gridSize);
            _gridDivisionsZ = (int)Math.Ceiling((_modelBounds.Max.Z - _modelBounds.Min.Z) / _gridSize);
            
            // Populate grid
            Count = 0;
            foreach (var item in items)
            {
                var location = _locationExtractor(item);
                if (location != null)
                {
                    int gridIndex = GetGridIndex(location);
                    
                    if (!_grid.ContainsKey(gridIndex))
                    {
                        _grid[gridIndex] = new List<T>();
                    }
                    _grid[gridIndex].Add(item);
                    Count++;
                }
            }
            
            IsBuilt = true;
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"[SpatialGrid] Built grid with {Count} items, " +
                    $"divisions: {_gridDivisionsX}x{_gridDivisionsY}x{_gridDivisionsZ}, " +
                    $"cell size: {_gridSize:F2} ft");
            }
        }

        /// <summary>
        /// Get items near a specific location within a radius.
        /// </summary>
        public IList<T> GetNearbyItems(XYZ location, double radius)
        {
            if (!IsBuilt || location == null)
                return new List<T>();

            var nearbyItems = new List<T>();
            
            // Calculate grid cell range to search
            var minGridX = (int)Math.Floor((location.X - radius - _modelBounds.Min.X) / _gridSize);
            var maxGridX = (int)Math.Ceiling((location.X + radius - _modelBounds.Min.X) / _gridSize);
            var minGridY = (int)Math.Floor((location.Y - radius - _modelBounds.Min.Y) / _gridSize);
            var maxGridY = (int)Math.Ceiling((location.Y + radius - _modelBounds.Min.Y) / _gridSize);
            var minGridZ = (int)Math.Floor((location.Z - radius - _modelBounds.Min.Z) / _gridSize);
            var maxGridZ = (int)Math.Ceiling((location.Z + radius - _modelBounds.Min.Z) / _gridSize);

            // Search grid cells
            for (int x = Math.Max(0, minGridX); x <= Math.Min(_gridDivisionsX - 1, maxGridX); x++)
            {
                for (int y = Math.Max(0, minGridY); y <= Math.Min(_gridDivisionsY - 1, maxGridY); y++)
                {
                    for (int z = Math.Max(0, minGridZ); z <= Math.Min(_gridDivisionsZ - 1, maxGridZ); z++)
                    {
                        int index = GetGridIndex(x, y, z);
                        if (_grid.ContainsKey(index))
                        {
                            // Filter by actual distance
                            foreach (var item in _grid[index])
                            {
                                var itemLocation = _locationExtractor(item);
                                if (itemLocation != null && itemLocation.DistanceTo(location) <= radius)
                                {
                                    nearbyItems.Add(item);
                                }
                            }
                        }
                    }
                }
            }

            return nearbyItems;
        }

        /// <summary>
        /// Get items within a bounding box.
        /// </summary>
        public IList<T> GetItemsInBounds(BoundingBoxXYZ boundingBox)
        {
            if (!IsBuilt || boundingBox == null)
                return new List<T>();

            var itemsInBounds = new List<T>();
            
            // Calculate grid cell range
            var minGridX = (int)Math.Floor((boundingBox.Min.X - _modelBounds.Min.X) / _gridSize);
            var maxGridX = (int)Math.Ceiling((boundingBox.Max.X - _modelBounds.Min.X) / _gridSize);
            var minGridY = (int)Math.Floor((boundingBox.Min.Y - _modelBounds.Min.Y) / _gridSize);
            var maxGridY = (int)Math.Ceiling((boundingBox.Max.Y - _modelBounds.Min.Y) / _gridSize);
            var minGridZ = (int)Math.Floor((boundingBox.Min.Z - _modelBounds.Min.Z) / _gridSize);
            var maxGridZ = (int)Math.Ceiling((boundingBox.Max.Z - _modelBounds.Min.Z) / _gridSize);

            // Search grid cells
            for (int x = Math.Max(0, minGridX); x <= Math.Min(_gridDivisionsX - 1, maxGridX); x++)
            {
                for (int y = Math.Max(0, minGridY); y <= Math.Min(_gridDivisionsY - 1, maxGridY); y++)
                {
                    for (int z = Math.Max(0, minGridZ); z <= Math.Min(_gridDivisionsZ - 1, maxGridZ); z++)
                    {
                        int index = GetGridIndex(x, y, z);
                        if (_grid.ContainsKey(index))
                        {
                            // Filter by actual bounds intersection
                            foreach (var item in _grid[index])
                            {
                                var itemLocation = _locationExtractor(item);
                                if (itemLocation != null && IsInBounds(itemLocation, boundingBox))
                                {
                                    itemsInBounds.Add(item);
                                }
                            }
                        }
                    }
                }
            }

            return itemsInBounds;
        }

        /// <summary>
        /// Clear the grid and free resources.
        /// </summary>
        public void Clear()
        {
            _grid.Clear();
            IsBuilt = false;
            Count = 0;
            _locationExtractor = null;
        }

        // Helper methods

        private BoundingBoxXYZ CalculateModelBounds(IEnumerable<T> items)
        {
            var minX = double.MaxValue;
            var minY = double.MaxValue;
            var minZ = double.MaxValue;
            var maxX = double.MinValue;
            var maxY = double.MinValue;
            var maxZ = double.MinValue;

            foreach (var item in items)
            {
                var location = _locationExtractor(item);
                if (location != null)
                {
                    minX = Math.Min(minX, location.X);
                    minY = Math.Min(minY, location.Y);
                    minZ = Math.Min(minZ, location.Z);
                    maxX = Math.Max(maxX, location.X);
                    maxY = Math.Max(maxY, location.Y);
                    maxZ = Math.Max(maxZ, location.Z);
                }
            }

            return new BoundingBoxXYZ 
            { 
                Min = new XYZ(minX, minY, minZ), 
                Max = new XYZ(maxX, maxY, maxZ) 
            };
        }

        private int GetGridIndex(XYZ location)
        {
            int gridX = (int)Math.Floor((location.X - _modelBounds.Min.X) / _gridSize);
            int gridY = (int)Math.Floor((location.Y - _modelBounds.Min.Y) / _gridSize);
            int gridZ = (int)Math.Floor((location.Z - _modelBounds.Min.Z) / _gridSize);
            return GetGridIndex(gridX, gridY, gridZ);
        }

        private int GetGridIndex(int x, int y, int z)
        {
            // 3D to 1D index mapping
            return z * (_gridDivisionsX * _gridDivisionsY) + y * _gridDivisionsX + x;
        }

        private bool IsInBounds(XYZ point, BoundingBoxXYZ box)
        {
            return point.X >= box.Min.X && point.X <= box.Max.X &&
                   point.Y >= box.Min.Y && point.Y <= box.Max.Y &&
                   point.Z >= box.Min.Z && point.Z <= box.Max.Z;
        }
    }
}
