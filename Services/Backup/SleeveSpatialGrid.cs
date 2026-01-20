using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class SleeveSpatialGrid
    {
        private readonly Dictionary<int, List<FamilyInstance>> _grid;
        private readonly BoundingBoxXYZ _modelBounds;
        private readonly double _gridSize;
        private readonly int _gridDivisionsX;
        private readonly int _gridDivisionsY;

        public SleeveSpatialGrid(List<FamilyInstance> sleeves, double gridSize = 10.0)
        {
            _grid = new Dictionary<int, List<FamilyInstance>>();
            _gridSize = gridSize;

            if (sleeves == null || !sleeves.Any())
            {
                _modelBounds = new BoundingBoxXYZ();
                return;
            }

            _modelBounds = CalculateModelBounds(sleeves);

            _gridDivisionsX = (int)Math.Ceiling((_modelBounds.Max.X - _modelBounds.Min.X) / _gridSize);
            _gridDivisionsY = (int)Math.Ceiling((_modelBounds.Max.Y - _modelBounds.Min.Y) / _gridSize);

            foreach (var sleeve in sleeves)
            {
                var location = (sleeve.Location as LocationPoint)?.Point;
                if (location != null)
                {
                    int gridX = (int)Math.Floor((location.X - _modelBounds.Min.X) / _gridSize);
                    int gridY = (int)Math.Floor((location.Y - _modelBounds.Min.Y) / _gridSize);
                    int index = gridY * _gridDivisionsX + gridX;

                    if (!_grid.ContainsKey(index))
                    {
                        _grid[index] = new List<FamilyInstance>();
                    }
                    _grid[index].Add(sleeve);
                }
            }
        }

        public List<FamilyInstance> GetNearbySleeves(XYZ location, double radius)
        {
            var nearbySleeves = new List<FamilyInstance>();
            var minGridX = (int)Math.Floor((location.X - radius - _modelBounds.Min.X) / _gridSize);
            var maxGridX = (int)Math.Floor((location.X + radius - _modelBounds.Min.X) / _gridSize);
            var minGridY = (int)Math.Floor((location.Y - radius - _modelBounds.Min.Y) / _gridSize);
            var maxGridY = (int)Math.Floor((location.Y + radius - _modelBounds.Min.Y) / _gridSize);

            for (int x = minGridX; x <= maxGridX; x++)
            {
                for (int y = minGridY; y <= maxGridY; y++)
                {
                    int index = y * _gridDivisionsX + x;
                    if (_grid.ContainsKey(index))
                    {
                        nearbySleeves.AddRange(_grid[index]);
                    }
                }
            }

            return nearbySleeves;
        }

        private BoundingBoxXYZ CalculateModelBounds(List<FamilyInstance> sleeves)
        {
            var minX = double.MaxValue;
            var minY = double.MaxValue;
            var minZ = double.MaxValue;
            var maxX = double.MinValue;
            var maxY = double.MinValue;
            var maxZ = double.MinValue;

            foreach (var sleeve in sleeves)
            {
                var location = (sleeve.Location as LocationPoint)?.Point;
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

            return new BoundingBoxXYZ { Min = new XYZ(minX, minY, minZ), Max = new XYZ(maxX, maxY, maxZ) };
        }
    }
}
