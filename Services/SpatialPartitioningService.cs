using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class SpatialPartitioningService
    {
        private readonly Dictionary<int, List<(Element, Transform?)>> _grid;
        private readonly BoundingBoxXYZ _modelBounds;
        private readonly double _gridSize;
        private readonly int _gridDivisionsX;
        private readonly int _gridDivisionsY;

        public SpatialPartitioningService(List<(Element, Transform?)> structuralElements, double gridSize = 10.0)
        {
            _grid = new Dictionary<int, List<(Element, Transform?)>>();
            _gridSize = gridSize;

            if (structuralElements == null || !structuralElements.Any())
            {
                _modelBounds = new BoundingBoxXYZ();
                return;
            }

            _modelBounds = CalculateModelBounds(structuralElements);

            _gridDivisionsX = (int)Math.Ceiling((_modelBounds.Max.X - _modelBounds.Min.X) / _gridSize);
            _gridDivisionsY = (int)Math.Ceiling((_modelBounds.Max.Y - _modelBounds.Min.Y) / _gridSize);

            foreach (var elementTuple in structuralElements)
            {
                var element = elementTuple.Item1;
                var transform = elementTuple.Item2;
                var boundingBox = element.get_BoundingBox(null);

                if (boundingBox != null)
                {
                    if (transform != null)
                    {
                        boundingBox = TransformBoundingBox(boundingBox, transform);
                    }

                    var minGridX = (int)Math.Floor((boundingBox.Min.X - _modelBounds.Min.X) / _gridSize);
                    var maxGridX = (int)Math.Floor((boundingBox.Max.X - _modelBounds.Min.X) / _gridSize);
                    var minGridY = (int)Math.Floor((boundingBox.Min.Y - _modelBounds.Min.Y) / _gridSize);
                    var maxGridY = (int)Math.Floor((boundingBox.Max.Y - _modelBounds.Min.Y) / _gridSize);

                    for (int x = minGridX; x <= maxGridX; x++)
                    {
                        for (int y = minGridY; y <= maxGridY; y++)
                        {
                            int index = y * _gridDivisionsX + x;
                            if (!_grid.ContainsKey(index))
                            {
                                _grid[index] = new List<(Element, Transform?)>();
                            }
                            _grid[index].Add(elementTuple);
                        }
                    }
                }
            }
        }

        public List<(Element, Transform?)> GetNearbyElements(Element mepElement)
        {
            var nearbyElements = new HashSet<(Element, Transform?)>();
            var boundingBox = mepElement.get_BoundingBox(null);

            if (boundingBox != null)
            {
                var minGridX = (int)Math.Floor((boundingBox.Min.X - _modelBounds.Min.X) / _gridSize);
                var maxGridX = (int)Math.Floor((boundingBox.Max.X - _modelBounds.Min.X) / _gridSize);
                var minGridY = (int)Math.Floor((boundingBox.Min.Y - _modelBounds.Min.Y) / _gridSize);
                var maxGridY = (int)Math.Floor((boundingBox.Max.Y - _modelBounds.Min.Y) / _gridSize);

                for (int x = minGridX; x <= maxGridX; x++)
                {
                    for (int y = minGridY; y <= maxGridY; y++)
                    {
                        int index = y * _gridDivisionsX + x;
                        if (_grid.ContainsKey(index))
                        {
                            foreach (var element in _grid[index])
                            {
                                nearbyElements.Add(element);
                            }
                        }
                    }
                }
            }

            return nearbyElements.ToList();
        }

        private BoundingBoxXYZ CalculateModelBounds(List<(Element, Transform?)> elements)
        {
            var minX = double.MaxValue;
            var minY = double.MaxValue;
            var minZ = double.MaxValue;
            var maxX = double.MinValue;
            var maxY = double.MinValue;
            var maxZ = double.MinValue;

            foreach (var elementTuple in elements)
            {
                var element = elementTuple.Item1;
                var transform = elementTuple.Item2;
                var boundingBox = element.get_BoundingBox(null);

                if (boundingBox != null)
                {
                    if (transform != null)
                    {
                        boundingBox = TransformBoundingBox(boundingBox, transform);
                    }

                    minX = Math.Min(minX, boundingBox.Min.X);
                    minY = Math.Min(minY, boundingBox.Min.Y);
                    minZ = Math.Min(minZ, boundingBox.Min.Z);
                    maxX = Math.Max(maxX, boundingBox.Max.X);
                    maxY = Math.Max(maxY, boundingBox.Max.Y);
                    maxZ = Math.Max(maxZ, boundingBox.Max.Z);
                }
            }

            return new BoundingBoxXYZ { Min = new XYZ(minX, minY, minZ), Max = new XYZ(maxX, maxY, maxZ) };
        }

        private BoundingBoxXYZ TransformBoundingBox(BoundingBoxXYZ bbox, Transform transform)
        {
            var min = transform.OfPoint(bbox.Min);
            var max = transform.OfPoint(bbox.Max);
            return new BoundingBoxXYZ
            {
                Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)),
                Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z))
            };
        }
    }
}
