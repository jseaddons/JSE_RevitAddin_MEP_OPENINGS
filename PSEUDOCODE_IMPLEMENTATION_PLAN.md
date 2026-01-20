# Pseudocode Implementation Plan for Spatial Partitioning

This document provides a detailed pseudocode implementation plan for the `SpatialPartitioningService`.

## 1. `SpatialPartitioningService` Class

This service will be responsible for creating and managing a 3D grid for spatial partitioning of structural elements.

### Class Structure

```csharp
public class SpatialPartitioningService
{
    private readonly Dictionary<int, List<(Element, Transform?)>> _grid;
    private readonly BoundingBoxXYZ _modelBounds;
    private readonly double _gridSize;
    private readonly int _gridDivisionsX;
    private readonly int _gridDivisionsY;

    public SpatialPartitioningService(List<(Element, Transform?)> structuralElements, double gridSize = 10.0)
    {
        // Constructor logic
    }

    public List<(Element, Transform?)> GetNearbyElements(Element mepElement)
    {
        // Method to get nearby structural elements
    }

    private int GetGridIndex(XYZ point)
    {
        // Helper method to calculate grid index
    }
}
```

### Constructor

The constructor will initialize the grid and register all structural elements.

```csharp
public SpatialPartitioningService(List<(Element, Transform?)> structuralElements, double gridSize = 10.0)
{
    _grid = new Dictionary<int, List<(Element, Transform?)>>();
    _gridSize = gridSize;

    // 1. Calculate the total bounding box of all structural elements
    _modelBounds = CalculateModelBounds(structuralElements);

    // 2. Determine the number of grid divisions
    _gridDivisionsX = (int)Math.Ceiling((_modelBounds.Max.X - _modelBounds.Min.X) / _gridSize);
    _gridDivisionsY = (int)Math.Ceiling((_modelBounds.Max.Y - _modelBounds.Min.Y) / _gridSize);

    // 3. Register each structural element in the grid
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

            // Determine the grid cells the element's bounding box intersects with
            var minGridX = (int)Math.Floor((boundingBox.Min.X - _modelBounds.Min.X) / _gridSize);
            var maxGridX = (int)Math.Floor((boundingBox.Max.X - _modelBounds.Min.X) / _gridSize);
            var minGridY = (int)Math.Floor((boundingBox.Min.Y - _modelBounds.Min.Y) / _gridSize);
            var maxGridY = (int)Math.Floor((boundingBox.Max.Y - _modelBounds.Min.Y) / _gridSize);

            // Register the element in all intersected grid cells
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
```

### `GetNearbyElements` Method

This method will return a list of structural elements that are in the same grid cells as the given MEP element.

```csharp
public List<(Element, Transform?)> GetNearbyElements(Element mepElement)
{
    var nearbyElements = new HashSet<(Element, Transform?)>();
    var boundingBox = mepElement.get_BoundingBox(null);

    if (boundingBox != null)
    {
        // Determine the grid cells the MEP element's bounding box intersects with
        var minGridX = (int)Math.Floor((boundingBox.Min.X - _modelBounds.Min.X) / _gridSize);
        var maxGridX = (int)Math.Floor((boundingBox.Max.X - _modelBounds.Min.X) / _gridSize);
        var minGridY = (int)Math.Floor((boundingBox.Min.Y - _modelBounds.Min.Y) / _gridSize);
        var maxGridY = (int)Math.Floor((boundingBox.Max.Y - _modelBounds.Min.Y) / _gridSize);

        // Collect elements from all intersected grid cells
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
```

## 2. Integration with `DuctSleevePlacerService`

The `DuctSleevePlacerService` will be modified to use the `SpatialPartitioningService`.

### Updated `PlaceAllDuctSleeves` Method

```csharp
public void PlaceAllDuctSleeves()
{
    // 1. Initialize the SpatialPartitioningService
    var spatialService = new SpatialPartitioningService(_structuralElements);

    foreach (var tuple in _ductTuples)
    {
        var duct = tuple.Item1;
        // ... (existing code to get duct line)

        // 2. Get nearby structural elements from the spatial service
        var nearbyStructuralElements = spatialService.GetNearbyElements(duct);

        // 3. Find intersections only with the nearby elements
        var intersections = MepIntersectionService.FindIntersections(duct, nearbyStructuralElements, _log);

        // ... (existing code to place sleeves)
    }
}
```

This pseudocode provides a clear and detailed plan for implementing the spatial partitioning optimization.
