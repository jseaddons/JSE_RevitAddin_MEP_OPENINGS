using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Generic spatial grid interface for efficient proximity queries.
    /// Provides O(1) insertion and O(k) proximity queries where k is the number of nearby items.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    /// <typeparam name="T">Type of items to index spatially</typeparam>
    public interface ISpatialGrid<T>
    {
        /// <summary>
        /// Build the spatial grid from a collection of items.
        /// </summary>
        /// <param name="items">Items to index</param>
        /// <param name="locationExtractor">Function to extract XYZ location from each item</param>
        void Build(IEnumerable<T> items, System.Func<T, XYZ> locationExtractor);
        
        /// <summary>
        /// Get items near a specific location within a radius.
        /// </summary>
        /// <param name="location">Center point for proximity search</param>
        /// <param name="radius">Search radius in Revit internal units (feet)</param>
        /// <returns>List of items within radius of location</returns>
        IList<T> GetNearbyItems(XYZ location, double radius);
        
        /// <summary>
        /// Get items within a bounding box.
        /// </summary>
        /// <param name="boundingBox">Bounding box for spatial query</param>
        /// <returns>List of items intersecting the bounding box</returns>
        IList<T> GetItemsInBounds(BoundingBoxXYZ boundingBox);
        
        /// <summary>
        /// Clear the grid and free resources.
        /// </summary>
        void Clear();
        
        /// <summary>
        /// Check if the grid has been built.
        /// </summary>
        bool IsBuilt { get; }
        
        /// <summary>
        /// Get the number of items currently indexed.
        /// </summary>
        int Count { get; }
        
        /// <summary>
        /// Get the grid cell size in Revit internal units (feet).
        /// </summary>
        double GridSize { get; }
    }
}
