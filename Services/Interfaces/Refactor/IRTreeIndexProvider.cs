using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Abstraction for R-Tree spatial index operations on clash zones.
    /// Provides efficient spatial queries for large datasets.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    public interface IRTreeIndexProvider
    {
        /// <summary>
        /// Build R-Tree index from clash zones for spatial queries.
        /// </summary>
        /// <param name="zones">Clash zones to index</param>
        void Build(IList<ClashZone> zones);
        
        /// <summary>
        /// Query zones within a bounding box using the R-Tree index.
        /// </summary>
        /// <param name="box">Bounding box for spatial query</param>
        /// <returns>List of clash zones intersecting the bounding box</returns>
        IList<ClashZone> Query(BoundingBoxXYZ box);
        
        /// <summary>
        /// Query zones near a specific point within a radius.
        /// </summary>
        /// <param name="point">Center point for query</param>
        /// <param name="radius">Search radius in Revit internal units (feet)</param>
        /// <returns>List of clash zones within radius of point</returns>
        IList<ClashZone> QueryNearby(XYZ point, double radius);
        
        /// <summary>
        /// Clear the index and free resources.
        /// </summary>
        void Clear();
        
        /// <summary>
        /// Check if the index has been built and is ready for queries.
        /// </summary>
        bool IsIndexed { get; }
        
        /// <summary>
        /// Get the number of zones currently indexed.
        /// </summary>
        int IndexedCount { get; }
    }
}
