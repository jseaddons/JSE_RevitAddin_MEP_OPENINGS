using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Unified service for zone pre-filtering using spatial optimizations.
    /// Composes spatial grid and R-tree index for efficient zone filtering.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    public interface IZonePreFilterService
    {
        /// <summary>
        /// Filter zones based on spatial criteria and configuration.
        /// </summary>
        /// <param name="zones">Input zones to filter</param>
        /// <param name="criteria">Filtering criteria including spatial bounds and predicates</param>
        /// <returns>Filtered list of zones matching criteria</returns>
        IList<ClashZone> Filter(IList<ClashZone> zones, FilterCriteria criteria);
        
        /// <summary>
        /// Get zones near a specific location using spatial optimizations.
        /// </summary>
        /// <param name="location">Center point for proximity search</param>
        /// <param name="radius">Search radius in Revit internal units (feet)</param>
        /// <returns>List of zones within radius of location</returns>
        IList<ClashZone> GetNearbyZones(XYZ location, double radius);
        
        /// <summary>
        /// Get zones within a bounding box.
        /// </summary>
        /// <param name="boundingBox">Bounding box for spatial query</param>
        /// <returns>List of zones intersecting the bounding box</returns>
        IList<ClashZone> GetZonesInBounds(BoundingBoxXYZ boundingBox);
    }
    
    /// <summary>
    /// Criteria for zone filtering operations.
    /// Immutable configuration for thread-safe filtering.
    /// </summary>
    public class FilterCriteria
    {
        /// <summary>
        /// Whether to use R-Tree index for spatial queries (faster for large datasets).
        /// </summary>
        public bool UseRTreeIndex { get; set; } = true;
        
        /// <summary>
        /// Whether to use spatial grid for proximity queries (faster for localized searches).
        /// </summary>
        public bool UseSpatialGrid { get; set; } = true;
        
        /// <summary>
        /// Optional section box to limit spatial queries.
        /// </summary>
        public BoundingBoxXYZ SectionBox { get; set; }
        
        /// <summary>
        /// Optional custom predicate for additional filtering logic.
        /// </summary>
        public Func<ClashZone, bool> CustomPredicate { get; set; }
        
        /// <summary>
        /// Minimum zone priority for filtering (zones below this priority are excluded).
        /// </summary>
        public int? MinimumPriority { get; set; }
        
        /// <summary>
        /// Whether to exclude zones that already have sleeves placed.
        /// </summary>
        public bool ExcludePlacedZones { get; set; } = false;
    }
}
