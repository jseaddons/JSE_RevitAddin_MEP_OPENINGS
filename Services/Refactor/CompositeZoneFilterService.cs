using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;  // For IZoneFilterService
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Spatial;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactor
{
    /// <summary>
    /// Composite zone filtering service using spatial grid and R-tree optimizations.
    /// Part of Team C - Optimization Layer Isolation.
    /// </summary>
    public class CompositeZoneFilterService : IZonePreFilterService
    {
        private readonly IRTreeIndexProvider _rTreeIndex;
        private readonly ISpatialGrid<ClashZone> _spatialGrid;
        private readonly IZoneFilterService _baseFilterService;
        
        /// <summary>
        /// Constructor with dependency injection.
        /// </summary>
        /// <param name="rTreeIndex">R-Tree index provider for spatial queries</param>
        /// <param name="spatialGrid">Spatial grid for proximity queries</param>
        /// <param name="baseFilterService">Base zone filter service (existing implementation)</param>
        public CompositeZoneFilterService(
            IRTreeIndexProvider rTreeIndex = null,
            ISpatialGrid<ClashZone> spatialGrid = null,
            IZoneFilterService baseFilterService = null)
        {
            _rTreeIndex = rTreeIndex ?? new ClashZoneRTreeIndex();
            _spatialGrid = spatialGrid ?? new SpatialGrid<ClashZone>(gridSize: 10.0); // 10 feet grid cells
            _baseFilterService = baseFilterService; // May be null - will use direct filtering
        }
        
        /// <summary>
        /// Filter zones based on spatial criteria and configuration.
        /// </summary>
        public IList<ClashZone> Filter(IList<ClashZone> zones, FilterCriteria criteria)
        {
            if (zones == null || zones.Count == 0)
                return new List<ClashZone>();
            
            IList<ClashZone> filtered = zones;
            
            // Build spatial indexes if needed
            if (criteria.UseSpatialGrid && !_spatialGrid.IsBuilt)
            {
                _spatialGrid.Build(zones, z => z.IntersectionPoint);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CompositeZoneFilter] Built spatial grid with {_spatialGrid.Count} zones");
                }
            }
            
            if (criteria.UseRTreeIndex && !_rTreeIndex.IsIndexed)
            {
                _rTreeIndex.Build(zones);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[CompositeZoneFilter] Built R-Tree index with {_rTreeIndex.IndexedCount} zones");
                }
            }
            
            // Apply section box filtering if specified
            if (criteria.SectionBox != null)
            {
                if (criteria.UseRTreeIndex && _rTreeIndex.IsIndexed)
                {
                    // Use R-Tree for efficient spatial query
                    filtered = _rTreeIndex.Query(criteria.SectionBox);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CompositeZoneFilter] R-Tree filtered: {zones.Count} → {filtered.Count} zones");
                    }
                }
                else if (criteria.UseSpatialGrid && _spatialGrid.IsBuilt)
                {
                    // Fallback to spatial grid
                    filtered = _spatialGrid.GetItemsInBounds(criteria.SectionBox);
                    
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Info($"[CompositeZoneFilter] Spatial grid filtered: {zones.Count} → {filtered.Count} zones");
                    }
                }
                else
                {
                    // Fallback to linear filtering
                    filtered = zones.Where(z => IsInBounds(z, criteria.SectionBox)).ToList();
                }
            }
            
            // Apply custom predicate if specified
            if (criteria.CustomPredicate != null)
            {
                filtered = filtered.Where(criteria.CustomPredicate).ToList();
            }
            
            // Apply minimum priority filter
            // TODO: ClashZone doesn't have Priority property - skip for now or map to ReadyForPlacement
            if (criteria.MinimumPriority.HasValue)
            {
                // Placeholder: Priority not implemented in ClashZone model
                // Could map: Priority 0 = not ready, Priority 1+ = ReadyForPlacement
                filtered = filtered.Where(z => z.ReadyForPlacement).ToList();
            }
            
            // Exclude placed zones if requested
            if (criteria.ExcludePlacedZones)
            {
                // ClashZone has SleeveInstanceId - if > 0, sleeve has been placed
                filtered = filtered.Where(z => z.SleeveInstanceId <= 0).ToList();
            }
            
            return filtered;
        }
        
        /// <summary>
        /// Get zones near a specific location using spatial optimizations.
        /// </summary>
        public IList<ClashZone> GetNearbyZones(XYZ location, double radius)
        {
            if (location == null)
                return new List<ClashZone>();
            
            // Prefer spatial grid for proximity queries (faster for localized searches)
            if (_spatialGrid.IsBuilt)
            {
                return _spatialGrid.GetNearbyItems(location, radius);
            }
            
            // Fallback to R-Tree if available
            if (_rTreeIndex.IsIndexed)
            {
                return _rTreeIndex.QueryNearby(location, radius);
            }
            
            // No spatial index available
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning("[CompositeZoneFilter] No spatial index built for proximity query");
            }
            return new List<ClashZone>();
        }
        
        /// <summary>
        /// Get zones within a bounding box.
        /// </summary>
        public IList<ClashZone> GetZonesInBounds(BoundingBoxXYZ boundingBox)
        {
            if (boundingBox == null)
                return new List<ClashZone>();
            
            // Prefer R-Tree for bounding box queries (optimized for this use case)
            if (_rTreeIndex.IsIndexed)
            {
                return _rTreeIndex.Query(boundingBox);
            }
            
            // Fallback to spatial grid
            if (_spatialGrid.IsBuilt)
            {
                return _spatialGrid.GetItemsInBounds(boundingBox);
            }
            
            // No spatial index available
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning("[CompositeZoneFilter] No spatial index built for bounds query");
            }
            return new List<ClashZone>();
        }
        
        // Helper methods
        
        private bool IsInBounds(ClashZone zone, BoundingBoxXYZ box)
        {
            if (zone.IntersectionPoint == null) return false;
            
            var pt = zone.IntersectionPoint;
            return pt.X >= box.Min.X && pt.X <= box.Max.X &&
                   pt.Y >= box.Min.Y && pt.Y <= box.Max.Y &&
                   pt.Z >= box.Min.Z && pt.Z <= box.Max.Z;
        }
    }
}
