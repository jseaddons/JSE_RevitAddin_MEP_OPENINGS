using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor;
using JSE_RevitAddin_MEP_OPENINGS.Services.Logging;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClashZoneManagement
{
    /// <summary>
    /// Team G: SOLID-compliant duct-damper filtering service.
    /// 
    /// This service handles the logic to skip ducts when dampers are present at the same intersection.
    /// SOLID: Single Responsibility - duct-damper filtering only.
    /// 
    /// ✅ PRESERVES ALL OPTIMIZATIONS:
    /// - Pre-calculates damper locations once (Calculate Once, Use Many Times)
    /// - Checks same wall requirement first (major optimization)
    /// - Uses HasDamperNearby flag to avoid re-checking on subsequent refreshes
    /// - Multiple proximity check methods (intersection point, bbox, distance)
    /// </summary>
    public class DuctDamperFilterService : IDuctDamperFilterService
    {
        private readonly ILogger _logger;
        
        /// <summary>
        /// Creates a new duct-damper filter service.
        /// </summary>
        /// <param name="logger">Optional logger for tracking operations</param>
        public DuctDamperFilterService(ILogger logger = null)
        {
            _logger = logger ?? LoggerAdapter.Default;
        }
        
        /// <summary>
        /// Pre-calculates all damper locations from storage and current intersections.
        /// This optimization calculates damper locations once and reuses them for all duct checks.
        /// </summary>
        public List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> PreCalculateDamperLocations(
            ClashZoneStorage storage,
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document)
        {
            var damperLocations = new List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)>();
            
            try
            {
                // ✅ STEP 1: Get dampers from saved storage (from previous refresh cycles)
                if (storage?.ClashZones != null)
                {
                    foreach (var clashZone in storage.ClashZones)
                    {
                        if (IsDamperClashZone(clashZone))
                        {
                            // Try to get the damper element from document
                            var damperElement = document.GetElement(clashZone.MepElementId);
                            
                            BoundingBoxXYZ damperBbox = null;
                            if (damperElement != null)
                            {
                                // Element is in active document or same linked file
                                damperBbox = damperElement.get_BoundingBox(null);
                            }
                            else
                            {
                                // Element might be in a different linked file - use intersection point approximation
                                if (Math.Abs(clashZone.IntersectionPointX) > 1e-9 ||
                                    Math.Abs(clashZone.IntersectionPointY) > 1e-9 ||
                                    Math.Abs(clashZone.IntersectionPointZ) > 1e-9)
                                {
                                    var intersectionPoint = new XYZ(
                                        clashZone.IntersectionPointX, 
                                        clashZone.IntersectionPointY, 
                                        clashZone.IntersectionPointZ);
                                    // Create a small bounding box around intersection point (approx 200mm = 0.66ft cube)
                                    const double approximateSize = 0.66; // 200mm in feet
                                    damperBbox = new BoundingBoxXYZ
                                    {
                                        Min = new XYZ(
                                            intersectionPoint.X - approximateSize, 
                                            intersectionPoint.Y - approximateSize, 
                                            intersectionPoint.Z - approximateSize),
                                        Max = new XYZ(
                                            intersectionPoint.X + approximateSize, 
                                            intersectionPoint.Y + approximateSize, 
                                            intersectionPoint.Z + approximateSize)
                                    };
                                    _logger.Debug($"Damper {clashZone.MepElementId} in linked file - using intersection point approximation", "DuctDamperFilter");
                                }
                            }
                            
                            if (damperBbox != null)
                            {
                                damperLocations.Add((clashZone.MepElementId, damperBbox, clashZone.StructuralElementId));
                            }
                        }
                    }
                }
                
                // ✅ STEP 2: Add dampers from current intersections (from current refresh cycle)
                foreach (var (mepElement, structuralElement, boundingBox, intersectionPoint) in currentIntersections)
                {
                    if (IsDamperElement(mepElement))
                    {
                        var damperBbox = mepElement.get_BoundingBox(null);
                        if (damperBbox != null)
                        {
                            // Avoid duplicates (check if already added from storage)
                            if (!damperLocations.Any(d => d.damperId == mepElement.Id))
                            {
                                damperLocations.Add((mepElement.Id, damperBbox, structuralElement.Id));
                            }
                        }
                    }
                }
                
                _logger.Info($"Pre-calculated {damperLocations.Count} damper locations", "DuctDamperFilter");
            }
            catch (Exception ex)
            {
                _logger.Error($"Error pre-calculating damper locations: {ex.Message}", ex, "DuctDamperFilter");
            }
            
            return damperLocations;
        }
        
        /// <summary>
        /// Checks if a duct should be skipped because a damper is nearby on the same wall.
        /// Returns true if duct should be skipped (damper takes priority).
        /// </summary>
        public bool ShouldSkipDuctForDamper(
            Element ductElement,
            ElementId wallId,
            XYZ intersectionPoint,
            List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> damperLocations,
            ClashZone existingClashZone = null)
        {
            try
            {
                // ✅ OPTIMIZATION: Check existing clash zone for saved flag (fastest check - no proximity calculation needed)
                if (existingClashZone != null && existingClashZone.HasDamperNearby)
                {
                    _logger.Debug($"Duct {ductElement.Id} - HasDamperNearby flag is true, skipping duct", "DuctDamperFilter");
                    return true; // Skip duct
                }
                
                // ✅ OPTIMIZATION: Check same wall requirement first (major optimization)
                var dampersOnSameWall = damperLocations?
                    .Where(d => d.wallId == wallId)
                    .ToList() ?? new List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)>();
                
                if (dampersOnSameWall.Count == 0)
                {
                    // ✅ Clear flag if damper no longer nearby (damper might have been deleted)
                    if (existingClashZone != null && existingClashZone.HasDamperNearby)
                    {
                        existingClashZone.HasDamperNearby = false;
                        _logger.Debug($"Cleared HasDamperNearby flag - damper no longer nearby for clash zone {existingClashZone.Id}", "DuctDamperFilter");
                    }
                    return false; // No dampers on same wall, don't skip
                }
                
                // Get duct bounding box
                var ductBbox = ductElement.get_BoundingBox(null);
                if (ductBbox == null)
                {
                    return false; // Cannot check without bounding box
                }
                
                // ✅ TOLERANCES: Same as legacy code
                const double intersectionTolerance = 0.2; // 200mm tolerance for damper at intersection point
                const double bboxProximityTolerance = 0.5; // 6 inches tolerance for connected duct-damper pairs
                
                // Check pre-calculated damper locations on SAME WALL
                foreach (var (damperId, damperBbox, _) in dampersOnSameWall)
                {
                    // Skip if it's the same element
                    if (damperId == ductElement.Id) continue;
                    
                    // ✅ METHOD 1: Check if damper center is near intersection point (most reliable)
                    XYZ damperCenter = (damperBbox.Min + damperBbox.Max) * 0.5;
                    double distanceToIntersection = damperCenter.DistanceTo(intersectionPoint);
                    
                    if (distanceToIntersection <= intersectionTolerance)
                    {
                        _logger.Debug($"Duct {ductElement.Id} intersection point is {distanceToIntersection:F4}ft from Damper {damperId} center - SKIP DUCT", "DuctDamperFilter");
                        return true; // Skip duct
                    }
                    
                    // ✅ METHOD 2: Check if damper bbox contains or is near intersection point
                    if (IsPointNearBoundingBox(intersectionPoint, damperBbox, intersectionTolerance))
                    {
                        _logger.Debug($"Duct {ductElement.Id} intersection point is within {intersectionTolerance}ft of Damper {damperId} bbox - SKIP DUCT", "DuctDamperFilter");
                        return true; // Skip duct
                    }
                    
                    // ✅ METHOD 3: Fallback - Check if duct and damper bounding boxes are within proximity
                    var bboxDistance = GetMinimumDistanceBetweenBoundingBoxes(ductBbox, damperBbox);
                    if (bboxDistance <= bboxProximityTolerance)
                    {
                        _logger.Debug($"Duct {ductElement.Id} bbox is {bboxDistance:F4}ft from Damper {damperId} bbox - SKIP DUCT", "DuctDamperFilter");
                        return true; // Skip duct
                    }
                }
                
                // ✅ Clear flag if damper no longer nearby (damper might have been deleted)
                if (existingClashZone != null && existingClashZone.HasDamperNearby)
                {
                    existingClashZone.HasDamperNearby = false;
                    _logger.Debug($"Cleared HasDamperNearby flag - damper no longer nearby for clash zone {existingClashZone.Id}", "DuctDamperFilter");
                }
                
                return false; // No damper nearby, don't skip
            }
            catch (Exception ex)
            {
                _logger.Warning($"Error checking duct-damper proximity: {ex.Message}", "DuctDamperFilter");
                return false; // Fail-safe: don't skip on error
            }
        }
        
        /// <summary>
        /// Checks if an element is a damper (Duct Accessory).
        /// </summary>
        public bool IsDamperElement(Element element)
        {
            try
            {
                // ✅ SIMPLE: If it's a Duct Accessory category, treat it as a damper
                return element.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Checks if a clash zone represents a damper.
        /// </summary>
        public bool IsDamperClashZone(ClashZone clashZone)
        {
            try
            {
                // ✅ SIMPLE: If it's a Duct Accessory category, treat it as a damper
                return string.Equals(clashZone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Check if a point is near a bounding box (within tolerance).
        /// </summary>
        private bool IsPointNearBoundingBox(XYZ point, BoundingBoxXYZ bbox, double tolerance)
        {
            try
            {
                // Expand bbox by tolerance
                var expandedMin = new XYZ(bbox.Min.X - tolerance, bbox.Min.Y - tolerance, bbox.Min.Z - tolerance);
                var expandedMax = new XYZ(bbox.Max.X + tolerance, bbox.Max.Y + tolerance, bbox.Max.Z + tolerance);
                
                // Check if point is within expanded bbox
                return point.X >= expandedMin.X && point.X <= expandedMax.X &&
                       point.Y >= expandedMin.Y && point.Y <= expandedMax.Y &&
                       point.Z >= expandedMin.Z && point.Z <= expandedMax.Z;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Calculate minimum distance between two bounding boxes.
        /// </summary>
        private double GetMinimumDistanceBetweenBoundingBoxes(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
        {
            try
            {
                // Find closest points between boxes
                var closest1 = new XYZ(
                    Math.Max(bbox1.Min.X, Math.Min(bbox1.Max.X, (bbox2.Min.X + bbox2.Max.X) / 2)),
                    Math.Max(bbox1.Min.Y, Math.Min(bbox1.Max.Y, (bbox2.Min.Y + bbox2.Max.Y) / 2)),
                    Math.Max(bbox1.Min.Z, Math.Min(bbox1.Max.Z, (bbox2.Min.Z + bbox2.Max.Z) / 2))
                );
                
                var closest2 = new XYZ(
                    Math.Max(bbox2.Min.X, Math.Min(bbox2.Max.X, (bbox1.Min.X + bbox1.Max.X) / 2)),
                    Math.Max(bbox2.Min.Y, Math.Min(bbox2.Max.Y, (bbox1.Min.Y + bbox1.Max.Y) / 2)),
                    Math.Max(bbox2.Min.Z, Math.Min(bbox2.Max.Z, (bbox1.Min.Z + bbox1.Max.Z) / 2))
                );
                
                return closest1.DistanceTo(closest2);
            }
            catch
            {
                return double.MaxValue; // Return large distance if calculation fails
            }
        }
    }
}

