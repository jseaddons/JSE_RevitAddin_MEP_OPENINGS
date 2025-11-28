using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team G: Interface for duct-damper filtering service.
    /// SOLID: Interface Segregation - focused interface for duct-damper filtering only.
    /// 
    /// This service handles the logic to skip ducts when dampers are present at the same intersection,
    /// ensuring only dampers get sleeves (not both duct and damper).
    /// </summary>
    public interface IDuctDamperFilterService
    {
        /// <summary>
        /// Pre-calculates all damper locations from storage and current intersections.
        /// This optimization calculates damper locations once and reuses them for all duct checks.
        /// </summary>
        /// <param name="storage">Clash zone storage (for existing damper clash zones)</param>
        /// <param name="currentIntersections">Current intersections (may contain dampers)</param>
        /// <param name="document">Revit document</param>
        /// <returns>List of damper locations (damper ID, bounding box, wall ID)</returns>
        List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> PreCalculateDamperLocations(
            ClashZoneStorage storage,
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document);
        
        /// <summary>
        /// Checks if a duct should be skipped because a damper is nearby on the same wall.
        /// Returns true if duct should be skipped (damper takes priority).
        /// </summary>
        /// <param name="ductElement">The duct element to check</param>
        /// <param name="wallId">The wall ID where the intersection occurs</param>
        /// <param name="intersectionPoint">The intersection point</param>
        /// <param name="damperLocations">Pre-calculated damper locations</param>
        /// <param name="existingClashZone">Existing clash zone for this duct (may have HasDamperNearby flag)</param>
        /// <returns>True if duct should be skipped (damper nearby), false otherwise</returns>
        bool ShouldSkipDuctForDamper(
            Element ductElement,
            ElementId wallId,
            XYZ intersectionPoint,
            List<(ElementId damperId, BoundingBoxXYZ bbox, ElementId wallId)> damperLocations,
            ClashZone existingClashZone = null);
        
        /// <summary>
        /// Checks if an element is a damper (Duct Accessory).
        /// </summary>
        /// <param name="element">Element to check</param>
        /// <returns>True if element is a damper</returns>
        bool IsDamperElement(Element element);
        
        /// <summary>
        /// Checks if a clash zone represents a damper.
        /// </summary>
        /// <param name="clashZone">Clash zone to check</param>
        /// <returns>True if clash zone represents a damper</returns>
        bool IsDamperClashZone(ClashZone clashZone);
    }
}

