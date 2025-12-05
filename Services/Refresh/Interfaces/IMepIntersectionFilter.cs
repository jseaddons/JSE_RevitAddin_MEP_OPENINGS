using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// ✅ SOLID: Single Responsibility + Dependency Inversion
    /// 
    /// Interface for filtering MEP intersections based on proximity rules.
    /// Implementations can filter ducts near dampers, cables near pipes, etc.
    /// 
    /// Clients depend on THIS interface, not concrete implementations.
    /// Makes the system extensible and testable.
    /// 
    /// Current implementations:
    /// - IDuctDamperProximityFilter: Filters ducts near dampers on same wall
    /// - (Future) ICablePipeProximityFilter: Filters cables near pipes
    /// - (Future) ICustomProximityFilter: Custom user-defined rules
    /// </summary>
    public interface IMepIntersectionFilter
    {
        /// <summary>
        /// Filter intersections based on proximity rules.
        /// 
        /// Returns subset of intersections that should be KEPT (all others are filtered OUT).
        /// Example: Duct-damper filter returns "ducts NOT near dampers" + "all dampers"
        /// </summary>
        /// <param name="allIntersections">
        /// All detected intersections from MepIntersectionService.
        /// Tuple format: (MEP element, Host element, Host bbox, Intersection point)
        /// </param>
        /// <param name="doc">Revit document for element access</param>
        /// <param name="logger">Logging callback for debugging and diagnostics</param>
        /// <returns>Filtered intersections to process (same format as input)</returns>
        List<(Element mepElement, Element hostElement, BoundingBoxXYZ hostBBox, XYZ intersectionPoint)> FilterIntersections(
            List<(Element mepElement, Element hostElement, BoundingBoxXYZ hostBBox, XYZ intersectionPoint)> allIntersections,
            Document doc,
            Action<string> logger);
    }

    /// <summary>
    /// ✅ SOLID: Interface Segregation
    /// 
    /// Specialized interface for duct-damper proximity filtering.
    /// Extends IMepIntersectionFilter for duct-damper-specific behavior.
    /// 
    /// Clients that ONLY need duct-damper filtering depend on THIS interface
    /// instead of the more general IMepIntersectionFilter.
    /// This follows Interface Segregation Principle.
    /// </summary>
    public interface IDuctDamperProximityFilter : IMepIntersectionFilter
    {
        // Can add duct-damper-specific methods here if needed in future
        // For now, just uses inherited FilterIntersections method
    }
}
