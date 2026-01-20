using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team G: Clash zone detection service interface.
    /// SOLID: Single Responsibility - detection only.
    /// 
    /// This interface abstracts clash zone detection operations, enabling:
    /// - Dependency injection for testability
    /// - SOLID compliance (DIP - depends on abstractions)
    /// - Preserved functionality (detection, validation, deduplication)
    /// </summary>
    public interface IClashZoneDetectionService
    {
        /// <summary>
        /// Detect new clash zones from intersections.
        /// Converts raw intersections into ClashZone objects with validation and deduplication.
        /// </summary>
        /// <param name="currentIntersections">List of intersections (MEP element, structural element, bounding box, intersection point)</param>
        /// <param name="document">Revit document</param>
        /// <param name="storage">Clash zone storage (for existing zones and to add new ones)</param>
        /// <param name="clearanceSettings">Optional clearance settings by category</param>
        /// <param name="selectedCategories">Optional list of selected MEP categories</param>
        /// <returns>List of newly detected clash zones</returns>
        List<ClashZone> DetectNewClashZones(
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document,
            ClashZoneStorage storage,
            Dictionary<string, double> clearanceSettings = null,
            List<string> selectedCategories = null);
    }
}

