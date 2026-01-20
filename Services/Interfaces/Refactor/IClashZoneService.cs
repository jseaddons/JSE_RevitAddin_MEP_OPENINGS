using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team G: Main clash zone service interface (orchestrator).
    /// SOLID: Depends on abstractions only.
    /// 
    /// This interface abstracts clash zone operations, enabling:
    /// - Dependency injection for testability
    /// - SOLID compliance (DIP - depends on abstractions)
    /// - Preserved functionality (cleanup, filtering, detection)
    /// </summary>
    public interface IClashZoneService
    {
        /// <summary>
        /// Cleanup invalid clash zones (removes null elements and duplicates).
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <returns>Number of clash zones removed</returns>
        int CleanupInvalidClashZones(Document document);
        
        /// <summary>
        /// Filter clash zones by current selection parameters.
        /// </summary>
        /// <param name="selectedReferenceFiles">List of selected reference files</param>
        /// <param name="currentClearanceSettings">Current clearance settings by category</param>
        /// <param name="currentPrefix">Current prefix for filter names</param>
        /// <param name="document">Revit document</param>
        /// <returns>Filtered list of clash zones</returns>
        List<ClashZone> FilterClashZonesByCurrentSelection(
            List<string> selectedReferenceFiles,
            Dictionary<string, double> currentClearanceSettings,
            string currentPrefix,
            Document document);
        
        /// <summary>
        /// Detect new clash zones from intersections.
        /// Converts raw intersections into ClashZone objects with validation and deduplication.
        /// </summary>
        /// <param name="currentIntersections">List of intersections (MEP element, structural element, bounding box, intersection point)</param>
        /// <param name="document">Revit document</param>
        /// <param name="clearanceSettings">Optional clearance settings by category</param>
        /// <param name="selectedCategories">Optional list of selected MEP categories</param>
        /// <returns>List of newly detected clash zones</returns>
        List<ClashZone> DetectNewClashZones(
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections,
            Document document,
            Dictionary<string, double> clearanceSettings = null,
            List<string> selectedCategories = null);
    }
}

