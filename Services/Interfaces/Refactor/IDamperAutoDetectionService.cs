using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team G: Interface for auto-detecting missing dampers.
    /// SOLID: Interface Segregation - focused interface for auto-detection only.
    /// 
    /// This service automatically detects dampers that user forgot to select,
    /// preventing incorrect sleeve placement when ducts have dampers at their ends.
    /// </summary>
    public interface IDamperAutoDetectionService
    {
        /// <summary>
        /// Auto-detects missing dampers when user selected ducts but forgot Duct Accessories.
        /// Enhances the intersection list with damper intersections found on the same walls as ducts.
        /// </summary>
        /// <param name="document">Revit document</param>
        /// <param name="currentIntersections">Current intersections (may not include dampers)</param>
        /// <returns>Enhanced intersection list with auto-detected dampers added</returns>
        List<(Element, Element, BoundingBoxXYZ, XYZ)> AutoDetectMissingDampers(
            Document document,
            List<(Element, Element, BoundingBoxXYZ, XYZ)> currentIntersections);
    }
}

