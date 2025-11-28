using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces.Refactor
{
    /// <summary>
    /// Team G: Clash zone validation operations (removed from ClashZoneService).
    /// SOLID: Single Responsibility - validation only.
    /// 
    /// This interface abstracts validation operations, enabling:
    /// - Dependency injection for testability
    /// - SOLID compliance (SRP - single responsibility)
    /// - Preserved validation logic
    /// </summary>
    public interface IClashZoneValidationService
    {
        /// <summary>
        /// Validate clash zone elements exist and still intersect.
        /// 
        /// ✅ PRESERVES:
        /// - MEP element existence check
        /// - Structural element existence check
        /// - Intersection point validation
        /// - Fail-safe error handling
        /// </summary>
        /// <param name="clashZone">Clash zone to validate</param>
        /// <param name="document">Revit document</param>
        /// <returns>True if clash zone is valid, false otherwise</returns>
        bool ValidateClashZone(ClashZone clashZone, Document document);
    }
}

