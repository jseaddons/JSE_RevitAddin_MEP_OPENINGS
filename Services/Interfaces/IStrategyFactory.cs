using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces
{
    /// <summary>
    /// Interface for creating sleeve placement strategies.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public interface IStrategyFactory
    {
        /// <summary>
        /// Create a placement strategy for the given category.
        /// </summary>
        ISleevePlacementStrategy CreateStrategy(string category, Document doc = null);
    }
}

