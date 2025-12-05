using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactored
{
    /// <summary>
    /// ✅ SOLID REFACTORED: Factory for creating sleeve placement strategies.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public class StrategyFactoryService : IStrategyFactory
    {
        public ISleevePlacementStrategy CreateStrategy(string category, Document doc = null)
        {
            return category switch
            {
                "Ducts" => new DuctPlacementStrategy(),
                "Pipes" => new PipePlacementStrategy(),
                "Cable Trays" => new CableTrayPlacementStrategy(),
                "Duct Accessories" => new DamperPlacementStrategy(doc),
                _ => throw new System.ArgumentException($"Unknown category: {category}")
            };
        }
    }
}

