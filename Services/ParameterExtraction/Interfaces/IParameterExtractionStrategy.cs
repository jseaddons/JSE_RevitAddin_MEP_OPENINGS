using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces
{
    /// <summary>
    /// Strategy interface for element-type-specific parameter extraction.
    /// Implements Strategy Pattern for extensibility (Open/Closed Principle).
    /// </summary>
    public interface IParameterExtractionStrategy
    {
        /// <summary>
        /// Determines if this strategy can handle the given element.
        /// </summary>
        /// <param name="element">The Revit element to check</param>
        /// <returns>True if this strategy should be used for this element</returns>
        bool CanHandle(Element element);

        /// <summary>
        /// Extract parameters from the element using strategy-specific logic.
        /// </summary>
        /// <param name="element">The Revit element</param>
        /// <returns>Extracted parameter snapshot</returns>
        ElementParameterSnapshot Extract(Element element);

        /// <summary>
        /// Priority of this strategy (higher = checked first).
        /// Used when multiple strategies can handle an element.
        /// </summary>
        int Priority { get; }
    }
}
