using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces
{
    /// <summary>
    /// Unified interface for extracting parameters from any Revit element.
    /// Replaces multiple scattered parameter extraction implementations.
    /// </summary>
    public interface IElementParameterExtractor
    {
        /// <summary>
        /// Extract parameters from a single element.
        /// </summary>
        /// <param name="element">The Revit element (MEP, damper, wall, floor, etc.)</param>
        /// <returns>Snapshot containing all extracted parameters</returns>
        ElementParameterSnapshot ExtractParameters(Element element);

        /// <summary>
        /// Extract parameters from a single element with custom options.
        /// </summary>
        /// <param name="element">The Revit element</param>
        /// <param name="options">Extraction options (e.g., which parameters to capture)</param>
        /// <returns>Snapshot containing extracted parameters</returns>
        ElementParameterSnapshot ExtractParameters(Element element, ParameterExtractionOptions options);

        /// <summary>
        /// Batch extract parameters from multiple elements.
        /// More efficient than calling ExtractParameters individually.
        /// </summary>
        /// <param name="elements">Collection of Revit elements</param>
        /// <returns>Dictionary mapping ElementId to parameter snapshots</returns>
        Dictionary<ElementId, ElementParameterSnapshot> ExtractBatch(IEnumerable<Element> elements);
    }
}
