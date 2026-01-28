using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Strategies;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction
{
    /// <summary>
    /// Unified parameter extraction service using Strategy Pattern.
    /// This is the main entry point for all element parameter extraction.
    /// 
    /// Usage:
    /// var extractor = new ElementParameterExtractor();
    /// var snapshot = extractor.ExtractParameters(element);
    /// </summary>
    public class ElementParameterExtractor : IElementParameterExtractor
    {
        private readonly List<IParameterExtractionStrategy> _strategies;

        /// <summary>
        /// Create extractor with default strategies.
        /// </summary>
        public ElementParameterExtractor()
        {
            _strategies = new List<IParameterExtractionStrategy>
            {
                new DamperParameterExtractionStrategy(),
                new RectangularMepParameterExtractionStrategy(),
                new WallParameterExtractionStrategy(),
                new FloorParameterExtractionStrategy(),
                new FramingParameterExtractionStrategy(),
                new DefaultParameterExtractionStrategy() // Fallback for unknown elements
            };

            // Sort by priority (descending - higher priority first)
            _strategies = _strategies.OrderByDescending(s => s.Priority).ToList();
        }

        /// <summary>
        /// Create extractor with custom strategies.
        /// </summary>
        public ElementParameterExtractor(IEnumerable<IParameterExtractionStrategy> strategies)
        {
            _strategies = strategies.OrderByDescending(s => s.Priority).ToList();
        }

        /// <summary>
        /// Extract parameters from a single element.
        /// Automatically selects the appropriate strategy.
        /// </summary>
        public ElementParameterSnapshot ExtractParameters(Element element)
        {
            if (element == null) return new ElementParameterSnapshot();

            var strategy = GetStrategy(element);
            return strategy.Extract(element);
        }

        /// <summary>
        /// Extract parameters with custom options.
        /// </summary>
        public ElementParameterSnapshot ExtractParameters(Element element, ParameterExtractionOptions options)
        {
            // For now, options are handled within strategies
            // Future: pass options to strategy.Extract(element, options)
            return ExtractParameters(element);
        }

        /// <summary>
        /// Batch extract parameters from multiple elements.
        /// More efficient than individual extraction due to strategy caching.
        /// </summary>
        public Dictionary<ElementId, ElementParameterSnapshot> ExtractBatch(IEnumerable<Element> elements)
        {
            var results = new Dictionary<ElementId, ElementParameterSnapshot>();
            
            if (elements == null) return results;

            foreach (var element in elements)
            {
                if (element?.Id == null) continue;
                results[element.Id] = ExtractParameters(element);
            }

            return results;
        }

        /// <summary>
        /// Get the appropriate strategy for an element.
        /// </summary>
        private IParameterExtractionStrategy GetStrategy(Element element)
        {
            foreach (var strategy in _strategies)
            {
                if (strategy.CanHandle(element))
                {
                    return strategy;
                }
            }

            // Should never happen if DefaultStrategy is registered
            return new DefaultParameterExtractionStrategy();
        }
    }
}
