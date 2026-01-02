using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Strategies
{
    /// <summary>
    /// Default/fallback strategy for elements that don't match any specific strategy.
    /// Extracts basic parameters (width, height, diameter) using common parameter names.
    /// </summary>
    public class DefaultParameterExtractionStrategy : IParameterExtractionStrategy
    {
        // Lowest priority - only used when no other strategy matches
        public int Priority => 0;

        /// <summary>
        /// Default strategy can handle any element (fallback).
        /// </summary>
        public bool CanHandle(Element element) => element != null;

        /// <summary>
        /// Extract basic parameters using common parameter names.
        /// </summary>
        public ElementParameterSnapshot Extract(Element element)
        {
            var snapshot = new ElementParameterSnapshot();
            if (element == null) return snapshot;

            // Extract all parameters
            foreach (Parameter param in element.Parameters)
            {
                string name = param.Definition?.Name ?? string.Empty;
                if (string.IsNullOrEmpty(name)) continue;

                object? value = param.StorageType switch
                {
                    StorageType.Double => param.AsDouble(),
                    StorageType.Integer => param.AsInteger(),
                    StorageType.String => param.AsString(),
                    StorageType.ElementId => param.AsElementId()?.IntegerValue,
                    _ => null
                };

                if (value != null)
                {
                    snapshot.AllParameters[name] = value;
                }
            }

            // Try common dimension parameters
            snapshot.Width = TryGetDoubleParameter(element, "Width", "DIMENSIONS_WIDTH");
            snapshot.Height = TryGetDoubleParameter(element, "Height", "DIMENSIONS_HEIGHT");
            snapshot.OuterDiameter = TryGetDoubleParameter(element, "Diameter", "DIMENSION_DIAMETER");

            // Try level parameter
            var levelParam = element.LookupParameter("Level") ?? element.LookupParameter("Reference Level");
            if (levelParam != null && levelParam.HasValue && levelParam.StorageType == StorageType.ElementId)
            {
                var levelId = levelParam.AsElementId();
                if (levelId != null && levelId != ElementId.InvalidElementId)
                {
                    var level = element.Document.GetElement(levelId) as Level;
                    if (level != null)
                    {
                        snapshot.LevelName = level.Name;
                        snapshot.LevelElevation = level.Elevation;
                    }
                }
            }

            return snapshot;
        }

        private double? TryGetDoubleParameter(Element element, params string[] paramNames)
        {
            foreach (var name in paramNames)
            {
                var param = element.LookupParameter(name);
                if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                {
                    return param.AsDouble();
                }
            }
            return null;
        }
    }
}
