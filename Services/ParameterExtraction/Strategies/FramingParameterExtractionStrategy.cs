using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Strategies
{
    /// <summary>
    /// Strategy for extracting parameters from Structural Framing (Beam/Column) elements.
    /// Extracts:
    /// - Structural Type (StructuralFraming)
    /// - Thickness/Dimensions
    /// - Fire Rating
    /// - Level Info
    /// </summary>
    public class FramingParameterExtractionStrategy : IParameterExtractionStrategy
    {
        public int Priority => 80;

        public bool CanHandle(Element element)
        {
            if (element?.Category == null) return false;
            return element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming ||
                   element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns;
        }

        public ElementParameterSnapshot Extract(Element element)
        {
             var snapshot = new ElementParameterSnapshot
            {
                StructuralType = "StructuralFraming"
            };

            // Extract Reference Level
            var levelParam = element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
            if (levelParam != null && levelParam.AsElementId() != ElementId.InvalidElementId)
            {
                var level = element.Document.GetElement(levelParam.AsElementId()) as Level;
                snapshot.LevelName = level?.Name;
                snapshot.LevelElevation = level?.Elevation;
            }

            // Extract Dimensions (Width/Height/Thickness depending on family)
            var type = element.Document.GetElement(element.GetTypeId()) as FamilySymbol;
            if (type != null)
            {
                // Try to find thickness/width/height from Type parameters
                 ExtractTypeDimension(type, "b", snapshot); // Width (common for concrete)
                 ExtractTypeDimension(type, "h", snapshot); // Height
                 ExtractTypeDimension(type, "Width", snapshot);
                 ExtractTypeDimension(type, "Height", snapshot);
            }

            // Fire Rating
            var fireRating = element.get_Parameter(BuiltInParameter.FIRE_RATING);
            if (fireRating != null && fireRating.HasValue)
            {
                snapshot.AllParameters["Fire Rating"] = fireRating.AsString();
            }

            // Extract all instance parameters
            foreach (Parameter param in element.Parameters)
            {
                string name = param.Definition?.Name ?? string.Empty;
                if (!string.IsNullOrEmpty(name))
                {
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
            }

            return snapshot;
        }

        private void ExtractTypeDimension(FamilySymbol type, string paramName, ElementParameterSnapshot snapshot)
        {
            var param = type.LookupParameter(paramName);
            if (param != null && param.HasValue && param.StorageType == StorageType.Double)
            {
                double val = param.AsDouble();
                // Map based on heuristic
                if (paramName.IndexOf("Width", StringComparison.OrdinalIgnoreCase) >= 0 || paramName == "b")
                {
                    snapshot.FramingThickness = val; // Often width acts as thickness for penetration
                }
                snapshot.AllParameters[paramName] = val;
            }
        }
    }
}
