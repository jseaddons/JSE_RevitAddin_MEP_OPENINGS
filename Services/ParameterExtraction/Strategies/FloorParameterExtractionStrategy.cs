using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Strategies
{
    /// <summary>
    /// Strategy for extracting parameters from Floor (Host) elements.
    /// Extracts:
    /// - Structural Type (Floor)
    /// - Thickness
    /// - Fire Rating
    /// - Level Info
    /// </summary>
    public class FloorParameterExtractionStrategy : IParameterExtractionStrategy
    {
        public int Priority => 80;

        public bool CanHandle(Element element)
        {
            return element is Floor;
        }

        public ElementParameterSnapshot Extract(Element element)
        {
            var snapshot = new ElementParameterSnapshot
            {
                StructuralType = "Floor"
            };

            if (element is Floor floor)
            {
                // Thickness (from Type parameter usually)
                var thicknessParam = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM) 
                                     ?? floor.get_Parameter(BuiltInParameter.STRUCTURAL_FLOOR_CORE_THICKNESS);
                
                if (thicknessParam != null && thicknessParam.HasValue)
                {
                    snapshot.Thickness = thicknessParam.AsDouble();
                    snapshot.WallThickness = snapshot.Thickness; // Reusing field for general thickness logic
                }

                // Level Info
                var levelParam = floor.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                if (levelParam != null && levelParam.AsElementId() != ElementId.InvalidElementId)
                {
                    var level = floor.Document.GetElement(levelParam.AsElementId()) as Level;
                    snapshot.LevelName = level?.Name;
                    snapshot.LevelElevation = level?.Elevation;
                }
            }

            // Extract all parameters (Fire Rating often instance)
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
                        StorageType.ElementId => param.AsElementId()?.GetIntegerValue(),
                        _ => null
                    };

                    if (value != null)
                    {
                        snapshot.AllParameters[name] = value;
                    }
                }
            }
            
            // Explicit check for Fire Rating if not found via loop
            if (!snapshot.AllParameters.ContainsKey("Fire Rating"))
            {
                 var fireRating = element.get_Parameter(BuiltInParameter.FIRE_RATING);
                if (fireRating != null && fireRating.HasValue)
                {
                    snapshot.AllParameters["Fire Rating"] = fireRating.AsString();
                }
            }

            return snapshot;
        }
    }
}