using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Strategies
{
    /// <summary>
    /// Strategy for extracting parameters from Wall (Host) elements.
    /// Extracts:
    /// - Structural Type (Wall)
    /// - Thickness (Width)
    /// - Orientation (BasisX/Y/Z)
    /// - Fire Rating
    /// </summary>
    public class WallParameterExtractionStrategy : IParameterExtractionStrategy
    {
        public int Priority => 80;

        public bool CanHandle(Element element)
        {
            return element is Wall;
        }

        public ElementParameterSnapshot Extract(Element element)
        {
            var snapshot = new ElementParameterSnapshot
            {
                StructuralType = "Wall"
            };

            if (element is Wall wall)
            {
                // Thickness
                snapshot.Thickness = wall.Width;
                snapshot.WallThickness = wall.Width;

                // Orientation
                if (wall.Location is LocationCurve locCurve && locCurve.Curve is Line line)
                {
                    XYZ direction = line.Direction;
                    snapshot.HostOrientation = FormatOrientation(direction);
                    
                    // Also clear specific orientation vectors
                    snapshot.OrientationX = direction.X;
                    snapshot.OrientationY = direction.Y;
                    snapshot.OrientationZ = direction.Z;
                }

                // Fire Rating
                var fireRating = wall.get_Parameter(BuiltInParameter.FIRE_RATING);
                if (fireRating != null && fireRating.HasValue)
                {
                    snapshot.AllParameters["Fire Rating"] = fireRating.AsString();
                }
            }

            // Extract all other parameters
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

        private string FormatOrientation(XYZ dir)
        {
            return $"({dir.X:F2}, {dir.Y:F2}, {dir.Z:F2})";
        }
    }
}
