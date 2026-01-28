using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Strategies
{
    /// <summary>
    /// Strategy for extracting parameters from RECTANGULAR MEP elements (Ducts, Cable Trays).
    /// Critically handles orientation (Cos/Sin) for floor sleeve placement.
    /// Excludes circular elements (Pipes, Round Ducts).
    /// </summary>
    public class RectangularMepParameterExtractionStrategy : IParameterExtractionStrategy
    {
        // High priority to catch ducts/trays before Default strategy
        public int Priority => 90;

        public bool CanHandle(Element element)
        {
            if (element == null) return false;

            // Check Category
            var builtinCat = (BuiltInCategory)element.Category.Id.IntegerValue;
            bool isCategoryMatch = builtinCat == BuiltInCategory.OST_DuctCurves ||
                                   builtinCat == BuiltInCategory.OST_CableTray;

            if (!isCategoryMatch) return false;

            // Check Shape/Profile (Exclude Circular)
            if (element is Duct duct)
            {
                // Verify it's not round/oval (roughly)
                // Best check: Does it have Width/Height? Round ducts usually don't.
                // Or check parameters.
                return HasRectangularDimensions(element);
            }

            if (element is CableTray)
            {
                // Cable trays are almost always rectangular/channel
                return true;
            }

            return false;
        }

        private bool HasRectangularDimensions(Element element)
        {
            // Check for Width and Height parameters
            bool hasWidth = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM) != null ||
                            element.LookupParameter("Width") != null;
            bool hasHeight = element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM) != null ||
                             element.LookupParameter("Height") != null;

            return hasWidth && hasHeight;
        }

        public ElementParameterSnapshot Extract(Element element)
        {
            var snapshot = new ElementParameterSnapshot();
            ExtractStandardParameters(element, snapshot);
            ExtractRectangularDimensions(element, snapshot);
            ExtractOrientation(element, snapshot);
            
            // Level info
            var levelParam = element.LookupParameter("Reference Level");
            if (levelParam != null && levelParam.AsElementId() != ElementId.InvalidElementId)
            {
                var level = element.Document.GetElement(levelParam.AsElementId()) as Level;
                snapshot.LevelName = level?.Name;
                snapshot.LevelElevation = level?.Elevation;
            }

            return snapshot;
        }

        private void ExtractStandardParameters(Element element, ElementParameterSnapshot snapshot)
        {
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
        }

        private void ExtractRectangularDimensions(Element element, ElementParameterSnapshot snapshot)
        {
            // Width
            var widthParam = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM) 
                             ?? element.LookupParameter("Width");
            if (widthParam != null) snapshot.Width = widthParam.AsDouble();

            // Height
            var heightParam = element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                              ?? element.LookupParameter("Height");
            if (heightParam != null) snapshot.Height = heightParam.AsDouble();
            
            // Size
             var sizeParam = element.LookupParameter("Size");
            if (sizeParam != null) snapshot.SizeParameterValue = sizeParam.AsString();
        }

        private void ExtractOrientation(Element element, ElementParameterSnapshot snapshot)
        {
            if (element.Location is LocationCurve locCurve)
            {
                XYZ start = locCurve.Curve.GetEndPoint(0);
                XYZ end = locCurve.Curve.GetEndPoint(1);
                XYZ direction = (end - start).Normalize();

                snapshot.OrientationX = direction.X;
                snapshot.OrientationY = direction.Y;
                snapshot.OrientationZ = direction.Z;

                // Calculate Rotation relative to X-axis
                // Angle in XY plane
                double angleRad = Math.Atan2(direction.Y, direction.X);
                
                // Normalize angle to [0, 2PI]
                if (angleRad < 0) angleRad += 2 * Math.PI;

                snapshot.RotationAngleRad = angleRad;
                snapshot.RotationAngleDeg = angleRad * (180.0 / Math.PI);

                // Pre-calculate Cos/Sin for placement optimization
                snapshot.RotationCos = Math.Cos(angleRad);
                snapshot.RotationSin = Math.Sin(angleRad);
                
                 snapshot.AngleToXRad = Math.Abs(Math.Atan2(direction.Y, direction.X)); 
                 snapshot.AngleToYRad = Math.Abs(Math.Atan2(direction.X, direction.Y)); // Complementary

                // Determine dominant direction
                if (Math.Abs(direction.Z) > 0.9)
                {
                    snapshot.OrientationDirection = direction.Z > 0 ? "+Z" : "-Z";
                }
                else if (Math.Abs(direction.X) > Math.Abs(direction.Y))
                {
                    snapshot.OrientationDirection = direction.X > 0 ? "+X" : "-X";
                }
                else
                {
                    snapshot.OrientationDirection = direction.Y > 0 ? "+Y" : "-Y";
                }
            }
        }
    }
}
