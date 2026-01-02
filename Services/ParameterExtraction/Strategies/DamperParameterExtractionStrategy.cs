using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ParameterExtraction.Strategies
{
    /// <summary>
    /// Strategy for extracting parameters from damper (Duct Accessory) elements.
    /// Handles damper-specific logic:
    /// - Prioritized width/height extraction (Damper Width > Width > DIMENSIONS_WIDTH)
    /// - Connector detection for asymmetric clearance
    /// - MSFD/Motorized vs Standard Fire Damper detection
    /// </summary>
    public class DamperParameterExtractionStrategy : IParameterExtractionStrategy
    {
        // Priority for damper elements (high priority - checked early)
        public int Priority => 100;

        #region Parameter Name Priorities

        // Width: prioritize 'Damper Width', fallback to others
        private static readonly string[] WidthParametersMain = { "Damper Width" };
        private static readonly string[] WidthParametersFallback = { "Width", "width", "DIMENSIONS_WIDTH" };

        // Height: prioritize 'Damper Height', fallback to others
        private static readonly string[] HeightParametersMain = { "Damper Height" };
        private static readonly string[] HeightParametersFallback = { "Height", "height", "DIMENSIONS_HEIGHT" };

        // Diameter (for round dampers)
        private static readonly string[] DiameterParameters = { "Diameter", "DIMENSION_DIAMETER" };

        // Level parameters (prioritized)
        private static readonly string[] LevelParameters =
        {
            "Reference Level",
            "Level",
            "Schedule Level",
            "Schedule of Level"
        };

        // Type keywords for MSFD/Motorized detection
        private static readonly string[] MotorizedKeywords = { "MSFD", "MSD", "MD", "Motorized", "Motor" };

        #endregion

        /// <summary>
        /// Check if this strategy can handle the element.
        /// Returns true for FamilyInstance elements in the "Duct Accessories" category.
        /// </summary>
        public bool CanHandle(Element element)
        {
            if (element is not FamilyInstance fi) return false;
            var category = fi.Category;
            if (category == null) return false;

            // Check for Duct Accessories category (contains dampers)
            return category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory;
        }

        /// <summary>
        /// Extract parameters from a damper element.
        /// </summary>
        public ElementParameterSnapshot Extract(Element element)
        {
            var snapshot = new ElementParameterSnapshot();
            if (element is not FamilyInstance fi) return snapshot;

            // Capture all parameters in single pass
            ExtractAllParameters(fi, snapshot);

            // Extract prioritized dimensions
            ExtractDimensions(fi, snapshot);

            // Extract type/family info for MSFD detection
            ExtractTypeInfo(fi, snapshot);

            // Detect if damper is Standard (not MSFD/Motorized)
            snapshot.IsStandardDamper = !IsMotorizedDamper(snapshot.TypeName, snapshot.FamilyName);

            // Extract connector info for asymmetric clearance
            ExtractConnectorInfo(fi, snapshot);

            // Extract level info
            ExtractLevelInfo(fi, snapshot);

            return snapshot;
        }

        #region Extraction Methods

        private void ExtractAllParameters(FamilyInstance fi, ElementParameterSnapshot snapshot)
        {
            foreach (Parameter param in fi.Parameters)
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

        private void ExtractDimensions(FamilyInstance fi, ElementParameterSnapshot snapshot)
        {
            // Width: Main priority first
            snapshot.Width = GetFirstMatchingParameter(fi, WidthParametersMain)
                          ?? GetFirstMatchingParameter(fi, WidthParametersFallback);

            // Height: Main priority first
            snapshot.Height = GetFirstMatchingParameter(fi, HeightParametersMain)
                           ?? GetFirstMatchingParameter(fi, HeightParametersFallback);

            // Diameter (for round dampers)
            snapshot.OuterDiameter = GetFirstMatchingParameter(fi, DiameterParameters);

            // Also check built-in diameter parameter
            if (snapshot.OuterDiameter == null)
            {
                var diamParam = fi.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                if (diamParam != null && diamParam.HasValue)
                {
                    snapshot.OuterDiameter = diamParam.AsDouble();
                }
            }

            // Extract Size parameter value (string)
            var sizeParam = fi.LookupParameter("Size");
            if (sizeParam != null && sizeParam.HasValue)
            {
                snapshot.SizeParameterValue = sizeParam.StorageType == StorageType.String
                    ? sizeParam.AsString()
                    : sizeParam.AsValueString();
            }
        }

        private void ExtractTypeInfo(FamilyInstance fi, ElementParameterSnapshot snapshot)
        {
            // Get type name
            var elementType = fi.Document.GetElement(fi.GetTypeId());
            snapshot.TypeName = elementType?.Name;

            // Get family name
            snapshot.FamilyName = fi.Symbol?.Family?.Name;
        }

        private bool IsMotorizedDamper(string? typeName, string? familyName)
        {
            // Check type name for motorized keywords
            if (!string.IsNullOrEmpty(typeName))
            {
                foreach (var keyword in MotorizedKeywords)
                {
                    if (typeName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
            }

            // Check family name for motorized keywords
            if (!string.IsNullOrEmpty(familyName))
            {
                foreach (var keyword in MotorizedKeywords)
                {
                    if (familyName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
            }

            return false;
        }

        private void ExtractConnectorInfo(FamilyInstance fi, ElementParameterSnapshot snapshot)
        {
            try
            {
                // Access MEPModel via property, not cast
                var mepModel = fi.MEPModel;
                var connectorManager = mepModel?.ConnectorManager;

                if (connectorManager == null)
                {
                    snapshot.HasMepConnector = false;
                    return;
                }

                // Find first connected connector
                foreach (Connector connector in connectorManager.Connectors)
                {
                    if (connector.IsConnected)
                    {
                        snapshot.HasMepConnector = true;

                        // Determine connector side based on direction
                        var direction = connector.CoordinateSystem.BasisZ;
                        snapshot.ConnectorSide = DetermineConnectorSide(direction);
                        return;
                    }
                }

                snapshot.HasMepConnector = false;
            }
            catch
            {
                snapshot.HasMepConnector = false;
            }
        }

        private string DetermineConnectorSide(XYZ direction)
        {
            // Determine which side the connector is on based on direction vector
            double absX = Math.Abs(direction.X);
            double absY = Math.Abs(direction.Y);
            double absZ = Math.Abs(direction.Z);

            if (absY >= absX && absY >= absZ)
            {
                return direction.Y > 0 ? "+Y" : "-Y";
            }
            else if (absX >= absY && absX >= absZ)
            {
                return direction.X > 0 ? "+X" : "-X";
            }
            else
            {
                return direction.Z > 0 ? "+Z" : "-Z";
            }
        }

        private void ExtractLevelInfo(FamilyInstance fi, ElementParameterSnapshot snapshot)
        {
            // Try level parameters in priority order
            foreach (var paramName in LevelParameters)
            {
                var param = fi.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    if (param.StorageType == StorageType.ElementId)
                    {
                        var levelId = param.AsElementId();
                        if (levelId != null && levelId != ElementId.InvalidElementId)
                        {
                            var level = fi.Document.GetElement(levelId) as Level;
                            if (level != null)
                            {
                                snapshot.LevelName = level.Name;
                                snapshot.LevelElevation = level.Elevation;
                                return;
                            }
                        }
                    }
                    else if (param.StorageType == StorageType.String)
                    {
                        snapshot.LevelName = param.AsString();
                        return;
                    }
                }
            }

            // Fallback: try to get level from instance
            if (fi.LevelId != null && fi.LevelId != ElementId.InvalidElementId)
            {
                var level = fi.Document.GetElement(fi.LevelId) as Level;
                if (level != null)
                {
                    snapshot.LevelName = level.Name;
                    snapshot.LevelElevation = level.Elevation;
                }
            }
        }

        private double? GetFirstMatchingParameter(FamilyInstance fi, string[] paramNames)
        {
            foreach (var name in paramNames)
            {
                var param = fi.LookupParameter(name);
                if (param != null && param.HasValue && param.StorageType == StorageType.Double)
                {
                    return param.AsDouble();
                }
            }
            return null;
        }

        #endregion
    }
}
