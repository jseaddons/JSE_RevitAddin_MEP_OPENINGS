using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for extracting opening parameters with user-configurable parameter names
    /// </summary>
    public class OpeningParameterExtractionService
    {
        private OpeningParameterConfiguration _config;
        
        public OpeningParameterExtractionService()
        {
            _config = new OpeningParameterConfiguration();
        }
        
        public OpeningParameterExtractionService(OpeningParameterConfiguration config)
        {
            _config = config ?? new OpeningParameterConfiguration();
        }
        
        /// <summary>
        /// Set parameter configuration
        /// </summary>
        public void SetConfiguration(OpeningParameterConfiguration config)
        {
            _config = config ?? new OpeningParameterConfiguration();
        }
        
        /// <summary>
        /// Get opening dimensions (Width, Height, Diameter)
        /// </summary>
        public OpeningDimensions GetOpeningDimensions(Element opening)
        {
            var dimensions = new OpeningDimensions();
            
            try
            {
                // Get width
                dimensions.Width = GetParameterValue(opening, _config.WidthParameterName);
                
                // Get height
                dimensions.Height = GetParameterValue(opening, _config.HeightParameterName);
                
                // Get diameter
                dimensions.Diameter = GetParameterValue(opening, _config.DiameterParameterName);
                
                // Generate dimension string
                dimensions.DimensionString = GenerateDimensionString(dimensions);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting opening dimensions: {ex.Message}");
            }
            
            return dimensions;
        }
        
        /// <summary>
        /// Get opening center from FFL (elevation)
        /// </summary>
        public double GetOpeningCenterFromFFL(Element opening)
        {
            try
            {
                return GetParameterValue(opening, _config.CenterFromFFLParameterName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting opening center from FFL: {ex.Message}");
                return 0.0;
            }
        }
        
        /// <summary>
        /// Get opening level information
        /// </summary>
        public OpeningLevelInfo GetOpeningLevelInfo(Element opening, Document doc)
        {
            var levelInfo = new OpeningLevelInfo();
            
            try
            {
                // Get level from opening
                var levelParam = opening.LookupParameter(_config.LevelParameterName);
                if (levelParam != null && levelParam.StorageType == StorageType.ElementId)
                {
                    var levelId = levelParam.AsElementId();
                    if (levelId != ElementId.InvalidElementId)
                    {
                        var level = doc.GetElement(levelId) as Level;
                        if (level != null)
                        {
                            levelInfo.LevelName = level.Name;
                            levelInfo.LevelElevation = level.Elevation;
                            levelInfo.LevelId = levelId;
                        }
                    }
                }
                
                // Get elevation from opening
                levelInfo.CenterFromFFL = GetParameterValue(opening, _config.CenterFromFFLParameterName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting opening level info: {ex.Message}");
            }
            
            return levelInfo;
        }
        
        /// <summary>
        /// Get ceiling level from FFL
        /// </summary>
        public double GetCeilingLevelFromFFL(Element opening)
        {
            try
            {
                return GetParameterValue(opening, _config.CeilingLevelFromFFLParameterName);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting ceiling level from FFL: {ex.Message}");
                return 0.0;
            }
        }
        
        /// <summary>
        /// Get mark value from opening (uses existing MarkParameterAddValue system)
        /// </summary>
        public string GetOpeningMarkValue(Element opening)
        {
            try
            {
                // Use the Mark parameter (which is what MarkParameterAddValue command uses)
                var markParam = opening.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                if (markParam != null)
                {
                    return markParam.AsString() ?? string.Empty;
                }
                
                return string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting opening mark value: {ex.Message}");
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Set mark value for opening (uses existing MarkParameterAddValue system)
        /// </summary>
        public bool SetOpeningMarkValue(Element opening, string markValue)
        {
            try
            {
                // Use the Mark parameter (which is what MarkParameterAddValue command uses)
                var markParam = opening.get_Parameter(BuiltInParameter.ALL_MODEL_MARK);
                if (markParam != null && !markParam.IsReadOnly)
                {
                    markParam.Set(markValue);
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting opening mark value: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Generate mark value based on opening family type (uses existing MarkParameterAddValue system)
        /// </summary>
        public string GenerateMarkValue(Element opening, string prefix = "", int index = 1)
        {
            try
            {
                string familyName = opening.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsString() ?? string.Empty;
                string typeName = opening.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsString() ?? string.Empty;
                
                // Use the same logic as MarkParameterAddValue command
                if (familyName.Contains("Pipe", StringComparison.OrdinalIgnoreCase))
                {
                    return $"{prefix}PO-{index:000}";
                }
                else if (familyName.Contains("Duct", StringComparison.OrdinalIgnoreCase))
                {
                    return $"{prefix}DO-{index:000}";
                }
                else if (familyName.Contains("Damper", StringComparison.OrdinalIgnoreCase))
                {
                    return $"{prefix}DA-{index:000}";
                }
                else if (familyName.Contains("CableTray", StringComparison.OrdinalIgnoreCase))
                {
                    return $"{prefix}CT-{index:000}";
                }
                else if (familyName.StartsWith("Cluster", StringComparison.OrdinalIgnoreCase) ||
                         familyName.IndexOf("Cluster", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         typeName.EndsWith("Rect", StringComparison.OrdinalIgnoreCase))
                {
                    return $"{prefix}CO-{index:000}";
                }
                else
                {
                    // Default fallback
                    return $"{prefix}O-{index:000}";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error generating mark value: {ex.Message}");
                return $"{prefix}O-{index:000}";
            }
        }
        
        #region Private Helper Methods
        
        private double GetParameterValue(Element element, string parameterName)
        {
            try
            {
                if (string.IsNullOrEmpty(parameterName)) return 0.0;
                
                var param = element.LookupParameter(parameterName);
                if (param != null)
                {
                    if (param.StorageType == StorageType.Double)
                    {
                        return param.AsDouble();
                    }
                    else if (param.StorageType == StorageType.String)
                    {
                        var stringValue = param.AsString();
                        if (double.TryParse(stringValue, out double result))
                        {
                            return result;
                        }
                    }
                }
                
                return 0.0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting parameter value for {parameterName}: {ex.Message}");
                return 0.0;
            }
        }
        
        private string GenerateDimensionString(OpeningDimensions dimensions)
        {
            try
            {
                if (dimensions.Diameter > 0 && Math.Abs(dimensions.Width - dimensions.Diameter) < 1)
                {
                    // Circular element
                    return $"{dimensions.Diameter:F0}";
                }
                else if (dimensions.Width > 0 && dimensions.Height > 0)
                {
                    // Rectangular element
                    return $"{dimensions.Width:F0}x{dimensions.Height:F0}";
                }
                else if (dimensions.Width > 0)
                {
                    // Width only
                    return $"{dimensions.Width:F0}";
                }
                else if (dimensions.Height > 0)
                {
                    // Height only
                    return $"{dimensions.Height:F0}";
                }
                else
                {
                    return "Unknown";
                }
            }
            catch
            {
                return "Unknown";
            }
        }
        
        #endregion
    }
    
    /// <summary>
    /// Configuration for opening parameter names
    /// </summary>
    public class OpeningParameterConfiguration
    {
        // Dimension parameters
        public string WidthParameterName { get; set; } = "Width";
        public string HeightParameterName { get; set; } = "Height";
        public string DiameterParameterName { get; set; } = "Outside Diameter";
        
        // Level and elevation parameters
        public string LevelParameterName { get; set; } = "Level";
        public string CenterFromFFLParameterName { get; set; } = "Center From FFL";
        public string CeilingLevelFromFFLParameterName { get; set; } = "Ceiling Level From FFL";
        
        // Note: Mark values are handled by the existing MarkParameterAddValue command system
        // No additional configuration needed for mark values
        
        public OpeningParameterConfiguration()
        {
        }
        
        public OpeningParameterConfiguration(
            string widthParam, 
            string heightParam, 
            string diameterParam,
            string levelParam,
            string centerFromFFLParam,
            string ceilingLevelFromFFLParam)
        {
            WidthParameterName = widthParam;
            HeightParameterName = heightParam;
            DiameterParameterName = diameterParam;
            LevelParameterName = levelParam;
            CenterFromFFLParameterName = centerFromFFLParam;
            CeilingLevelFromFFLParameterName = ceilingLevelFromFFLParam;
        }
    }
    
    /// <summary>
    /// Opening dimensions data
    /// </summary>
    public class OpeningDimensions
    {
        public double Width { get; set; }
        public double Height { get; set; }
        public double Diameter { get; set; }
        public string DimensionString { get; set; } = string.Empty;
        
        public OpeningDimensions()
        {
        }
        
        public OpeningDimensions(double width, double height, double diameter)
        {
            Width = width;
            Height = height;
            Diameter = diameter;
        }
    }
    
    /// <summary>
    /// Opening level information
    /// </summary>
    public class OpeningLevelInfo
    {
        public string LevelName { get; set; } = string.Empty;
        public double LevelElevation { get; set; }
        public double CenterFromFFL { get; set; }
        public ElementId LevelId { get; set; } = ElementId.InvalidElementId;
        
        public OpeningLevelInfo()
        {
        }
        
        public OpeningLevelInfo(string levelName, double levelElevation, double centerFromFFL, ElementId levelId)
        {
            LevelName = levelName;
            LevelElevation = levelElevation;
            CenterFromFFL = centerFromFFL;
            LevelId = levelId;
        }
    }
}
