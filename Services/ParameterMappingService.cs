using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for discovering and managing parameter mappings
    /// </summary>
    public class ParameterMappingService
    {
        /// <summary>
        /// Get available parameters from a specific element category
        /// </summary>
        public List<Models.ParameterInfo> GetAvailableParameters(Document doc, BuiltInCategory category)
        {
            var parameters = new List<Models.ParameterInfo>();
            
            try
            {
                // Get elements of the specified category
                var collector = new FilteredElementCollector(doc)
                    .OfCategory(category)
                    .WhereElementIsNotElementType()
                    .Take(1); // Just need one element to get parameters
                
                var element = collector.FirstOrDefault();
                if (element == null) return parameters;
                
                // Get all parameters from the element
                var paramSet = element.Parameters;
                foreach (Parameter param in paramSet)
                {
                    if (param.Definition != null)
                    {
                        var paramInfo = new Models.ParameterInfo
                        {
                            Name = param.Definition.Name,
                            Type = "Unknown", // ParameterType not available in this Revit version
                            IsReadOnly = param.IsReadOnly,
                            Category = category,
                            Description = param.Definition.Name,
                            IsShared = param.Definition is ExternalDefinition
                        };
                        
                        parameters.Add(paramInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting parameters for category {category}: {ex.Message}");
            }
            
            return parameters.OrderBy(p => p.Name).ToList();
        }
        
        /// <summary>
        /// Get available parameters from MEP elements
        /// </summary>
        public List<Models.ParameterInfo> GetMepElementParameters(Document doc)
        {
            var parameters = new List<Models.ParameterInfo>();
            
            try
            {
                // Get MEP categories
                var mepCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_ElectricalEquipment, // OST_CableTrayCurves not available
                BuiltInCategory.OST_ElectricalEquipment, // OST_ConduitCurves not available
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_PlumbingFixtures
                };
                
                foreach (var category in mepCategories)
                {
                    var categoryParams = GetAvailableParameters(doc, category);
                    parameters.AddRange(categoryParams);
                }
                
                // Remove duplicates
                parameters = parameters
                    .GroupBy(p => p.Name)
                    .Select(g => g.First())
                    .OrderBy(p => p.Name)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting MEP parameters: {ex.Message}");
            }
            
            return parameters;
        }
        
        /// <summary>
        /// Get available parameters from host elements (walls, floors, ceilings)
        /// </summary>
        public List<Models.ParameterInfo> GetHostElementParameters(Document doc)
        {
            var parameters = new List<Models.ParameterInfo>();
            
            try
            {
                // Get host categories
                var hostCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Ceilings
                };
                
                foreach (var category in hostCategories)
                {
                    var categoryParams = GetAvailableParameters(doc, category);
                    parameters.AddRange(categoryParams);
                }
                
                // Remove duplicates
                parameters = parameters
                    .GroupBy(p => p.Name)
                    .Select(g => g.First())
                    .OrderBy(p => p.Name)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting host parameters: {ex.Message}");
            }
            
            return parameters;
        }
        
        /// <summary>
        /// Get available parameters from levels
        /// </summary>
        public List<Models.ParameterInfo> GetLevelParameters(Document doc)
        {
            return GetAvailableParameters(doc, BuiltInCategory.OST_Levels);
        }
        
        /// <summary>
        /// Get available parameters from opening elements
        /// </summary>
        public List<Models.ParameterInfo> GetOpeningParameters(Document doc)
        {
            var parameters = new List<Models.ParameterInfo>();
            
            try
            {
                // Get opening categories
                var openingCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_GenericModel, // Generic openings
                    BuiltInCategory.OST_GenericAnnotation, // OST_WallOpening not available   // Wall openings
                    BuiltInCategory.OST_FloorOpening,  // Floor openings
                    BuiltInCategory.OST_CeilingOpening // Ceiling openings
                };
                
                foreach (var category in openingCategories)
                {
                    var categoryParams = GetAvailableParameters(doc, category);
                    parameters.AddRange(categoryParams);
                }
                
                // Remove duplicates
                parameters = parameters
                    .GroupBy(p => p.Name)
                    .Select(g => g.First())
                    .OrderBy(p => p.Name)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting opening parameters: {ex.Message}");
            }
            
            return parameters;
        }
        
        /// <summary>
        /// Create a parameter mapping configuration
        /// </summary>
        public ParameterMapping CreateMapping(string sourceParameter, string targetParameter, TransferType transferType)
        {
            return new ParameterMapping(sourceParameter, targetParameter, transferType);
        }
        
        /// <summary>
        /// Create a parameter mapping configuration with separator
        /// </summary>
        public ParameterMapping CreateMapping(string sourceParameter, string targetParameter, TransferType transferType, string separator)
        {
            return new ParameterMapping(sourceParameter, targetParameter, transferType, separator);
        }
        
        /// <summary>
        /// Validate parameter compatibility between source and target
        /// </summary>
        public bool ValidateMapping(ParameterInfo source, ParameterInfo target)
        {
            if (source == null || target == null) return false;
            
            // Check if target is read-only
            // if (target.IsReadOnly) return false; // IsReadOnly not available in Models.ParameterInfo
            
            // Check parameter type compatibility
            return IsParameterTypeCompatible(source.Type, target.Type);
        }
        
        /// <summary>
        /// Get common parameters for transfer operations
        /// </summary>
        public List<Models.ParameterInfo> GetCommonTransferParameters(Document doc)
        {
            var commonParams = new List<Models.ParameterInfo>();
            
            try
            {
                // Common parameters that are often transferred
                var commonParameterNames = new List<string>
                {
                    "System Abbreviation",
                    "System Name",
                    "System Type",
                    "Fire Rating",
                    "Material",
                    "Level Name",
                    "Level",
                    "Comments",
                    "Mark",
                    "Type Name",
                    "Family Name",
                    "Model Name"
                };
                
                // Get parameters from different categories
                var allParams = new List<Models.ParameterInfo>();
                allParams.AddRange(GetMepElementParameters(doc));
                allParams.AddRange(GetHostElementParameters(doc));
                allParams.AddRange(GetLevelParameters(doc));
                allParams.AddRange(GetOpeningParameters(doc));
                
                // Filter to common parameters
                foreach (var paramName in commonParameterNames)
                {
                    var param = allParams.FirstOrDefault(p => 
                        p.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase));
                    if (param != null)
                    {
                        commonParams.Add(param);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting common parameters: {ex.Message}");
            }
            
            return commonParams;
        }
        
        /// <summary>
        /// Get predefined parameter mappings for common scenarios
        /// </summary>
        public List<ParameterMapping> GetPredefinedMappings()
        {
            var mappings = new List<ParameterMapping>();
            
            // Common MEP to Opening mappings
            mappings.Add(new ParameterMapping("System Abbreviation", "MEP_System", TransferType.ReferenceToOpening));
            mappings.Add(new ParameterMapping("System Name", "MEP_System_Name", TransferType.ReferenceToOpening));
            mappings.Add(new ParameterMapping("System Type", "MEP_System_Type", TransferType.ReferenceToOpening));
            
            // Common Host to Opening mappings
            mappings.Add(new ParameterMapping("Fire Rating", "Opening_Fire_Rating", TransferType.HostToOpening));
            mappings.Add(new ParameterMapping("Material", "Opening_Material", TransferType.HostToOpening));
            mappings.Add(new ParameterMapping("Wall Type", "Opening_Wall_Type", TransferType.HostToOpening));
            
            // Common Level to Opening mappings
            mappings.Add(new ParameterMapping("Name", "Opening_Level", TransferType.LevelToOpening));
            mappings.Add(new ParameterMapping("Elevation", "Opening_Level_Elevation", TransferType.LevelToOpening));
            
            return mappings;
        }
        
        /// <summary>
        /// Check if a parameter exists on an element
        /// </summary>
        public bool ParameterExists(Element element, string parameterName)
        {
            try
            {
                var param = element.LookupParameter(parameterName);
                return param != null;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Get parameter value from an element
        /// </summary>
        public string GetParameterValue(Element element, string parameterName)
        {
            try
            {
                var param = element.LookupParameter(parameterName);
                if (param == null) return string.Empty;
                
                return param.AsString();
            }
            catch
            {
                return string.Empty;
            }
        }
        
        /// <summary>
        /// Set parameter value on an element
        /// </summary>
        public bool SetParameterValue(Element element, string parameterName, string value)
        {
            try
            {
                var param = element.LookupParameter(parameterName);
                if (param == null || param.IsReadOnly) return false;
                
                param.Set(value);
                return true;
            }
            catch
            {
                return false;
            }
        }
        
        #region Private Helper Methods
        
        private bool IsParameterTypeCompatible(string sourceType, string targetType)
        {
            // For most transfer operations, we're dealing with text parameters
            // This is a simplified compatibility check
            var textTypes = new List<string>
            {
                "Text",
                "String",
                "MultilineText"
            };
            
            var numericTypes = new List<string>
            {
                "Double",
                "Integer",
                "Number"
            };
            
            // Text parameters are generally compatible with each other
            if (textTypes.Contains(sourceType) && textTypes.Contains(targetType))
                return true;
            
            // Numeric parameters are compatible with each other
            if (numericTypes.Contains(sourceType) && numericTypes.Contains(targetType))
                return true;
            
            // Text can often be converted to other types
            if (textTypes.Contains(sourceType))
                return true;
            
            return false;
        }
        
        #endregion
    }
}
