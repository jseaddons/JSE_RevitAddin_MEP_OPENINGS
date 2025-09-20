using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class ParameterInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public bool IsInstanceParameter { get; set; }
        public BuiltInParameter? BuiltInParameter { get; set; }
        public List<string> Values { get; set; } = new List<string>();
    }

    public class ParameterExtractionService
    {
        public List<ParameterInfo> GetParametersForCategory(Document document, BuiltInCategory category, bool includeInstanceParams = true, bool includeTypeParams = true)
        {
            var parameters = new List<ParameterInfo>();

            try
            {
                // Get all elements of the specified category
                var collector = new FilteredElementCollector(document)
                    .OfCategory(category)
                    .WhereElementIsNotElementType();

                var elements = collector.ToElements();
                if (elements.Count == 0) return parameters;

                // Sample a few elements to extract parameters
                var sampleElements = elements.Take(10).ToList();

                // Get parameters from instance elements
                if (includeInstanceParams)
                {
                    foreach (var element in sampleElements)
                    {
                        ExtractParametersFromElement(element, parameters, true);
                    }
                }

                // Get parameters from element types
                if (includeTypeParams)
                {
                    foreach (var element in sampleElements)
                    {
                        var elementType = document.GetElement(element.GetTypeId());
                        if (elementType != null)
                        {
                            ExtractParametersFromElement(elementType, parameters, false);
                        }
                    }
                }

                // Remove duplicates and sort
                parameters = parameters
                    .GroupBy(p => p.Name)
                    .Select(g => g.First())
                    .OrderBy(p => p.Name)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting parameters for category {category}: {ex.Message}");
            }

            return parameters;
        }

        public List<ParameterInfo> GetParametersForMepCategories(Document document, List<MepCategory> categories)
        {
            var allParameters = new List<ParameterInfo>();

            foreach (var category in categories)
            {
                var builtinCategories = GetBuiltInCategoriesForMepCategory(category);
                foreach (var builtinCategory in builtinCategories)
                {
                    var categoryParams = GetParametersForCategory(document, builtinCategory);
                    allParameters.AddRange(categoryParams);
                }
            }

            // Remove duplicates and sort
            return allParameters
                .GroupBy(p => p.Name)
                .Select(g => g.First())
                .OrderBy(p => p.Name)
                .ToList();
        }

        private void ExtractParametersFromElement(Element element, List<ParameterInfo> parameters, bool isInstance)
        {
            try
            {
                foreach (Parameter param in element.Parameters)
                {
                    if (param == null || string.IsNullOrEmpty(param.Definition?.Name)) continue;

                    var paramInfo = new ParameterInfo
                    {
                        Name = param.Definition?.Name ?? "Unknown",
                        Type = param.StorageType.ToString(),
                        IsInstanceParameter = isInstance,
                        BuiltInParameter = param.Id.IntegerValue < 0 ?
                            (BuiltInParameter)param.Id.IntegerValue : null
                    };

                    // Try to get some sample values
                    if (param.HasValue)
                    {
                        try
                        {
                            string value = GetParameterValueAsString(param);
                            if (!string.IsNullOrEmpty(value) && !paramInfo.Values.Contains(value))
                            {
                                paramInfo.Values.Add(value);
                            }
                        }
                        catch
                        {
                            // Ignore value extraction errors
                        }
                    }

                    parameters.Add(paramInfo);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting parameters from element: {ex.Message}");
            }
        }

        private string GetParameterValueAsString(Parameter param)
        {
            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return param.AsString() ?? string.Empty;
                    case StorageType.Integer:
                        return param.AsInteger().ToString();
                    case StorageType.Double:
                        return param.AsDouble().ToString();
                    case StorageType.ElementId:
                        var elemId = param.AsElementId();
                        return elemId.IntegerValue.ToString();
                    default:
                        return string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        private List<BuiltInCategory> GetBuiltInCategoriesForMepCategory(MepCategory category)
        {
            return category switch
            {
                MepCategory.Pipes => new List<BuiltInCategory> { BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting },
                MepCategory.Ducts => new List<BuiltInCategory> { BuiltInCategory.OST_DuctCurves },
                MepCategory.DuctAccessories => new List<BuiltInCategory> { BuiltInCategory.OST_DuctAccessory },
                MepCategory.DuctFittings => new List<BuiltInCategory> { BuiltInCategory.OST_DuctFitting },
                MepCategory.CableTrays => new List<BuiltInCategory> { BuiltInCategory.OST_CableTray },
                MepCategory.Conduits => new List<BuiltInCategory> { BuiltInCategory.OST_Conduit },
                _ => new List<BuiltInCategory>()
            };
        }

        public List<string> GetParameterNamesForDisplay(List<ParameterInfo> parameters)
        {
            return parameters.Select(p => p.Name).ToList();
        }

        public List<string> GetParameterValuesForParameter(List<ParameterInfo> parameters, string parameterName)
        {
            var param = parameters.FirstOrDefault(p => p.Name == parameterName);
            return param?.Values ?? new List<string>();
        }
    }
}
