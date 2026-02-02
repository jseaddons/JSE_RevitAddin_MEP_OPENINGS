using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for analyzing MEP element dimensions and calculating service sizes with clearance
    /// </summary>
    public class MepElementAnalysisService
    {
        private double _defaultClearance = 50.0; // Default clearance in mm
        private string _clearanceSuffix = "mm A.SPACE";
        
        /// <summary>
        /// Get MEP element dimensions
        /// </summary>
        public MepElementDimensions GetElementDimensions(Element mepElement)
        {
            var dimensions = new MepElementDimensions();
            
            try
            {
                // Get element type
                dimensions.ElementType = GetElementType(mepElement);
                
                // Try to get dimensions based on element type
                if (mepElement is MEPCurve mepCurve)
                {
                    GetMepCurveDimensions(mepCurve, dimensions);
                }
                else if (mepElement is FamilyInstance familyInstance)
                {
                    GetFamilyInstanceDimensions(familyInstance, dimensions);
                }
                else
                {
                    // Fallback to parameter-based dimension extraction
                    GetParameterBasedDimensions(mepElement, dimensions);
                }
                
                // Generate dimension string
                dimensions.DimensionString = GenerateDimensionString(dimensions);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting element dimensions: {ex.Message}");
                dimensions.ElementType = "Unknown";
                dimensions.DimensionString = "Unknown";
            }
            
            return dimensions;
        }
        
        /// <summary>
        /// Calculate total service size including clearance using direct parameters
        /// </summary>
        public ServiceSizeCalculation CalculateServiceSize(List<Element> mepElements, double clearance = 0)
        {
            var calculation = new ServiceSizeCalculation();

            try
            {
                if (clearance <= 0)
                    clearance = _defaultClearance;
                
                calculation.AnnularSpace = clearance;
                calculation.ServiceDimensions = new List<MepElementDimensions>();
                
                // Get dimensions directly from MEP element parameters
                var mepDimensions = new List<string>();
                
                foreach (var element in mepElements)
                {
                    // Get size directly from element parameters
                    var size = GetElementSizeFromParameters(element);
                    if (!string.IsNullOrEmpty(size))
                    {
                        mepDimensions.Add(size);
                    }
                }
                
                // Get opening size directly from opening parameters
                var openingSize = GetOpeningSizeFromParameters(mepElements);
                
                // Generate simple calculation string: MEP_SIZE + CLEARANCE = OPENING_SIZE
                var mepSizeStr = string.Join(" + ", mepDimensions);
                calculation.CalculationString = $"{mepSizeStr} +{clearance}{_clearanceSuffix} ={openingSize}";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error calculating service size: {ex.Message}");
                calculation.CalculationString = "Error calculating service size";
            }
            
            return calculation;
        }
        
        /// <summary>
        /// Analyze cluster opening with multiple MEP elements
        /// </summary>
        public ClusterOpeningAnalysis AnalyzeClusterOpening(ElementId openingId, List<Element> mepElements)
        {
            var analysis = new ClusterOpeningAnalysis
            {
                OpeningId = openingId,
                MepElements = mepElements,
                ServiceTypes = new List<string>(),
                TotalServiceSize = CalculateServiceSize(mepElements)
            };
            
            try
            {
                // Get service types from elements
                var abbreviationService = new ServiceTypeAbbreviationService();
                foreach (var element in mepElements)
                {
                    var serviceType = abbreviationService.GetAbbreviatedServiceTypeFromElement(element);
                    if (!string.IsNullOrEmpty(serviceType) && !analysis.ServiceTypes.Contains(serviceType))
                    {
                        analysis.ServiceTypes.Add(serviceType);
                    }
                }
                
                // Combine service types
                analysis.CombinedServiceType = string.Join(" & ", analysis.ServiceTypes);
                
                // Get opening center from FFL (placeholder - would need host element analysis)
                analysis.OpeningCenterFromFFL = 0; // TODO: Implement with HostElementAnalysisService
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error analyzing cluster opening: {ex.Message}");
            }
            
            return analysis;
        }
        
        /// <summary>
        /// Set clearance parameters
        /// </summary>
        public void SetClearanceParameters(double clearance, string suffix = "mm A.SPACE")
        {
            _defaultClearance = clearance;
            _clearanceSuffix = suffix;
        }
        
        /// <summary>
        /// Get current clearance parameters
        /// </summary>
        public (double clearance, string suffix) GetClearanceParameters()
        {
            return (_defaultClearance, _clearanceSuffix);
        }
        
        #region Private Helper Methods
        
        private string GetElementType(Element element)
        {
            try
            {
                // Try to get from category
                var category = element.Category?.Name;
                if (!string.IsNullOrEmpty(category))
                {
                    return category;
                }
                
                // Try to get from family name
                var familyName = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsString();
                if (!string.IsNullOrEmpty(familyName))
                {
                    return familyName;
                }
                
                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }
        
        private void GetMepCurveDimensions(MEPCurve mepCurve, MepElementDimensions dimensions)
        {
            try
            {
                // Get diameter for pipes and conduits
                var diameterParam = mepCurve.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
                if (diameterParam != null)
                {
                    dimensions.Diameter = diameterParam.AsDouble() * 304.8; // Convert to mm
                    dimensions.Width = dimensions.Diameter;
                    dimensions.Height = dimensions.Diameter;
                    return;
                }
                
                // Get width and height for ducts and cable trays
                var widthParam = mepCurve.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                var heightParam = mepCurve.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                
                if (widthParam != null)
                {
                    dimensions.Width = widthParam.AsDouble() * 304.8; // Convert to mm
                }
                
                if (heightParam != null)
                {
                    dimensions.Height = heightParam.AsDouble() * 304.8; // Convert to mm
                }
                
                // If no specific parameters found, try generic parameters
                if (dimensions.Width <= 0 || dimensions.Height <= 0)
                {
                    GetGenericDimensions(mepCurve, dimensions);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting MEP curve dimensions: {ex.Message}");
            }
        }
        
        private void GetFamilyInstanceDimensions(FamilyInstance familyInstance, MepElementDimensions dimensions)
        {
            try
            {
                // Try to get dimensions from family instance parameters
                GetGenericDimensions(familyInstance, dimensions);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting family instance dimensions: {ex.Message}");
            }
        }
        
        private void GetParameterBasedDimensions(Element element, MepElementDimensions dimensions)
        {
            try
            {
                GetGenericDimensions(element, dimensions);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting parameter-based dimensions: {ex.Message}");
            }
        }
        
        private void GetGenericDimensions(Element element, MepElementDimensions dimensions)
        {
            try
            {
                // Common dimension parameter names to try
                var dimensionParams = new List<string>
                {
                    "Width", "Height", "Outside Diameter",
                    "Outer Diameter", "Inner Diameter", "Diameter",
                    "Size", "Dimension",
                    "Nominal Size", "Actual Size"
                };
                
                foreach (var paramName in dimensionParams)
                {
                    var param = element.LookupParameter(paramName);
                    if (param != null)
                    {
                        var value = param.AsDouble();
                        if (value > 0)
                        {
                            // Convert to mm if needed (assuming Revit units)
                            value = value * 304.8;
                            
                            if (paramName.Contains("Width") || paramName.Contains("Size"))
                            {
                                dimensions.Width = Math.Max(dimensions.Width, value);
                            }
                            else if (paramName.Contains("Height"))
                            {
                                dimensions.Height = Math.Max(dimensions.Height, value);
                            }
                            else if (paramName.Contains("Diameter"))
                            {
                                dimensions.Diameter = Math.Max(dimensions.Diameter, value);
                                if (dimensions.Width <= 0) dimensions.Width = dimensions.Diameter;
                                if (dimensions.Height <= 0) dimensions.Height = dimensions.Diameter;
                            }
                        }
                    }
                }
                
                // If still no dimensions found, try bounding box
                if (dimensions.Width <= 0 && dimensions.Height <= 0)
                {
                    GetBoundingBoxDimensions(element, dimensions);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting generic dimensions: {ex.Message}");
            }
        }
        
        private void GetBoundingBoxDimensions(Element element, MepElementDimensions dimensions)
        {
            try
            {
                var geom = element.get_Geometry(Helpers.GeometryOptionsFactory.CreateIntersectionOptions());
                if (geom != null)
                {
                    var bbox = geom.GetBoundingBox();
                    if (bbox != null)
                    {
                        var size = bbox.Max - bbox.Min;
                        dimensions.Width = Math.Abs(size.X) * 304.8; // Convert to mm
                        dimensions.Height = Math.Abs(size.Y) * 304.8; // Convert to mm
                        
                        // Use the larger dimension as diameter for circular elements
                        dimensions.Diameter = Math.Max(dimensions.Width, dimensions.Height);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting bounding box dimensions: {ex.Message}");
            }
        }
        
        private string GenerateDimensionString(MepElementDimensions dimensions)
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
        
        /// <summary>
        /// Get element size directly from parameters
        /// </summary>
        private string GetElementSizeFromParameters(Element element)
        {
            try
            {
                // Try common size parameter names
                var sizeParams = new List<string>
                {
                    "Size",
                    "Nominal Size", 
                    "Actual Size",
                    "Outside Diameter",
                    "Outer Diameter",
                    "Inner Diameter",
                    "Width",
                    "Height",
                    "Diameter"
                };
                
                foreach (var paramName in sizeParams)
                {
                    var param = element.LookupParameter(paramName);
                    if (param != null)
                    {
                        var value = param.AsString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            return value;
                        }
                    }
                }
                
                // Try to get from family and type name
                var familyType = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsString();
                if (!string.IsNullOrEmpty(familyType))
                {
                    // Extract size from family name if it contains dimensions
                    if (!string.IsNullOrEmpty(familyType) && (familyType.Contains("x") || familyType.Contains("mm") || familyType.Contains("in")))
                    {
                        return familyType;
                    }
                }
                
                return "Unknown";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting element size from parameters: {ex.Message}");
                return "Unknown";
            }
        }
        
        /// <summary>
        /// Get opening size directly from opening parameters
        /// </summary>
        private string GetOpeningSizeFromParameters(List<Element> mepElements)
        {
            try
            {
                // For now, return a placeholder - this would be calculated based on MEP elements + clearance
                // In a real implementation, this would get the actual opening size from the opening element
                return "Opening_Size"; // Placeholder
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting opening size from parameters: {ex.Message}");
                return "Unknown";
            }
        }
        
        #endregion
    }
}
