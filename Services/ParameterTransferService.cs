using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for transferring parameters from various sources to openings
    /// </summary>
    public class ParameterTransferService
    {
        private readonly ParameterRenamingService _renamingService;
        private readonly ParameterMappingService _mappingService;
        private readonly ServiceTypeAbbreviationService _abbreviationService;
        private readonly MepElementAnalysisService _mepAnalysisService;
        
        public ParameterTransferService()
        {
            _renamingService = new ParameterRenamingService();
            _mappingService = new ParameterMappingService();
            _abbreviationService = new ServiceTypeAbbreviationService();
            _mepAnalysisService = new MepElementAnalysisService();
        }
        
        /// <summary>
        /// Transfer parameters from MEP elements (references) to openings
        /// </summary>
        public ParameterTransferResult TransferFromReferenceElements(
            Document doc, 
            List<ElementId> openingIds, 
            ParameterMapping mapping)
        {
            var result = new ParameterTransferResult();
            
            try
            {
                if (!mapping.IsEnabled)
                {
                    result.Success = true;
                    result.Message = "Mapping is disabled, skipping transfer.";
                    return result;
                }
                
                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();
                
                using (var transaction = new Transaction(doc, "Transfer Parameters from Reference Elements"))
                {
                    transaction.Start();
                    
                    foreach (var openingId in openingIds)
                    {
                        try
                        {
                            var opening = doc.GetElement(openingId);
                            if (opening == null) continue;
                            
                            // Get MEP elements that intersect with this opening
                            var mepElements = GetMepElementsInOpening(doc, opening);
                            
                            if (mepElements.Count == 0)
                            {
                                result.Warnings.Add($"No MEP elements found for opening {openingId}");
                                continue;
                            }
                            
                            // Transfer parameter from first MEP element (or combine if multiple)
                            var transferSuccess = TransferParameterFromElements(
                                doc, opening, mepElements, mapping);
                            
                            if (transferSuccess)
                                transferredCount++;
                            else
                                failedCount++;
                        }
                        catch (Exception ex)
                        {
                            failedCount++;
                            errors.Add($"Error transferring to opening {openingId}: {ex.Message}");
                        }
                    }
                    
                    transaction.Commit();
                }
                
                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} parameters, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }
            
            return result;
        }
        
        /// <summary>
        /// Transfer parameters from host elements (walls, floors, ceilings) to openings
        /// </summary>
        public ParameterTransferResult TransferFromHostElements(
            Document doc, 
            List<ElementId> openingIds, 
            ParameterMapping mapping)
        {
            var result = new ParameterTransferResult();
            
            try
            {
                if (!mapping.IsEnabled)
                {
                    result.Success = true;
                    result.Message = "Mapping is disabled, skipping transfer.";
                    return result;
                }
                
                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();
                
                using (var transaction = new Transaction(doc, "Transfer Parameters from Host Elements"))
                {
                    transaction.Start();
                    
                    foreach (var openingId in openingIds)
                    {
                        try
                        {
                            var opening = doc.GetElement(openingId);
                            if (opening == null) continue;
                            
                            // Get host elements (walls, floors, ceilings) that contain this opening
                            var hostElements = GetHostElementsForOpening(doc, opening);
                            
                            if (hostElements.Count == 0)
                            {
                                result.Warnings.Add($"No host elements found for opening {openingId}");
                                continue;
                            }
                            
                            // Transfer parameter from host elements
                            var transferSuccess = TransferParameterFromElements(
                                doc, opening, hostElements, mapping);
                            
                            if (transferSuccess)
                                transferredCount++;
                            else
                                failedCount++;
                        }
                        catch (Exception ex)
                        {
                            failedCount++;
                            errors.Add($"Error transferring to opening {openingId}: {ex.Message}");
                        }
                    }
                    
                    transaction.Commit();
                }
                
                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} parameters, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }
            
            return result;
        }
        
        /// <summary>
        /// Transfer parameters from levels to openings
        /// </summary>
        public ParameterTransferResult TransferFromLevels(
            Document doc, 
            List<ElementId> openingIds, 
            ParameterMapping mapping)
        {
            var result = new ParameterTransferResult();
            
            try
            {
                if (!mapping.IsEnabled)
                {
                    result.Success = true;
                    result.Message = "Mapping is disabled, skipping transfer.";
                    return result;
                }
                
                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();
                
                using (var transaction = new Transaction(doc, "Transfer Parameters from Levels"))
                {
                    transaction.Start();
                    
                    foreach (var openingId in openingIds)
                    {
                        try
                        {
                            var opening = doc.GetElement(openingId);
                            if (opening == null) continue;
                            
                            // Get level for this opening
                            var level = GetLevelForOpening(doc, opening);
                            
                            if (level == null)
                            {
                                result.Warnings.Add($"No level found for opening {openingId}");
                                continue;
                            }
                            
                            // Transfer parameter from level
                            var transferSuccess = TransferParameterFromElement(
                                doc, opening, level, mapping);
                            
                            if (transferSuccess)
                                transferredCount++;
                            else
                                failedCount++;
                        }
                        catch (Exception ex)
                        {
                            failedCount++;
                            errors.Add($"Error transferring to opening {openingId}: {ex.Message}");
                        }
                    }
                    
                    transaction.Commit();
                }
                
                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} parameters, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }
            
            return result;
        }
        
        /// <summary>
        /// Transfer service size calculations with clearance to openings
        /// </summary>
        public ParameterTransferResult TransferServiceSizeCalculations(
            Document doc, 
            List<ElementId> openingIds, 
            string targetParameter,
            double clearance = 50.0)
        {
            var result = new ParameterTransferResult();
            
            try
            {
                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();
                
                using (var transaction = new Transaction(doc, "Transfer Service Size Calculations"))
                {
                    transaction.Start();
                    
                    foreach (var openingId in openingIds)
                    {
                        try
                        {
                            var opening = doc.GetElement(openingId);
                            if (opening == null) continue;
                            
                            // Get MEP elements that intersect with this opening
                            var mepElements = GetMepElementsInOpening(doc, opening);
                            
                            if (mepElements.Count == 0)
                            {
                                result.Warnings.Add($"No MEP elements found for opening {openingId}");
                                continue;
                            }
                            
                            // Calculate service size with clearance
                            var serviceSizeCalculation = _mepAnalysisService.CalculateServiceSize(mepElements, clearance);
                            
                            // Set parameter value
                            var param = opening.LookupParameter(targetParameter);
                            if (param != null && !param.IsReadOnly)
                            {
                                param.Set(serviceSizeCalculation.CalculationString);
                                transferredCount++;
                            }
                            else
                            {
                                failedCount++;
                                errors.Add($"Cannot set parameter {targetParameter} on opening {openingId}");
                            }
                        }
                        catch (Exception ex)
                        {
                            failedCount++;
                            errors.Add($"Error transferring service size calculation to opening {openingId}: {ex.Message}");
                        }
                    }
                    
                    transaction.Commit();
                }
                
                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} service size calculations, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Service size calculation transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }
            
            return result;
        }
        
        /// <summary>
        /// Transfer model names to openings
        /// </summary>
        public ParameterTransferResult TransferModelNames(
            Document doc, 
            List<ElementId> openingIds, 
            string targetParameter)
        {
            var result = new ParameterTransferResult();
            
            try
            {
                var transferredCount = 0;
                var failedCount = 0;
                var errors = new List<string>();
                
                using (var transaction = new Transaction(doc, "Transfer Model Names"))
                {
                    transaction.Start();
                    
                    foreach (var openingId in openingIds)
                    {
                        try
                        {
                            var opening = doc.GetElement(openingId);
                            if (opening == null) continue;
                            
                            // Get model name
                            var modelName = doc.Title;
                            
                            // Set parameter value
                            var param = opening.LookupParameter(targetParameter);
                            if (param != null && !param.IsReadOnly)
                            {
                                param.Set(modelName);
                                transferredCount++;
                            }
                            else
                            {
                                failedCount++;
                                errors.Add($"Cannot set parameter {targetParameter} on opening {openingId}");
                            }
                        }
                        catch (Exception ex)
                        {
                            failedCount++;
                            errors.Add($"Error transferring model name to opening {openingId}: {ex.Message}");
                        }
                    }
                    
                    transaction.Commit();
                }
                
                result.Success = true;
                result.TransferredCount = transferredCount;
                result.FailedCount = failedCount;
                result.Errors = errors;
                result.Message = $"Transferred {transferredCount} model names, {failedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }
            
            return result;
        }
        
        /// <summary>
        /// Execute complete parameter transfer configuration
        /// </summary>
        public ParameterTransferResult ExecuteTransferConfiguration(
            Document doc, 
            List<ElementId> openingIds, 
            ParameterTransferConfiguration config)
        {
            var result = new ParameterTransferResult();
            var allResults = new List<ParameterTransferResult>();
            
            try
            {
                // Execute each mapping
                foreach (var mapping in config.Mappings)
                {
                    ParameterTransferResult mappingResult = null;
                    
                    switch (mapping.TransferType)
                    {
                        case TransferType.ReferenceToOpening:
                            mappingResult = TransferFromReferenceElements(doc, openingIds, mapping);
                            break;
                        case TransferType.HostToOpening:
                            mappingResult = TransferFromHostElements(doc, openingIds, mapping);
                            break;
                        case TransferType.LevelToOpening:
                            mappingResult = TransferFromLevels(doc, openingIds, mapping);
                            break;
                    }
                    
                    if (mappingResult != null)
                        allResults.Add(mappingResult);
                }
                
                // Transfer model names if enabled
                if (config.TransferModelNames)
                {
                    var modelResult = TransferModelNames(doc, openingIds, config.ModelNameParameter);
                    allResults.Add(modelResult);
                }
                
                // Transfer service size calculations if enabled
                if (config.TransferServiceSizeCalculations)
                {
                    // Set clearance parameters
                    _mepAnalysisService.SetClearanceParameters(config.DefaultClearance, config.ClearanceSuffix);
                    
                    var serviceSizeResult = TransferServiceSizeCalculations(
                        doc, openingIds, config.ServiceSizeCalculationParameter, config.DefaultClearance);
                    allResults.Add(serviceSizeResult);
                }
                
                // Combine results
                result.Success = allResults.All(r => r.Success);
                result.TransferredCount = allResults.Sum(r => r.TransferredCount);
                result.FailedCount = allResults.Sum(r => r.FailedCount);
                result.Errors = allResults.SelectMany(r => r.Errors).ToList();
                result.Warnings = allResults.SelectMany(r => r.Warnings).ToList();
                result.Message = $"Transfer completed: {result.TransferredCount} successful, {result.FailedCount} failed.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = $"Configuration transfer failed: {ex.Message}";
                result.Errors.Add(ex.Message);
            }
            
            return result;
        }
        
        #region Private Helper Methods
        
        private List<Element> GetMepElementsInOpening(Document doc, Element opening)
        {
            var mepElements = new List<Element>();
            
            try
            {
                // Get MEP elements that intersect with the opening
                var collector = new FilteredElementCollector(doc)
                    .OfClass(typeof(MEPCurve))
                    .WhereElementIsNotElementType();
                
                foreach (Element mepElement in collector)
                {
                    if (ElementsIntersect(opening, mepElement))
                    {
                        mepElements.Add(mepElement);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but continue
                System.Diagnostics.Debug.WriteLine($"Error getting MEP elements: {ex.Message}");
            }
            
            return mepElements;
        }
        
        private List<Element> GetHostElementsForOpening(Document doc, Element opening)
        {
            var hostElements = new List<Element>();
            
            try
            {
                // Get host elements (walls, floors, ceilings) that contain this opening
                var hostCategories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Ceilings
                };
                
                var filter = new ElementMulticategoryFilter(hostCategories);
                var collector = new FilteredElementCollector(doc)
                    .WherePasses(filter)
                    .WhereElementIsNotElementType();
                
                foreach (Element hostElement in collector)
                {
                    if (ElementsIntersect(opening, hostElement))
                    {
                        hostElements.Add(hostElement);
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error but continue
                System.Diagnostics.Debug.WriteLine($"Error getting host elements: {ex.Message}");
            }
            
            return hostElements;
        }
        
        private Element GetLevelForOpening(Document doc, Element opening)
        {
            try
            {
                // Get level from opening's location
                var location = opening.Location as LocationPoint;
                if (location != null)
                {
                    var level = doc.GetElement(opening.LevelId) as Level;
                    return level;
                }
            }
            catch (Exception ex)
            {
                // Log error but continue
                System.Diagnostics.Debug.WriteLine($"Error getting level: {ex.Message}");
            }
            
            return null;
        }
        
        private bool TransferParameterFromElements(
            Document doc, 
            Element targetElement, 
            List<Element> sourceElements, 
            ParameterMapping mapping)
        {
            if (sourceElements.Count == 0) return false;
            
            if (sourceElements.Count == 1)
            {
                return TransferParameterFromElement(doc, targetElement, sourceElements[0], mapping);
            }
            else
            {
                // Multiple elements - combine values
                return TransferParameterFromMultipleElements(doc, targetElement, sourceElements, mapping);
            }
        }
        
        private bool TransferParameterFromElement(
            Document doc, 
            Element targetElement, 
            Element sourceElement, 
            ParameterMapping mapping)
        {
            try
            {
                var sourceParam = sourceElement.LookupParameter(mapping.SourceParameter);
                if (sourceParam == null) return false;
                
                var targetParam = targetElement.LookupParameter(mapping.TargetParameter);
                if (targetParam == null || targetParam.IsReadOnly) return false;
                
                var value = sourceParam.AsString();
                if (string.IsNullOrEmpty(value)) return false;
                
                // Apply renaming if conditions exist
                value = _renamingService.ApplyRenaming(value, mapping.SourceParameter);
                
                // Apply service type abbreviation if this is a service type parameter
                if (IsServiceTypeParameter(mapping.SourceParameter))
                {
                    value = _abbreviationService.GetAbbreviation(value, mapping.SourceParameter);
                }
                
                targetParam.Set(value);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error transferring parameter: {ex.Message}");
                return false;
            }
        }
        
        private bool TransferParameterFromMultipleElements(
            Document doc, 
            Element targetElement, 
            List<Element> sourceElements, 
            ParameterMapping mapping)
        {
            try
            {
                var values = new List<string>();
                
                foreach (var sourceElement in sourceElements)
                {
                    var sourceParam = sourceElement.LookupParameter(mapping.SourceParameter);
                    if (sourceParam != null)
                    {
                        var value = sourceParam.AsString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            // Apply renaming if conditions exist
                            value = _renamingService.ApplyRenaming(value, mapping.SourceParameter);
                            
                            // Apply service type abbreviation if this is a service type parameter
                            if (IsServiceTypeParameter(mapping.SourceParameter))
                            {
                                value = _abbreviationService.GetAbbreviation(value, mapping.SourceParameter);
                            }
                            
                            values.Add(value);
                        }
                    }
                }
                
                if (values.Count == 0) return false;
                
                var targetParam = targetElement.LookupParameter(mapping.TargetParameter);
                if (targetParam == null || targetParam.IsReadOnly) return false;
                
                var combinedValue = string.Join(mapping.Separator, values.Distinct());
                targetParam.Set(combinedValue);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error transferring multiple parameters: {ex.Message}");
                return false;
            }
        }
        
        private bool ElementsIntersect(Element element1, Element element2)
        {
            try
            {
                var geom1 = element1.get_Geometry(new Options());
                var geom2 = element2.get_Geometry(new Options());
                
                if (geom1 == null || geom2 == null) return false;
                
                // Simple bounding box intersection check
                var bbox1 = geom1.GetBoundingBox();
                var bbox2 = geom2.GetBoundingBox();
                
                // Custom intersection check since BoundingBoxXYZ.Intersects doesn't exist
                return bbox1.Min.X <= bbox2.Max.X && bbox1.Max.X >= bbox2.Min.X &&
                       bbox1.Min.Y <= bbox2.Max.Y && bbox1.Max.Y >= bbox2.Min.Y &&
                       bbox1.Min.Z <= bbox2.Max.Z && bbox1.Max.Z >= bbox2.Min.Z;
            }
            catch
            {
                return false;
            }
        }
        
        private bool IsServiceTypeParameter(string parameterName)
        {
            var serviceTypeParameters = new List<string>
            {
                "System Abbreviation",
                "System Name",
                "System Type",
                "Family Name",
                "Type Name"
            };
            
            return serviceTypeParameters.Any(param => 
                string.Equals(param, parameterName, StringComparison.OrdinalIgnoreCase));
        }
        
        #endregion
    }
}
