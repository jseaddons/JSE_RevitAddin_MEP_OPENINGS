using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents a parameter mapping configuration for transfer operations
    /// </summary>
    public class ParameterMapping
    {
        public string SourceParameter { get; set; } = string.Empty;
        public string TargetParameter { get; set; } = string.Empty;
        public string Separator { get; set; } = ";";
        public bool IsEnabled { get; set; }
        public TransferType TransferType { get; set; }
        public string Description { get; set; } = string.Empty;
        
        public ParameterMapping()
        {
        }
        
        public ParameterMapping(string sourceParameter, string targetParameter, TransferType transferType)
        {
            SourceParameter = sourceParameter;
            TargetParameter = targetParameter;
            TransferType = transferType;
            IsEnabled = true;
        }
        
        public ParameterMapping(string sourceParameter, string targetParameter, TransferType transferType, string separator)
        {
            SourceParameter = sourceParameter;
            TargetParameter = targetParameter;
            TransferType = transferType;
            Separator = separator;
            IsEnabled = true;
        }
    }
    
    /// <summary>
    /// Defines the type of parameter transfer operation
    /// </summary>
    public enum TransferType
    {
        ReferenceToOpening,
        HostToOpening,
        LevelToOpening,
        ModelNameToOpening
    }
    
    /// <summary>
    /// Represents a renaming condition for parameter values
    /// </summary>
    public class RenamingCondition
    {
        public string OriginalValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
        public string ParameterName { get; set; } = string.Empty;
        public bool IsEnabled { get; set; } = true;
        
        public RenamingCondition()
        {
        }
        
        public RenamingCondition(string originalValue, string newValue, string parameterName)
        {
            OriginalValue = originalValue;
            NewValue = newValue;
            ParameterName = parameterName;
            IsEnabled = true;
        }
    }
    
    /// <summary>
    /// Contains information about a Revit parameter
    /// </summary>
    public class ParameterInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public bool IsReadOnly { get; set; }
        public BuiltInCategory Category { get; set; }
        public string Description { get; set; } = string.Empty;
        public bool IsShared { get; set; }
        public string Group { get; set; } = string.Empty;
        
        public ParameterInfo()
        {
        }
        
        public ParameterInfo(string name, string type, BuiltInCategory category)
        {
            Name = name;
            Type = type;
            Category = category;
        }
    }
    
    /// <summary>
    /// Configuration for parameter transfer operations
    /// </summary>
    public class ParameterTransferConfiguration
    {
        public List<ParameterMapping> Mappings { get; set; } = new List<ParameterMapping>();
        public List<RenamingCondition> RenamingConditions { get; set; } = new List<RenamingCondition>();
        public bool TransferModelNames { get; set; }
        public string ModelNameParameter { get; set; } = "Model_Name";
        public bool ValidateBeforeTransfer { get; set; } = true;
        public bool CreateMissingParameters { get; set; } = false;
        public string SourceCategoryName { get; set; } = string.Empty; // e.g., "Ducts", "Duct Accessories"
        
        // Service size calculation settings
        public bool TransferServiceSizeCalculations { get; set; } = false;
        public string ServiceSizeCalculationParameter { get; set; } = "Service_Size_Calculation";
        public double DefaultClearance { get; set; } = 50.0;
        public string ClearanceSuffix { get; set; } = "mm A.SPACE";
        
        public ParameterTransferConfiguration()
        {
        }
    }
    
    /// <summary>
    /// Result of a parameter transfer operation
    /// </summary>
    public class ParameterTransferResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public int TransferredCount { get; set; }
        public int FailedCount { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        public DateTime TransferDate { get; set; }
        
        public ParameterTransferResult()
        {
            TransferDate = DateTime.Now;
        }
        
        public ParameterTransferResult(bool success, string message)
        {
            Success = success;
            Message = message;
            TransferDate = DateTime.Now;
        }
    }
    
    /// <summary>
    /// Validation error for parameter transfer operations
    /// </summary>
    public class ParameterTransferValidationError
    {
        public string ParameterName { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
        public ValidationErrorType ErrorType { get; set; }
        public ElementId ElementId { get; set; } = ElementId.InvalidElementId;
        
        public ParameterTransferValidationError()
        {
        }
        
        public ParameterTransferValidationError(string parameterName, string errorMessage, ValidationErrorType errorType)
        {
            ParameterName = parameterName;
            ErrorMessage = errorMessage;
            ErrorType = errorType;
        }
    }
    
    /// <summary>
    /// Types of validation errors
    /// </summary>
    public enum ValidationErrorType
    {
        ParameterNotFound,
        ParameterReadOnly,
        TypeMismatch,
        ElementNotFound,
        PermissionDenied,
        InvalidValue
    }
    
    /// <summary>
    /// Contains MEP element dimension information
    /// </summary>
    public class MepElementDimensions
    {
        public double Width { get; set; }
        public double Height { get; set; }
        public double Diameter { get; set; }
        public string DimensionString { get; set; } = string.Empty;
        public string ElementType { get; set; } = string.Empty;
        
        public MepElementDimensions()
        {
        }
        
        public MepElementDimensions(double width, double height, string elementType)
        {
            Width = width;
            Height = height;
            ElementType = elementType;
        }
    }
    
    /// <summary>
    /// Contains service size calculation with clearance
    /// </summary>
    public class ServiceSizeCalculation
    {
        public List<MepElementDimensions> ServiceDimensions { get; set; } = new List<MepElementDimensions>();
        public double AnnularSpace { get; set; }
        public double TotalWidth { get; set; }
        public double TotalHeight { get; set; }
        public string CalculationString { get; set; } = string.Empty;
        
        public ServiceSizeCalculation()
        {
        }
        
        public ServiceSizeCalculation(List<MepElementDimensions> dimensions, double clearance)
        {
            ServiceDimensions = dimensions;
            AnnularSpace = clearance;
        }
    }
    
    /// <summary>
    /// Contains cluster opening analysis with multiple MEP elements
    /// </summary>
    public class ClusterOpeningAnalysis
    {
        public ElementId OpeningId { get; set; } = ElementId.InvalidElementId;
        public List<Element> MepElements { get; set; } = new List<Element>();
        public List<string> ServiceTypes { get; set; } = new List<string>();
        public string CombinedServiceType { get; set; } = string.Empty;
        public ServiceSizeCalculation TotalServiceSize { get; set; } = new ServiceSizeCalculation();
        public double OpeningCenterFromFFL { get; set; }
        
        public ClusterOpeningAnalysis()
        {
        }
        
        public ClusterOpeningAnalysis(ElementId openingId, List<Element> mepElements)
        {
            OpeningId = openingId;
            MepElements = mepElements;
        }
    }
}
