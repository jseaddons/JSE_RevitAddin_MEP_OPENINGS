# Opening Schedule Implementation Plan

## Overview
This document outlines the implementation plan to generate opening schedules similar to the attached "BUILDER'S CUT LEVEL 23 SCHEDULE FOR ELECTRICAL WALL OPENING" format. The schedule will extract data from Revit openings and their associated MEP elements to create comprehensive construction documentation.

## Schedule Column Analysis

### Required Columns:
1. **TAG NO**: Unique identifier for each opening
2. **SIZE OF OPENING (WXH/DIA)**: Width × Height or Diameter of the opening
3. **TYPE OF SERVICE**: Type of MEP service passing through the opening
4. **CENTER OF SERVICE FROM FFL**: Vertical distance from Finished Floor Level to service center
5. **SERVICE SIZE (OUTER DIMENSION) + ANNULAR SPACE = TOTAL**: MEP element dimensions + clearance = opening size
6. **FALSE CEILING LEVEL FROM FFL**: Vertical distance from FFL to ceiling level

## Data Source Mapping

### 1. TAG NO
- **Source**: Custom parameter in opening family
- **Implementation**: Tag generation service (existing)
- **Format**: Sequential numbering (E01, E02, etc.)

### 2. SIZE OF OPENING (WXH/DIA)
- **Source**: Opening family parameters (Width, Height, Diameter)
- **Implementation**: Direct parameter extraction
- **Format**: "WIDTH×HEIGHT" or "DIAMETER"

### 3. TYPE OF SERVICE
- **Source**: MEP element types passing through opening
- **Implementation**: Service type classification service
- **Format**: Abbreviated service types (EDB, DB, ELV, NCT, TELECOM, GSM)

### 4. CENTER OF SERVICE FROM FFL
- **Source**: Opening family parameter (elevation from FFL)
- **Implementation**: Direct parameter extraction
- **Format**: "+XXXX" (millimeters from FFL)

### 5. SERVICE SIZE + ANNULAR SPACE = TOTAL
- **Source**: MEP element dimensions + clearance parameters
- **Implementation**: MEP element analysis + clearance calculation
- **Format**: "SERVICE_DIM + CLEARANCE = TOTAL_DIM"

### 6. FALSE CEILING LEVEL FROM FFL
- **Source**: Host element (wall/floor) ceiling level parameter
- **Implementation**: Host element analysis
- **Format**: "+XXXX" (millimeters from FFL)

## Implementation Architecture

### 1. Core Services

#### A. OpeningScheduleService
```csharp
public class OpeningScheduleService
{
    // Generate complete opening schedule
    public OpeningSchedule GenerateSchedule(Document doc, 
        List<ElementId> openingIds, 
        ScheduleConfiguration config);
    
    // Export schedule to various formats
    public void ExportSchedule(OpeningSchedule schedule, 
        string filePath, 
        ExportFormat format);
    
    // Validate schedule data
    public List<ScheduleValidationError> ValidateSchedule(OpeningSchedule schedule);
}
```

#### B. ServiceTypeClassificationService
```csharp
public class ServiceTypeClassificationService
{
    // Classify MEP element service type
    public string ClassifyServiceType(Element mepElement);
    
    // Get service abbreviation
    public string GetServiceAbbreviation(string serviceType);
    
    // Handle multiple services in cluster opening
    public List<string> ClassifyClusterServices(List<Element> mepElements);
    
    // Combine multiple service types
    public string CombineServiceTypes(List<string> serviceTypes);
}
```

#### C. MepElementAnalysisService
```csharp
public class MepElementAnalysisService
{
    // Get MEP element dimensions
    public MepElementDimensions GetElementDimensions(Element mepElement);
    
    // Calculate total service size including clearance
    public ServiceSizeCalculation CalculateServiceSize(
        List<Element> mepElements, 
        double clearance);
    
    // Handle cluster opening analysis
    public ClusterOpeningAnalysis AnalyzeClusterOpening(
        ElementId openingId, 
        List<Element> mepElements);
}
```

#### D. HostElementAnalysisService
```csharp
public class HostElementAnalysisService
{
    // Get ceiling level from host element
    public double GetCeilingLevelFromFFL(Element hostElement);
    
    // Get FFL reference level
    public double GetFFLReference(Element hostElement);
    
    // Calculate opening center from FFL
    public double GetOpeningCenterFromFFL(Element opening, Element hostElement);
}
```

### 2. Data Models

#### A. OpeningScheduleItem
```csharp
public class OpeningScheduleItem
{
    public string TagNumber { get; set; }
    public string OpeningSize { get; set; }
    public string ServiceType { get; set; }
    public double CenterFromFFL { get; set; }
    public string ServiceSizeCalculation { get; set; }
    public double CeilingLevelFromFFL { get; set; }
    public ElementId OpeningId { get; set; }
    public List<ElementId> MepElementIds { get; set; }
    public ElementId HostElementId { get; set; }
}
```

#### B. OpeningSchedule
```csharp
public class OpeningSchedule
{
    public string Title { get; set; }
    public string Level { get; set; }
    public string Discipline { get; set; }
    public List<OpeningScheduleItem> Items { get; set; }
    public DateTime GeneratedDate { get; set; }
    public string ProjectName { get; set; }
}
```

#### C. MepElementDimensions
```csharp
public class MepElementDimensions
{
    public double Width { get; set; }
    public double Height { get; set; }
    public double Diameter { get; set; }
    public string DimensionString { get; set; }
    public string ElementType { get; set; }
}
```

#### D. ServiceSizeCalculation
```csharp
public class ServiceSizeCalculation
{
    public List<MepElementDimensions> ServiceDimensions { get; set; }
    public double AnnularSpace { get; set; }
    public double TotalWidth { get; set; }
    public double TotalHeight { get; set; }
    public string CalculationString { get; set; }
}
```

#### E. ClusterOpeningAnalysis
```csharp
public class ClusterOpeningAnalysis
{
    public ElementId OpeningId { get; set; }
    public List<Element> MepElements { get; set; }
    public List<string> ServiceTypes { get; set; }
    public string CombinedServiceType { get; set; }
    public ServiceSizeCalculation TotalServiceSize { get; set; }
    public double OpeningCenterFromFFL { get; set; }
}
```

### 3. Service Type Classification

#### A. Service Type Mapping
```csharp
public static class ServiceTypeMappings
{
    public static readonly Dictionary<string, string> ServiceAbbreviations = new()
    {
        { "Electrical Distribution Board", "EDB" },
        { "Distribution Board", "DB" },
        { "Extra-Low Voltage", "ELV" },
        { "Network Cable Tray", "NCT" },
        { "Telecommunications", "TELECOM" },
        { "Global System for Mobile", "GSM" },
        { "Pipe", "P" },
        { "Duct", "D" },
        { "Cable Tray", "CT" },
        { "Conduit", "C" }
    };
    
    public static readonly Dictionary<string, string> CategoryMappings = new()
    {
        { "Electrical Equipment", "EDB" },
        { "Pipe Fitting", "P" },
        { "Duct Fitting", "D" },
        { "Cable Tray Fitting", "CT" },
        { "Conduit Fitting", "C" }
    };
}
```

#### B. Service Type Classification Logic
```csharp
public class ServiceTypeClassifier
{
    public string ClassifyByCategory(Element mepElement)
    {
        var category = mepElement.Category?.Name;
        return ServiceTypeMappings.CategoryMappings
            .GetValueOrDefault(category, "Unknown");
    }
    
    public string ClassifyByFamilyName(Element mepElement)
    {
        var familyName = mepElement.get_Parameter(BuiltInParameter.ELEM_FAMILY_NAME)?.AsString();
        return ServiceTypeMappings.ServiceAbbreviations
            .GetValueOrDefault(familyName, "Unknown");
    }
    
    public string ClassifyBySystemType(Element mepElement)
    {
        var systemType = mepElement.get_Parameter(BuiltInParameter.RBS_SYSTEM_TYPE_PARAM)?.AsString();
        return ServiceTypeMappings.ServiceAbbreviations
            .GetValueOrDefault(systemType, "Unknown");
    }
}
```

### 4. Cluster Opening Handling

#### A. Cluster Opening Detection
```csharp
public class ClusterOpeningDetector
{
    public bool IsClusterOpening(ElementId openingId, List<Element> mepElements)
    {
        return mepElements.Count > 1;
    }
    
    public List<Element> GetMepElementsInOpening(ElementId openingId)
    {
        // Get all MEP elements that intersect with the opening
        var opening = openingId.Document.GetElement(openingId);
        var openingGeometry = opening.get_Geometry(new Options());
        
        var intersectingElements = new List<Element>();
        // Implementation to find intersecting MEP elements
        return intersectingElements;
    }
}
```

#### B. Cluster Service Combination
```csharp
public class ClusterServiceCombiner
{
    public string CombineServiceTypes(List<string> serviceTypes)
    {
        var uniqueTypes = serviceTypes.Distinct().ToList();
        
        if (uniqueTypes.Count == 1)
            return uniqueTypes.First();
        
        return string.Join(" & ", uniqueTypes);
    }
    
    public ServiceSizeCalculation CombineServiceSizes(
        List<MepElementDimensions> dimensions, 
        double clearance)
    {
        var totalWidth = dimensions.Sum(d => d.Width) + clearance;
        var totalHeight = dimensions.Max(d => d.Height) + clearance;
        
        var calculationString = string.Join(" + ", 
            dimensions.Select(d => $"{d.Width}x{d.Height} {d.ElementType}")) + 
            $" +{clearance}mm A.SPACE ={totalWidth}x{totalHeight}";
        
        return new ServiceSizeCalculation
        {
            ServiceDimensions = dimensions,
            AnnularSpace = clearance,
            TotalWidth = totalWidth,
            TotalHeight = totalHeight,
            CalculationString = calculationString
        };
    }
}
```

## UI Implementation

### 1. Schedule Generation Dialog
```
┌─────────────────────────────────────────────────────────────┐
│ Opening Schedule Generation                                │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│ Schedule Configuration:                                     │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ Title: [BUILDER'S CUT LEVEL 23 SCHEDULE FOR ELECTRICAL] │ │
│ │ Level: [Level 23 ▼]                                    │ │
│ │ Discipline: [Electrical ▼]                             │ │
│ │ Tag Prefix: [E]                                        │ │
│ │ Starting Number: [01]                                  │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ Opening Selection:                                          │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ [✓] All Openings in Project                            │ │
│ │ [✓] Selected Openings Only                             │ │
│ │ [✓] Filter by Level                                    │ │
│ │ [✓] Filter by Discipline                               │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ Export Options:                                             │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ Format: [Excel ▼] [CSV ▼] [PDF ▼]                     │ │
│ │ File Path: [C:\Schedules\Electrical_Openings.xlsx]      │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│ │  [Generate] │  │  [Preview]   │  │  [Cancel]    │         │
│ └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────┘
```

### 2. Schedule Preview Dialog
```
┌─────────────────────────────────────────────────────────────┐
│ Schedule Preview                                            │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│ BUILDER'S CUT LEVEL 23 SCHEDULE FOR ELECTRICAL WALL OPENING │
│                                                             │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ TAG │ SIZE OF    │ TYPE OF │ CENTER OF │ SERVICE SIZE │ │
│ │ NO  │ OPENING    │ SERVICE │ SERVICE   │ + A.SPACE    │ │
│ │     │ (WXH/DIA)  │         │ FROM FFL  │ = TOTAL      │ │
│ ├─────────────────────────────────────────────────────────┤ │
│ │ E01 │ 150x150    │ EDB     │ +2615     │ 50x50 EDB    │ │
│ │     │            │         │           │ +50mm A.SPACE│ │
│ │     │            │         │           │ =150x150     │ │
│ ├─────────────────────────────────────────────────────────┤ │
│ │ E02 │ 150x150    │ DB      │ +2990     │ 50x50 DB     │ │
│ │     │            │         │           │ +50mm A.SPACE│ │
│ │     │            │         │           │ =150x150     │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│ │   [Export]  │  │  [Refresh]  │  │  [Close]     │         │
│ └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────┘
```

## Implementation Steps

### Phase 1: Core Services (Week 1-2)
1. **Create OpeningScheduleService**
   - Implement schedule generation logic
   - Add export functionality
   - Implement validation

2. **Create ServiceTypeClassificationService**
   - Implement service type classification
   - Add abbreviation mapping
   - Handle cluster service combination

3. **Create MepElementAnalysisService**
   - Implement MEP element dimension extraction
   - Add clearance calculation
   - Handle cluster opening analysis

### Phase 2: Data Models (Week 2-3)
1. **Define OpeningScheduleItem model**
2. **Define OpeningSchedule model**
3. **Define MepElementDimensions model**
4. **Define ServiceSizeCalculation model**
5. **Define ClusterOpeningAnalysis model**

### Phase 3: Service Type Classification (Week 3-4)
1. **Implement ServiceTypeMappings**
2. **Create ServiceTypeClassifier**
3. **Add cluster service combination logic**
4. **Implement service abbreviation system**

### Phase 4: UI Implementation (Week 4-5)
1. **Create ScheduleGenerationDialog**
2. **Create SchedulePreviewDialog**
3. **Implement export functionality**
4. **Add validation and error handling**

### Phase 5: Integration and Testing (Week 5-6)
1. **Integrate with existing opening system**
2. **Add to main dialog**
3. **Implement command integration**
4. **Add comprehensive testing**

## Integration Points

### 1. Main Dialog Integration
- Add "Generate Schedule" button to main dialog
- Integrate with existing opening management workflow
- Add schedule generation to opening creation/update process

### 2. Command Integration
- Create `GenerateOpeningScheduleCommand`
- Add to ribbon interface
- Implement external event handling

### 3. Export Integration
- Integrate with existing export system
- Add multiple format support (Excel, CSV, PDF)
- Implement template-based export

## Configuration Management

### 1. Schedule Templates
- Create configurable schedule templates
- Support different disciplines (Electrical, Mechanical, Plumbing)
- Allow custom column configurations

### 2. Service Type Configuration
- Configurable service type mappings
- Custom abbreviation definitions
- Discipline-specific service classifications

### 3. Export Configuration
- Configurable export formats
- Template-based export options
- Custom file naming conventions

## Error Handling

### 1. Data Validation
- Validate opening parameters
- Check MEP element accessibility
- Verify host element properties

### 2. Service Type Validation
- Handle unknown service types
- Validate service combinations
- Check abbreviation mappings

### 3. Export Validation
- Validate export file paths
- Check export format compatibility
- Handle export failures

## Performance Considerations

### 1. Batch Processing
- Process multiple openings simultaneously
- Implement progress reporting
- Add cancellation support

### 2. Memory Management
- Efficient element querying
- Minimize geometry calculations
- Implement cleanup procedures

### 3. Caching
- Cache service type classifications
- Cache MEP element dimensions
- Implement smart refresh logic

## Future Enhancements

### 1. Advanced Features
- Custom schedule templates
- Multi-level schedule generation
- Automated schedule updates

### 2. Integration Features
- Export to BIM 360
- Integration with other MEP tools
- Automated report generation

### 3. Reporting Features
- Schedule comparison tools
- Change tracking
- Audit trails

## Conclusion

This implementation plan provides a comprehensive approach to generating opening schedules that match the format and functionality of the attached electrical wall opening schedule. The modular architecture allows for incremental development and testing, while the integration points ensure seamless operation with the existing opening management system.

The implementation handles both single MEP element openings and cluster openings with multiple MEP elements, providing accurate service type classification, dimension calculations, and comprehensive schedule generation capabilities.
