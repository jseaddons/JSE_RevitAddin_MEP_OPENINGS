# Parameter Service Implementation Plan

## Overview
This document outlines the implementation plan for a Parameter Service UI that transfers parameter values from Host elements, Reference elements (MEP), and Levels to Openings, based on the conVoid parameter transfer functionality described in [conVoid Parameter Transfer Guide](https://www.support.conclass.tech/convoid-revit-parameter-transfer-4).

## Core Functionality

### 1. Parameter Transfer Categories
The Parameter Service will support three main transfer categories:

#### A. Reference Element to Openings
- Transfer parameter values from MEP elements (Provision for Voids) to openings
- Examples: System Abbreviation, MEP system information, element properties
- Source: MEP elements (Pipes, Ducts, Cable Trays, Conduits)

#### B. Host to Opening
- Transfer Host properties (walls, floors, ceilings) to openings
- Examples: Fire rating, material information, structural properties
- Support for multiple hosts with separator functionality
- Source: Host elements that openings penetrate

#### C. Level to Openings
- Transfer level parameter values to openings
- Examples: Level name, elevation information, building information
- Source: Revit levels

#### D. Model Name Transfer
- Built-in parameter to transfer linked model names to openings
- Useful for coordination across multiple models

## UI Design Specifications

### 1. Main Parameter Transfer Dialog
```
┌─────────────────────────────────────────────────────────────┐
│ Parameter Transfer Settings                                │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│ Reference Element to Openings                              │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ Source Parameter: [System Abbreviation ▼]              │ │
│ │ Target Parameter: [MEP_System ▼]                       │ │
│ │ [✓] Transfer from Reference Elements                    │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ Host to Opening                                             │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ Source Parameter: [Fire Rating ▼]                     │ │
│ │ Target Parameter: [Opening_Fire_Rating ▼]             │ │
│ │ Separator: [; ] (for multiple hosts)                    │ │
│ │ [✓] Transfer from Host Elements                        │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ Level to Openings                                           │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ Source Parameter: [Level Name ▼]                       │ │
│ │ Target Parameter: [Opening_Level ▼]                     │ │
│ │ [✓] Transfer from Level                                 │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ Model Information                                           │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ [✓] Transfer Model Name to Opening                     │ │
│ │ Target Parameter: [Model_Name ▼]                       │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│ │    [OK]     │  │  [Cancel]   │  │   [Help]    │         │
│ └─────────────┘  └─────────────┘  └─────────────┘         │
└─────────────────────────────────────────────────────────────┘
```

### 2. Parameter Value Renaming Dialog
```
┌─────────────────────────────────────────────────────────────┐
│ Parameter Value Renaming                                   │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│ Renaming Conditions:                                        │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ Original Value    │ New Value    │ Actions              │ │
│ │ Sprinklers        │ SPR          │ [Edit] [Delete]      │ │
│ │ Hydronic          │ S            │ [Edit] [Delete]      │ │
│ │ Sanitary          │ S            │ [Edit] [Delete]      │ │
│ │ Supply Air        │ V            │ [Edit] [Delete]      │ │
│ │ Exhaust Air       │ V            │ [Edit] [Delete]      │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ ┌─────────────┐  ┌─────────────┐  ┌─────────────┐         │
│ │   [Add]     │  │ [Import CSV]│  │ [Export CSV]│         │
│ └─────────────┘  └─────────────┘  └─────────────┘         │
│                                                             │
│ ┌─────────────┐  ┌─────────────┐                          │
│ │    [OK]     │  │  [Cancel]   │                          │
│ └─────────────┘  └─────────────┘                          │
└─────────────────────────────────────────────────────────────┘
```

## Implementation Architecture

### 1. Core Services

#### A. ParameterTransferService
```csharp
public class ParameterTransferService
{
    // Transfer parameters from MEP elements to openings
    public void TransferFromReferenceElements(Document doc, 
        List<ElementId> openingIds, 
        ParameterMapping mapping);
    
    // Transfer parameters from host elements to openings
    public void TransferFromHostElements(Document doc, 
        List<ElementId> openingIds, 
        ParameterMapping mapping);
    
    // Transfer parameters from levels to openings
    public void TransferFromLevels(Document doc, 
        List<ElementId> openingIds, 
        ParameterMapping mapping);
    
    // Transfer model names to openings
    public void TransferModelNames(Document doc, 
        List<ElementId> openingIds, 
        string targetParameter);
}
```

#### B. ParameterMappingService
```csharp
public class ParameterMappingService
{
    // Get available parameters from element category
    public List<ParameterInfo> GetAvailableParameters(Document doc, 
        BuiltInCategory category);
    
    // Create parameter mapping configuration
    public ParameterMapping CreateMapping(string sourceParam, 
        string targetParam, 
        string separator = ";");
    
    // Validate parameter compatibility
    public bool ValidateMapping(ParameterInfo source, ParameterInfo target);
}
```

#### C. ParameterRenamingService
```csharp
public class ParameterRenamingService
{
    // Apply renaming conditions to parameter values
    public string ApplyRenaming(string originalValue, 
        List<RenamingCondition> conditions);
    
    // Import renaming conditions from CSV
    public List<RenamingCondition> ImportFromCsv(string csvPath);
    
    // Export renaming conditions to CSV
    public void ExportToCsv(List<RenamingCondition> conditions, string csvPath);
}
```

### 2. Data Models

#### A. ParameterMapping
```csharp
public class ParameterMapping
{
    public string SourceParameter { get; set; }
    public string TargetParameter { get; set; }
    public string Separator { get; set; } = ";";
    public bool IsEnabled { get; set; }
    public TransferType TransferType { get; set; }
}

public enum TransferType
{
    ReferenceToOpening,
    HostToOpening,
    LevelToOpening,
    ModelNameToOpening
}
```

#### B. RenamingCondition
```csharp
public class RenamingCondition
{
    public string OriginalValue { get; set; }
    public string NewValue { get; set; }
    public string ParameterName { get; set; }
}
```

#### C. ParameterInfo
```csharp
public class ParameterInfo
{
    public string Name { get; set; }
    public string Type { get; set; }
    public bool IsReadOnly { get; set; }
    public BuiltInCategory Category { get; set; }
}
```

### 3. UI Components

#### A. ParameterTransferDialog (WinForms)
- Main dialog for parameter transfer configuration
- Tabbed interface for different transfer types
- Parameter selection dropdowns
- Preview of transfer results

#### B. ParameterRenamingDialog (WinForms)
- Grid-based interface for renaming conditions
- CSV import/export functionality
- Add/Edit/Delete operations for conditions

#### C. ParameterPreviewPanel (WinForms)
- Real-time preview of parameter transfers
- Validation feedback
- Error reporting

## Implementation Steps

### Phase 1: Core Services (Week 1-2)
1. **Create ParameterTransferService**
   - Implement basic parameter reading/writing
   - Add element filtering logic
   - Handle parameter type validation

2. **Create ParameterMappingService**
   - Implement parameter discovery
   - Add mapping validation
   - Create parameter compatibility checks

3. **Create ParameterRenamingService**
   - Implement renaming logic
   - Add CSV import/export
   - Create condition management

### Phase 2: Data Models (Week 2-3)
1. **Define ParameterMapping model**
2. **Define RenamingCondition model**
3. **Define ParameterInfo model**
4. **Create configuration persistence**

### Phase 3: UI Implementation (Week 3-4)
1. **Create ParameterTransferDialog**
   - Design main interface
   - Implement parameter selection
   - Add transfer type tabs

2. **Create ParameterRenamingDialog**
   - Design renaming interface
   - Implement CSV functionality
   - Add condition management

3. **Create ParameterPreviewPanel**
   - Design preview interface
   - Implement real-time updates
   - Add validation feedback

### Phase 4: Integration (Week 4-5)
1. **Integrate with existing opening system**
2. **Add to main dialog**
3. **Implement command integration**
4. **Add error handling and logging**

### Phase 5: Testing and Refinement (Week 5-6)
1. **Unit testing for services**
2. **UI testing and validation**
3. **Performance optimization**
4. **User feedback integration**

## Integration Points

### 1. Main Dialog Integration
- Add "Parameter Transfer" button to main dialog
- Integrate with existing opening management workflow
- Add parameter transfer to opening creation/update process

### 2. Command Integration
- Create `ParameterTransferCommand`
- Add to ribbon interface
- Implement external event handling

### 3. Configuration Management
- Integrate with existing profile system
- Add parameter transfer settings to user profiles
- Implement settings persistence

## CSV File Format

### Renaming Conditions CSV Structure
```csv
Original Value,New Value,Parameter Name
Sprinklers,SPR,System Abbreviation
Hydronic,S,System Abbreviation
Sanitary,S,System Abbreviation
Supply Air,V,System Abbreviation
Exhaust Air,V,System Abbreviation
```

## Error Handling

### 1. Parameter Validation
- Check parameter existence
- Validate parameter types
- Handle read-only parameters

### 2. Element Validation
- Verify element accessibility
- Handle deleted elements
- Check element permissions

### 3. Transfer Validation
- Validate transfer results
- Handle transfer failures
- Provide rollback capability

## Performance Considerations

### 1. Batch Processing
- Process multiple openings simultaneously
- Implement progress reporting
- Add cancellation support

### 2. Memory Management
- Efficient parameter reading
- Minimize element queries
- Implement cleanup procedures

### 3. Transaction Management
- Use Revit transactions efficiently
- Implement rollback on failure
- Minimize transaction scope

## Future Enhancements

### 1. Advanced Features
- Conditional parameter transfer
- Formula-based parameter values
- Custom parameter creation

### 2. Integration Features
- Export parameter mappings
- Import from other projects
- Template management

### 3. Reporting Features
- Transfer reports
- Parameter audit trails
- Change tracking

## Testing Strategy

### 1. Unit Tests
- Service method testing
- Parameter validation testing
- Renaming logic testing

### 2. Integration Tests
- End-to-end transfer testing
- UI interaction testing
- Error scenario testing

### 3. User Acceptance Testing
- Real project testing
- Performance validation
- Usability testing

## Documentation Requirements

### 1. User Documentation
- Parameter transfer guide
- CSV format documentation
- Troubleshooting guide

### 2. Developer Documentation
- API documentation
- Service architecture guide
- Extension points documentation

### 3. Configuration Documentation
- Parameter mapping examples
- Best practices guide
- Performance tuning guide

## Conclusion

This implementation plan provides a comprehensive approach to creating a Parameter Service UI that matches the functionality described in the conVoid documentation. The modular architecture allows for incremental development and testing, while the integration points ensure seamless operation with the existing opening management system.

The implementation follows Revit API best practices and provides a user-friendly interface for parameter transfer operations, making it easy for users to enrich their openings with valuable parameter information from various sources.
