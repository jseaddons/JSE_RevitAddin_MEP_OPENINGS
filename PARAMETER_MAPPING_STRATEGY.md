# Parameter Mapping Strategy After Cluster Sleeve Placement

## Overview
After cluster sleeves are placed, we need to map parameters from MEP elements to the cluster sleeves. This involves:
1. **Multiple MEP elements** inside each cluster sleeve
2. **Parameter transfer** from MEP elements to cluster sleeves
3. **User input** for prefix configuration
4. **Separate service classes** for host and reference element mapping

## Current Workflow Issues

### Problem 1: MarkParameterAddValue Command UI Dependency
- `MarkParameterAddValue` command requires user input for prefix
- This breaks the automated workflow
- User has to manually input prefix during orchestration

### Problem 2: Parameter Mapping Complexity
- Each cluster sleeve contains multiple MEP elements
- Need to decide which MEP element's parameters to use
- Host vs Reference element parameters need different handling

## Proposed Solution Architecture

### 1. UI Integration for Prefix Input
**Location:** Main UI Parameter Service Section (bottom)
**Implementation:**
```
┌─────────────────────────────────────────┐
│ Parameter Mapping Configuration          │
├─────────────────────────────────────────┤
│ Prefix for Cluster Sleeves: [____]      │
│ □ Enable Parameter Mapping               │
│ □ Map Host Element Parameters           │
│ □ Map Reference Element Parameters      │
└─────────────────────────────────────────┘
```

### 2. Service Class Architecture

#### A. ParameterMappingOrchestrator
**Purpose:** Main orchestrator for parameter mapping workflow
**Responsibilities:**
- Coordinate parameter mapping after cluster placement
- Call appropriate service classes based on category
- Handle prefix configuration from UI

#### B. ReferenceElementParameterMapper
**Purpose:** Handle parameter mapping for reference elements (MEP)
**Responsibilities:**
- Extract parameters from MEP elements within cluster sleeves
- Apply parameters to cluster sleeves
- Handle parameter conflicts (multiple MEP elements)

#### C. HostElementParameterMapper
**Purpose:** Handle parameter mapping for host elements
**Responsibilities:**
- Extract parameters from host elements (walls, structural framing, floors)
- Apply parameters to cluster sleeves
- Handle host-specific parameter logic

### 3. Parameter Mapping Strategy

#### For Reference Elements (MEP):
```
Cluster Sleeve → Find MEP Elements Inside → Extract Parameters → Apply to Sleeve
```

**Parameter Selection Logic:**
1. **Primary MEP Element:** Largest element by size/area
2. **Parameter Priority:** System Type > Service Type > Level > Dimensions
3. **Conflict Resolution:** Use most common value or first occurrence

#### For Host Elements:
```
Cluster Sleeve → Find Host Elements (Wall/Structural/Floor) → Extract Parameters → Apply to Sleeve
```

**Parameter Selection Logic:**
1. **Host Type Priority:** Wall > Structural Framing > Floor
2. **Parameter Priority:** Level > Material > Dimensions
3. **Conflict Resolution:** Use host element with highest priority

### 4. Implementation Workflow

#### Step 1: UI Integration
```csharp
// In EmergencyMainDialog.cs - Parameter Service Section
private void CreateParameterMappingSection()
{
    // Add prefix input field
    // Add checkboxes for mapping options
    // Store configuration in OpeningSettings
}
```

#### Step 2: Service Classes
```csharp
// ParameterMappingOrchestrator.cs
public class ParameterMappingOrchestrator
{
    public void MapParametersAfterClustering(List<ClusterSleeve> clusterSleeves, 
                                           OpeningSettings settings)
    {
        // Get prefix from settings
        // Call appropriate mappers based on configuration
        // Handle errors and logging
    }
}

// ReferenceElementParameterMapper.cs
public class ReferenceElementParameterMapper
{
    public void MapMepParametersToSleeve(ClusterSleeve sleeve, string prefix)
    {
        // Find MEP elements inside sleeve
        // Extract and prioritize parameters
        // Apply to sleeve with prefix
    }
}

// HostElementParameterMapper.cs
public class HostElementParameterMapper
{
    public void MapHostParametersToSleeve(ClusterSleeve sleeve, string prefix)
    {
        // Find host elements
        // Extract and prioritize parameters
        // Apply to sleeve with prefix
    }
}
```

#### Step 3: Orchestrator Integration
```csharp
// In OpeningCommandOrchestrator.cs
private void ExecuteParameterMapping(List<ClusterSleeve> clusterSleeves)
{
    if (settings.EnableParameterMapping)
    {
        var parameterOrchestrator = new ParameterMappingOrchestrator();
        parameterOrchestrator.MapParametersAfterClustering(clusterSleeves, settings);
    }
}
```

### 5. Parameter Mapping Rules

#### Reference Element Parameters (MEP):
| Parameter | Source | Priority | Example |
|-----------|--------|----------|---------|
| System Type | Primary MEP | High | "Supply Air" |
| Service Type | Primary MEP | High | "HVAC" |
| Level | All MEP | Medium | "Level 1" |
| Width/Height | Largest MEP | Medium | "600mm" |
| Diameter | Largest Pipe | Medium | "150mm" |

#### Host Element Parameters:
| Parameter | Source | Priority | Example |
|-----------|--------|----------|---------|
| Level | Host Element | High | "Level 1" |
| Material | Host Element | Medium | "Concrete" |
| Thickness | Host Element | Low | "200mm" |
| Fire Rating | Host Element | Medium | "2 Hour" |

### 6. Prefix Configuration

#### Default Prefixes:
- **Reference Elements:** "MEP_"
- **Host Elements:** "HOST_"
- **Combined:** "CLUSTER_"

#### User Customization:
- Allow user to set custom prefix
- Validate prefix format (alphanumeric + underscore)
- Store in OpeningSettings for persistence

### 7. Error Handling

#### Common Scenarios:
1. **No MEP Elements Found:** Log warning, skip parameter mapping
2. **Parameter Conflicts:** Use priority rules, log conflicts
3. **Invalid Parameters:** Skip invalid parameters, continue with valid ones
4. **Prefix Validation:** Show error message, use default prefix

### 8. Integration Points

#### A. Main UI Integration:
- Add parameter mapping section to Parameter Service
- Store configuration in OpeningSettings
- Validate user input

#### B. Orchestrator Integration:
- Call parameter mapping after cluster placement
- Pass configuration from UI
- Handle errors gracefully

#### C. Service Integration:
- Use existing ParameterTransferService for actual parameter setting
- Integrate with ParameterExtractionService for parameter discovery
- Use ParameterRenamingService for prefix handling

### 9. Testing Strategy

#### Unit Tests:
- Test parameter extraction from MEP elements
- Test parameter priority logic
- Test prefix application
- Test error handling

#### Integration Tests:
- Test full workflow from cluster placement to parameter mapping
- Test with different MEP categories
- Test with multiple elements in cluster

### 10. Future Enhancements

#### Advanced Features:
- **Parameter Templates:** Predefined parameter sets for different scenarios
- **Conditional Mapping:** Map parameters based on MEP element properties
- **Batch Operations:** Apply same parameters to multiple cluster sleeves
- **Parameter Validation:** Check parameter values before applying

## Implementation Priority

### Phase 1: Core Infrastructure
1. Create ParameterMappingOrchestrator
2. Add UI section for prefix configuration
3. Integrate with OpeningCommandOrchestrator

### Phase 2: Reference Element Mapping
1. Create ReferenceElementParameterMapper
2. Implement MEP parameter extraction logic
3. Test with different MEP categories

### Phase 3: Host Element Mapping
1. Create HostElementParameterMapper
2. Implement host parameter extraction logic
3. Test with different host types

### Phase 4: Advanced Features
1. Parameter conflict resolution
2. Advanced prefix options
3. Parameter validation and error handling

## Conclusion

This strategy provides a comprehensive approach to parameter mapping after cluster sleeve placement. By integrating the prefix configuration into the main UI and creating dedicated service classes, we can automate the parameter mapping process while maintaining flexibility and user control.

The key is to separate concerns:
- **UI:** Handle user configuration
- **Orchestrator:** Coordinate the workflow
- **Mappers:** Handle specific parameter logic
- **Services:** Provide reusable functionality

This approach ensures maintainability, testability, and extensibility for future enhancements.

