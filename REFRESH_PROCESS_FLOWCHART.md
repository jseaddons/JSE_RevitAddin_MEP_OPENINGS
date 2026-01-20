# Refresh Process Flowchart

## Mermaid Flowchart Code

```mermaid
flowchart TD
    A([User clicks Refresh Button]) --> B[OnRefreshClick Event Handler]
    B --> C[Call Refresh() Method]

    C --> D{Check Filter Selection}
    D -->|No Filters Selected| E[Show Warning Dialog<br/>"Please select at least one filter"]
    D -->|Filters Selected| F[Get Selected Filters]

    E --> Z([End - User Action Required])

    F --> G[Initialize Progress Bar<br/>Status: "Analyzing clash zones..."]
    G --> H[Get Current Document]

    H --> I{Document Available?}
    I -->|No| J[Log Warning<br/>"No document available"]
    I -->|Yes| K[Collect Structural Elements<br/>from Linked Files]

    J --> Z

    K --> L[Collect MEP Elements<br/>Pipes, Ducts, Cable Trays]
    L --> M[Find Intersections<br/>MEP vs Structural]

    M --> N{Intersections Found?}
    N -->|No| O[Log Warning<br/>"No intersections found"]
    N -->|Yes| P[Create ClashZoneStorage]

    O --> Q[Update Status<br/>"No intersections found"]
    Q --> R[Complete Progress Bar]
    R --> S([End - No Action Required])

    P --> T[Convert Intersections<br/>to ClashZone Objects]
    T --> U[Find First Enabled Filter]

    U --> V{Enabled Filter Found?}
    V -->|No| W[Log Warning<br/>"No enabled filter found"]
    V -->|Yes| X[Save ClashZones to Filter]

    W --> R

    X --> Y[Update Filter<br/>LastModified = Now]
    Y --> AA[Log Success<br/>"Saved X clash zones to filter"]
    AA --> BB[Complete Progress Bar<br/>Status: "Refresh completed successfully"]
    BB --> CC([End - Success])

    style A fill:#e1f5fe
    style E fill:#ffebee
    style J fill:#ffebee
    style O fill:#fff3e0
    style W fill:#fff3e0
    style CC fill:#e8f5e8
    style S fill:#f3e5f5
    style Z fill:#ffebee
```

## Process Steps Breakdown

### **Code Architecture Changes (Latest Update)**

#### **Before (Monolithic OnRefreshClick)**
- All refresh logic was contained within the `OnRefreshClick` event handler
- Event handler was responsible for both event handling and business logic
- Difficult to test and maintain

#### **After (Separated Concerns)**
- `OnRefreshClick` event handler: Simple event handling only
- `Refresh()` method: Core business logic following flowchart process
- Clean separation of concerns for better maintainability

### 1. **User Interaction**
- User clicks the Refresh button
- `OnRefreshClick` event handler is triggered
- Handler calls the `Refresh()` method (new architecture)

### 2. **Filter Validation**
- `Refresh()` method checks if at least one filter is selected
- Prompts user with warning dialog if no filters selected
- Returns early if validation fails

### 3. **Document Access**
- Retrieves current Revit document from constructor or UIDocument
- Handles document unavailability gracefully
- Aborts process if no document available

### 4. **Element Collection**
- **Structural Elements**: From linked architectural files
- **MEP Elements**: From current document (pipes, ducts, cable trays)
- Handles coordinate transformations for linked elements

### 5. **Intersection Detection**
- Uses `MepIntersectionService.FindIntersections()`
- Processes each MEP element against structural elements
- Returns list of intersection points and bounding boxes

### 6. **Clash Zone Storage**
- Creates `ClashZoneStorage` object
- Converts intersections to `ClashZone` objects
- Saves to first enabled filter
- Updates filter timestamp

### 7. **Completion**
- Updates progress bar to 100%
- Logs success/failure messages
- Enables refresh button for next use

## Error Handling Paths

### **No Filters Selected**
- Shows warning dialog
- Returns to user for action
- No processing occurs

### **No Document Available**
- Logs warning message
- Aborts processing
- User must check Revit state

### **No Intersections Found**
- Logs informational message
- Completes process normally
- No clash zones to save

### **No Enabled Filter**
- Logs warning message
- Completes process without saving
- User must enable a filter

## Success Criteria

### **Complete Success**
- Filters selected ✓
- Document available ✓
- Intersections found ✓
- Enabled filter available ✓
- Clash zones saved ✓

### **Partial Success**
- Process completes but with warnings
- Some data may be missing
- User informed of limitations

### **Failure**
- Critical error prevents completion
- User prompted for corrective action
- Process aborts gracefully

## Implementation Changes Log

### **Latest Update: Code Architecture Refactoring (2025-09-20)**

#### **Changes Made:**
1. **Extracted Core Logic**: Moved all refresh business logic from `OnRefreshClick` event handler into a new `Refresh()` method
2. **Separated Concerns**: Event handler now only handles UI events, business logic is in dedicated method
3. **Improved Maintainability**: Refresh logic is now easier to test, debug, and modify independently
4. **Flowchart Alignment**: Code structure now directly matches the documented process flow

#### **Code Changes:**
- **Before**: `OnRefreshClick` contained ~80 lines of mixed event handling and business logic
- **After**: `OnRefreshClick` contains 4 lines calling `Refresh()`, `Refresh()` contains all business logic

#### **Benefits:**
- ✅ **Testability**: Business logic can be unit tested independently
- ✅ **Maintainability**: Changes to refresh logic don't affect event handling
- ✅ **Readability**: Clear separation between UI and business logic
- ✅ **Reusability**: Refresh logic can be called from other places if needed

#### **Files Modified:**
- `Views/EmergencyMainDialog.cs`: Refactored refresh implementation
- `REFRESH_PROCESS_FLOWCHART.md`: Updated documentation to reflect new architecture

---

*This flowchart represents the current implementation of the Refresh functionality. For the most up-to-date process flow, refer to the source code in `EmergencyMainDialog.cs`.*
