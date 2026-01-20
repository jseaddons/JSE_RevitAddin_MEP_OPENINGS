# Combine Sleeve Feature - Implementation Plan

**Document Version:** 1.0  
**Status:** Planning Phase  
**Related Document:** `SLEEVE_PLACEMENT_METHODOLOGY_REFACTORED.md`

---

## Table of Contents

1. [Overview](#1-overview)
2. [Feature Requirements](#2-feature-requirements)
3. [Architecture Design](#3-architecture-design)
4. [UI Design](#4-ui-design)
5. [Implementation Phases](#5-implementation-phases)
6. [Technical Details](#6-technical-details)
7. [Integration Points](#7-integration-points)

---

## 1. Overview

### 1.1 Purpose

The **Combine Sleeve** feature allows users to combine:
- **Cluster sleeves** with **individual sleeves** from **different categories**
- Example: Combine a cluster sleeve (Ducts) with an individual sleeve (Pipes) on the same wall

### 1.2 Key Differences from Clustering

| Feature | Clustering | Combine Sleeve |
|---------|-----------|----------------|
| **Scope** | Within same category | Cross-category |
| **Input** | Multiple individual sleeves | Cluster sleeve + Individual sleeve(s) |
| **Output** | New cluster sleeve | Combined cluster sleeve |
| **Wall Grouping** | Same wall group | Same wall group (strict) |
| **Rotation** | Supports rotated (45°, 225°) | Straight axis only (0°, 90°, 180°, 270°) |

### 1.3 Use Cases

1. **Auto Mode**: User selects a category, system automatically finds and combines compatible sleeves
2. **Manual Mode**: User manually selects 2 sleeves to combine

---

## 2. Feature Requirements

### 2.1 Wall Group Restrictions (CRITICAL)

**Strict Wall Group Isolation:**
- ✅ **Wall X** sleeves can ONLY combine with **Wall X** sleeves
- ✅ **Wall Y** sleeves can ONLY combine with **Wall Y** sleeves  
- ✅ **Floor** sleeves can ONLY combine with **Floor** sleeves
- ❌ **Wall X** CANNOT combine with **Wall Y** or **Floor**
- ❌ **Wall Y** CANNOT combine with **Wall X** or **Floor**
- ❌ **Floor** CANNOT combine with **Wall X** or **Wall Y**

### 2.2 Mode Requirements

#### 2.2.1 Auto Mode

**User Workflow:**
1. User selects target category (e.g., "Pipes")
2. System finds all cluster sleeves and individual sleeves from selected category
3. System groups by wall type (Wall X, Wall Y, Floor)
4. System combines sleeves within same wall group only
5. System creates combined cluster sleeves

**Validation:**
- If user selects incompatible wall groups → Show error: "Please select only Wall X sleeves" or "Please select only Wall Y sleeves"
- Auto mode handles this automatically by filtering

#### 2.2.2 Manual Mode

**User Workflow:**
1. User selects first sleeve (cluster or individual)
2. User selects second sleeve (cluster or individual)
3. System validates:
   - Same wall group (Wall X, Wall Y, or Floor)
   - Straight axis-aligned (0°, 90°, 180°, 270°)
4. If valid → Combine
5. If invalid → Show error message

**Error Messages:**
- "Cannot combine: Sleeves are on different wall groups (Wall X vs Wall Y)"
- "Cannot combine: Sleeves are not straight axis-aligned"
- "Cannot combine: One sleeve is on Wall X, other is on Floor"

### 2.3 Rotation Requirements

**Phase 1 (Current):**
- ✅ **Straight axis-aligned only** (0°, 90°, 180°, 270°)
- ❌ **No rotated sleeves** (45°, 225°, etc.) - Future phase

**Phase 2 (Future):**
- Support rotated axis-aligned sleeves (45°, 225°, etc.)

---

## 3. Architecture Design

### 3.1 Service Structure

```
Services/
├── Combining/
│   ├── ICombineSleeveService.cs          # Main interface
│   ├── CombineSleeveService.cs           # Main orchestrator
│   ├── Auto/
│   │   └── AutoCombineService.cs         # Auto mode logic
│   ├── Manual/
│   │   └── ManualCombineService.cs       # Manual mode logic
│   ├── Validation/
│   │   └── CombineValidationService.cs   # Wall group & rotation validation
│   ├── Algorithm/
│   │   └── CombineAlgorithmService.cs    # Bounding box calculation (reuse clustering)
│   └── Placement/
│       └── CombinePlacementService.cs    # Combined sleeve placement
```

### 3.2 Command Structure

```
Commands/
└── CombineSleeveCommand.cs               # ICommand implementation
```

### 3.3 UI Structure

```
Views/
└── CombineSleeveDialog.xaml              # WPF dialog with radio buttons
```

### 3.4 Reuse Existing Clustering Code

**Key Insight:** Combine sleeve uses **same algorithm** as clustering, but with:
- Different input (cluster + individual vs multiple individuals)
- Cross-category support
- Wall group validation

**Reusable Components:**
- ✅ `ClusterAlgorithmService` - BFS proximity algorithm
- ✅ `ClusterPlacementService` - Sleeve placement logic
- ✅ `ClusterBoundingBoxServices` - Bounding box calculation
- ✅ `ClusterDataService` - Data loading (already supports all categories)

**New Components:**
- `CombineValidationService` - Wall group validation
- `CombineSleeveService` - Orchestration

---

## 4. UI Design

### 4.1 Dialog Layout

```
┌─────────────────────────────────────────┐
│  Combine Sleeve                         │
├─────────────────────────────────────────┤
│                                         │
│  Mode Selection:                        │
│  ○ Auto                                 │
│  ● Manual                               │
│                                         │
│  ┌─────────────────────────────────┐   │
│  │ Auto Mode Options:               │   │
│  │                                   │   │
│  │  Select Category:                │   │
│  │  [Dropdown: Ducts ▼]             │   │
│  │                                   │   │
│  │  [Combine] [Cancel]               │   │
│  └─────────────────────────────────┘   │
│                                         │
│  ┌─────────────────────────────────┐   │
│  │ Manual Mode Options:             │   │
│  │                                   │   │
│  │  Select First Sleeve:            │   │
│  │  [Pick Element]                  │   │
│  │                                   │   │
│  │  Select Second Sleeve:           │   │
│  │  [Pick Element]                  │   │
│  │                                   │   │
│  │  [Combine] [Cancel]               │   │
│  └─────────────────────────────────┘   │
│                                         │
└─────────────────────────────────────────┘
```

### 4.2 Radio Button Behavior

- **Auto Mode Selected:**
  - Show category dropdown
  - Hide manual selection controls
  - Enable "Combine" button

- **Manual Mode Selected:**
  - Show "Pick Element" buttons
  - Hide category dropdown
  - Enable "Combine" button only when both sleeves selected

### 4.3 Error Messages

**Display Location:** TaskDialog or status bar

**Messages:**
- "Cannot combine: Sleeves are on different wall groups"
- "Cannot combine: Sleeves are not straight axis-aligned"
- "Please select only Wall X sleeves" (when user selects mixed groups)
- "Please select only Wall Y sleeves" (when user selects mixed groups)

---

## 5. Implementation Phases

### Phase 1: Core Infrastructure

**Tasks:**
1. ✅ Create `ICombineSleeveService` interface
2. ✅ Create `CombineSleeveService` orchestrator
3. ✅ Create `CombineValidationService` (wall group validation)
4. ✅ Create `CombineSleeveCommand` (ICommand)
5. ✅ Create `CombineSleeveDialog.xaml` (UI)

**Deliverables:**
- Service interfaces and base classes
- Command structure
- UI dialog (non-functional)

### Phase 2: Manual Mode

**Tasks:**
1. ✅ Implement `ManualCombineService`
2. ✅ Implement element picker (Revit selection)
3. ✅ Implement wall group validation
4. ✅ Implement rotation validation (straight axis only)
5. ✅ Implement bounding box calculation (reuse clustering)
6. ✅ Implement combined sleeve placement
7. ✅ Test manual mode end-to-end

**Deliverables:**
- Working manual mode
- Validation logic
- Error handling

### Phase 3: Auto Mode

**Tasks:**
1. ✅ Implement `AutoCombineService`
2. ✅ Implement category filtering
3. ✅ Implement wall group grouping
4. ✅ Implement automatic combination logic
5. ✅ Implement batch processing
6. ✅ Test auto mode end-to-end

**Deliverables:**
- Working auto mode
- Batch processing
- Progress reporting

### Phase 4: Integration & Testing

**Tasks:**
1. ✅ Integrate with existing clustering code
2. ✅ Test cross-category combinations
3. ✅ Test wall group restrictions
4. ✅ Test error handling
5. ✅ Performance testing
6. ✅ User acceptance testing

**Deliverables:**
- Fully integrated feature
- Test results
- Documentation

---

## 6. Technical Details

### 6.1 Wall Group Detection

**Method:** Extract from `ClashZone.StructuralElementType` and `ClashZone.StructuralElementNormal`

**Logic:**
```csharp
public enum WallGroup
{
    WallX,      // Normal = (1,0,0) or (-1,0,0)
    WallY,      // Normal = (0,1,0) or (0,-1,0)
    Floor       // Normal = (0,0,1) or (0,0,-1)
}

public WallGroup GetWallGroup(ClashZone clashZone)
{
    var normal = clashZone.StructuralElementNormal;
    if (normal == null) return WallGroup.Floor; // Default
    
    // Check X-axis walls
    if (Math.Abs(normal.X) > 0.9 && Math.Abs(normal.Y) < 0.1)
        return WallGroup.WallX;
    
    // Check Y-axis walls
    if (Math.Abs(normal.Y) > 0.9 && Math.Abs(normal.X) < 0.1)
        return WallGroup.WallY;
    
    // Check floors
    if (Math.Abs(normal.Z) > 0.9)
        return WallGroup.Floor;
    
    return WallGroup.Floor; // Default
}
```

### 6.2 Rotation Validation

**Method:** Check if sleeve is straight axis-aligned

**Logic:**
```csharp
public bool IsStraightAxisAligned(ClashZone clashZone)
{
    double angleDeg = clashZone.MepElementRotationAngle * 180 / Math.PI;
    
    // Normalize to 0-360
    while (angleDeg < 0) angleDeg += 360;
    while (angleDeg >= 360) angleDeg -= 360;
    
    // Check if close to 0°, 90°, 180°, or 270°
    double threshold = 2.0; // 2 degree tolerance
    return Math.Abs(angleDeg) < threshold ||
           Math.Abs(angleDeg - 90) < threshold ||
           Math.Abs(angleDeg - 180) < threshold ||
           Math.Abs(angleDeg - 270) < threshold;
}
```

### 6.3 Bounding Box Calculation

**Reuse:** `ClusterBoundingBoxServices.GetClusterBoundingBox()`

**Input:**
- List of sleeves (cluster sleeve + individual sleeves)
- All must be same wall group
- All must be straight axis-aligned

**Output:**
- Combined bounding box (width, height, depth, midpoint)

### 6.4 Combined Sleeve Placement

**Reuse:** `ClusterPlacementService.PlaceClusterSleeve()`

**Modifications:**
- Use combined bounding box
- Set `MEP_Category` parameter to "Combined" or comma-separated categories
- Set `IsCombined` parameter to `true`
- Preserve original sleeve IDs in `ClashZoneIdsJson`

---

## 7. Integration Points

### 7.1 Existing Services

**Reuse:**
- ✅ `ClusterDataService` - Load sleeves from database
- ✅ `ClusterAlgorithmService` - Proximity algorithm (if needed)
- ✅ `ClusterPlacementService` - Placement logic
- ✅ `ClusterBoundingBoxServices` - Bounding box calculation
- ✅ `ClashZoneRepository` - Database operations

**New:**
- `CombineSleeveService` - Orchestration
- `CombineValidationService` - Validation
- `CombineSleeveCommand` - Command

### 7.2 Database Schema

**No Changes Required:**
- `ClusterSleeves` table already supports cross-category
- `SleeveSnapshots` table already supports cluster sleeves
- `ClashZones` table already has all required data

**Optional Enhancement:**
- Add `IsCombined` flag to `ClusterSleeves` table
- Add `CombinedCategories` column (comma-separated)

### 7.3 UI Integration

**Location:** Main ribbon or toolbar

**Button:**
- Icon: 🔗 (link/combine icon)
- Tooltip: "Combine Sleeve"
- Command: `CombineSleeveCommand`

---

## 8. Algorithm Details

### 8.1 Manual Mode Algorithm

```
1. User selects first sleeve (S1)
2. User selects second sleeve (S2)
3. Validate:
   a. Get wall group for S1 (WG1)
   b. Get wall group for S2 (WG2)
   c. If WG1 != WG2 → Error: "Different wall groups"
   d. Check if S1 is straight axis-aligned
   e. Check if S2 is straight axis-aligned
   f. If not straight → Error: "Not straight axis-aligned"
4. If valid:
   a. Load sleeve data (bounding boxes, placement points)
   b. Calculate combined bounding box
   c. Place combined sleeve
   d. Delete original sleeves (or mark as combined)
   e. Save to database
```

### 8.2 Auto Mode Algorithm

```
1. User selects category (C)
2. Load all sleeves:
   a. Cluster sleeves (all categories)
   b. Individual sleeves (category C)
3. Group by wall group:
   a. Wall X group
   b. Wall Y group
   c. Floor group
4. For each wall group:
   a. Find cluster sleeves in this group
   b. Find individual sleeves (category C) in this group
   c. For each cluster sleeve:
      - Find nearby individual sleeves (within tolerance)
      - If found → Combine
   d. For remaining individual sleeves:
      - Find nearby individual sleeves (within tolerance)
      - If found → Combine into new cluster
5. Save all combined sleeves to database
```

### 8.3 Proximity Check

**Reuse:** Same tolerance as clustering (`JoinOpeningsDistance` setting)

**Logic:**
```csharp
double tolerance = settingsModel.JoinOpeningsDistance; // mm
double toleranceInternal = UnitUtils.ConvertToInternalUnits(tolerance, UnitTypeId.Millimeters);

// Check if sleeves are within tolerance (edge-to-edge or center-to-center)
bool areNearby = CalculateDistance(sleeve1, sleeve2) <= toleranceInternal;
```

---

## 9. Error Handling

### 9.1 Validation Errors

**Wall Group Mismatch:**
- Error: "Cannot combine: Sleeves are on different wall groups"
- Action: Show which wall groups (e.g., "Wall X vs Wall Y")
- User Action: Select sleeves from same wall group

**Rotation Mismatch:**
- Error: "Cannot combine: Sleeves are not straight axis-aligned"
- Action: Show rotation angles
- User Action: Select only straight axis-aligned sleeves

**Category Mismatch (Auto Mode):**
- Error: "No sleeves found for selected category"
- Action: Show available categories
- User Action: Select different category

### 9.2 Placement Errors

**Transaction Rollback:**
- If placement fails → Rollback transaction
- Show error message
- Preserve original sleeves

**Database Errors:**
- Log error to `combine_sleeve_errors.log`
- Show user-friendly message
- Preserve original sleeves

---

## 10. Testing Strategy

### 10.1 Unit Tests

- Wall group detection
- Rotation validation
- Bounding box calculation
- Validation logic

### 10.2 Integration Tests

- Manual mode end-to-end
- Auto mode end-to-end
- Cross-category combinations
- Wall group restrictions

### 10.3 User Acceptance Tests

- Real-world scenarios
- Performance with large datasets
- Error handling
- UI usability

---

## 11. Future Enhancements

### Phase 2: Rotated Sleeves

- Support 45°, 225° rotations
- Rotated bounding box calculation
- Rotated placement logic

### Phase 3: Multi-Category Combinations

- Combine 3+ categories
- Smart category naming
- Category priority rules

### Phase 4: Visual Feedback

- Highlight compatible sleeves
- Preview combined bounding box
- 3D visualization

---

## 12. Success Criteria

✅ **Functional:**
- Manual mode works for 2 sleeves
- Auto mode works for selected category
- Wall group restrictions enforced
- Straight axis validation works
- Combined sleeves placed correctly

✅ **Performance:**
- Auto mode processes 100+ sleeves in < 30 seconds
- Manual mode responds in < 2 seconds

✅ **User Experience:**
- Clear error messages
- Intuitive UI
- Helpful tooltips

---

## 13. Dependencies

**Required:**
- Existing clustering services (already implemented)
- Database schema (no changes needed)
- UI framework (WPF)

**Optional:**
- Progress bar for auto mode
- Undo/redo support
- Batch processing UI

---

**End of Document**

