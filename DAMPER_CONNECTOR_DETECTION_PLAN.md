# Damper Connector Detection - Implementation Plan

## Executive Summary

This document outlines the plan to implement a connector-based damper detection system that:
1. **Removes dependency on damper type names** (MSFD, MSD, etc.)
2. **Uses only connector detection** - checks if damper has MEP connector and which side
3. **Saves connector information to database** for use during placement
4. **Applies clearance based on connector presence and side** from UI settings

---

## 1. Current State Analysis

### 1.1 What We Have (OOP Architecture)

✅ **Services Created:**
- `IDamperConnectorDetector` - Interface for detecting connector presence and side
- `DamperConnectorDetector` - Implementation using Revit API connectors
- `IDamperTypeDetector` - Interface for detecting damper type (still used for classification, but NOT for clearance)
- `DamperTypeDetector` - Implementation that classifies damper types
- `DamperConnectorService` - Orchestrates connector detection

✅ **Data Model:**
- `DamperConnectorInfo` - Value object with:
  - `HasMepConnector` (bool)
  - `ConnectorSide` (string: "Left", "Right", "Top", "Bottom", or empty)
  - `IsStandardDamper` (bool) - for classification only
  - `DamperType` (string) - for classification only
  - `RequiresMepSideClearance` (bool) - for classification only

✅ **ClashZone Model Properties:**
- `IsMSFDDamper` (bool) - **⚠️ OBSOLETE** - Currently set when connector found
- `DamperConnectorSide` (string) - ✅ EXISTS - Stores connector side
- `IsStandardDamper` (bool) - ✅ EXISTS - Stores damper type classification

✅ **ClashZoneService_Legacy:**
- `GetDamperConnectorInfo()` method calls `DamperConnectorService.DetectConnectorInfo()`
- Populates `ClashZone` properties during refresh:
  ```csharp
  DamperConnectorSide = connectorSide;
  IsMSFDDamper = hasMepConnector && !string.IsNullOrEmpty(connectorSide);
  IsStandardDamper = damperConnectorInfo.IsStandardDamper;
  ```

### 1.2 What's Missing (Database Persistence)

❌ **Database Schema:**
- **NO** `HasMepConnector` column in `ClashZones` table
- **NO** `DamperConnectorSide` column in `ClashZones` table
- Current schema only has columns added via migrations (not visible in CREATE TABLE statement)

❌ **Repository Layer:**
- `ClashZoneRepository.InsertClashZone()` - **NO** parameters for connector info
- `ClashZoneRepository.UpdateClashZone()` - **NO** parameters for connector info
- `ClashZoneRepository.AddClashZoneParameters()` - **NO** connector info parameters

❌ **Placement Strategy:**
- `DamperPlacementStrategy.GetDamperPlacementAdjustment()` - Uses obsolete `IsMSFDDamper` flag
- Logic checks `IsMSFDDamper && !string.IsNullOrEmpty(DamperConnectorSide)` - but if DB doesn't have values, placement can't compute!

### 1.3 Current Flow (Broken)

```
REFRESH:
1. ClashZoneService_Legacy.GetDamperConnectorInfo()
   → Calls DamperConnectorService.DetectConnectorInfo()
   → Returns DamperConnectorInfo with HasMepConnector, ConnectorSide, etc.
2. ClashZoneService_Legacy.CreateClashZone()
   → Sets ClashZone.IsMSFDDamper = hasMepConnector && !string.IsNullOrEmpty(connectorSide)
   → Sets ClashZone.DamperConnectorSide = connectorSide
3. Save to Database
   → ❌ Connector info NOT saved (no columns exist!)
   → Only XML persistence has the values

PLACEMENT:
1. Load ClashZone from Database
   → ❌ HasMepConnector info missing (not in DB)
   → ❌ DamperConnectorSide may be empty (not persisted)
2. DamperPlacementStrategy checks IsMSFDDamper flag
   → ❌ Flag may be false (not loaded from DB)
   → Falls into wrong clearance logic
```

---

## 2. Required Changes

### 2.1 Database Schema Migration

**Add to ClashZones Table:**
```sql
ALTER TABLE ClashZones ADD COLUMN HasMepConnector INTEGER DEFAULT 0;
ALTER TABLE ClashZones ADD COLUMN DamperConnectorSide TEXT DEFAULT '';
```

**Migration Details:**
- Column: `HasMepConnector` (INTEGER: 0 = false, 1 = true)
- Column: `DamperConnectorSide` (TEXT: "Left", "Right", "Top", "Bottom", or empty string)
- Default values: `HasMepConnector = 0`, `DamperConnectorSide = ''`
- Add to `EnsureSchemaUpgraded()` method in `SleeveDbContext.cs`

**Index (Optional but Recommended):**
```sql
CREATE INDEX IF NOT EXISTS idx_clashzones_has_mep_connector 
ON ClashZones(HasMepConnector) 
WHERE HasMepConnector = 1;
```

### 2.2 ClashZone Model Updates

**Remove Obsolete Property:**
- ❌ **REMOVE:** `IsMSFDDamper` (bool) - Line 517 in `ClashZone.cs`
- This flag is confusing and not descriptive

**Add/Keep Properties:**
- ✅ **KEEP:** `DamperConnectorSide` (string) - Already exists at line 510
- ✅ **ADD:** `HasMepConnector` (bool) - New property

**Updated Properties:**
```csharp
/// <summary>
/// ✅ OOP METHOD: Whether this damper has an MEP connector detected
/// Set to true when connector is found (regardless of damper type)
/// Strategy will use this flag along with DamperConnectorSide to determine clearance from UI settings
/// </summary>
public bool HasMepConnector { get; set; } = false;

/// <summary>
/// ✅ OOP METHOD: Connector side direction ("Left", "Right", "Top", "Bottom")
/// Set when MEP connector is found and side is detected
/// Used by strategy to determine which side gets MEP clearance (from UI settings)
/// </summary>
public string DamperConnectorSide { get; set; } = string.Empty;

/// <summary>
/// ✅ DEPRECATED: Use HasMepConnector instead
/// Kept for backward compatibility during migration
/// </summary>
[Obsolete("Use HasMepConnector instead")]
public bool IsMSFDDamper { get; set; } = false;
```

### 2.3 ClashZoneService_Legacy Updates

**Update Populate Logic:**
```csharp
// ✅ OOP METHOD: Use DamperConnectorService to detect connector info
var damperConnectorInfo = GetDamperConnectorInfo(mepElement, mepCategory);
string connectorSide = damperConnectorInfo.ConnectorSide;
bool hasMepConnector = damperConnectorInfo.HasMepConnector; // ✅ NEW: Direct flag

// Create ClashZone with connector info
var clashZone = new ClashZone
{
    // ... other properties ...
    HasMepConnector = hasMepConnector, // ✅ NEW: Direct flag
    DamperConnectorSide = connectorSide,
    IsMSFDDamper = hasMepConnector && !string.IsNullOrEmpty(connectorSide), // ⚠️ DEPRECATED: Keep for backward compat
    IsStandardDamper = damperConnectorInfo.IsStandardDamper, // For classification only
    // ... other properties ...
};
```

**Important Notes:**
- Do NOT check damper type name anymore
- Only check: Does damper have connector? → `HasMepConnector = true`
- If connector exists, detect side → `DamperConnectorSide = "Left"/"Right"/"Top"/"Bottom"`

### 2.4 Repository Layer Updates

**Add Parameters to Insert/Update:**
```csharp
// In AddClashZoneParameters() method:
cmd.Parameters.AddWithValue("@HasMepConnector", clashZone.HasMepConnector ? 1 : 0);
cmd.Parameters.AddWithValue("@DamperConnectorSide", clashZone.DamperConnectorSide ?? string.Empty);
```

**Update SQL INSERT Statement:**
```sql
INSERT INTO ClashZones (
    -- ... existing columns ...
    HasMepConnector,
    DamperConnectorSide
) VALUES (
    -- ... existing values ...
    @HasMepConnector,
    @DamperConnectorSide
)
```

**Update SQL UPDATE Statement:**
```sql
UPDATE ClashZones SET
    -- ... existing columns ...
    HasMepConnector = @HasMepConnector,
    DamperConnectorSide = @DamperConnectorSide,
    UpdatedAt = CURRENT_TIMESTAMP
WHERE ClashZoneId = @ClashZoneId
```

**Update Load Logic:**
When loading ClashZone from database, populate:
```csharp
clashZone.HasMepConnector = reader.GetInt32("HasMepConnector") == 1;
clashZone.DamperConnectorSide = reader.GetString("DamperConnectorSide") ?? string.Empty;
```

### 2.5 Placement Strategy Updates

**Update DamperPlacementStrategy Logic:**
```csharp
// ✅ OOP METHOD: Check if connector was detected (should apply to ALL dampers with connectors)
// For ANY damper with connector: Use MEP+Other on width (100+50=150mm)
// For damper without connector: Use Other on all sides (50mm)

if (clashZone.HasMepConnector && !string.IsNullOrEmpty(clashZone.DamperConnectorSide))
{
    // Damper HAS connector - apply asymmetric clearance
    // Width: Base + MEP side (100mm) + Other side (50mm) = Base + 150mm
    // Height: Base + Other (50mm) + Other (50mm) = Base + 100mm
    
    string connectorSide = clashZone.DamperConnectorSide;
    double mepSideClearance = mepClearance; // From UI: 100mm
    double otherSideClearance = otherClearance; // From UI: 50mm
    
    // Determine which dimension gets the MEP clearance
    double finalWidth, finalHeight;
    
    if (connectorSide == "Left" || connectorSide == "Right")
    {
        // Connector on left or right → MEP clearance on width
        finalWidth = damperWidth + insulationContribution + mepSideClearance + otherSideClearance; // 150mm total
        finalHeight = damperHeight + insulationContribution + (2 * otherSideClearance); // 100mm total
    }
    else if (connectorSide == "Top" || connectorSide == "Bottom")
    {
        // Connector on top or bottom → MEP clearance on height
        finalWidth = damperWidth + insulationContribution + (2 * otherSideClearance); // 100mm total
        finalHeight = damperHeight + insulationContribution + mepSideClearance + otherSideClearance; // 150mm total
    }
    else
    {
        // Unknown side → use symmetric clearance
        finalWidth = damperWidth + insulationContribution + (2 * otherSideClearance);
        finalHeight = damperHeight + insulationContribution + (2 * otherSideClearance);
    }
    
    // Calculate offset if needed (optional - for asymmetric positioning)
    XYZ offsetVector = XYZ.Zero;
    // ... offset calculation if needed ...
    
    return (offsetVector, finalWidth, finalHeight);
}
else
{
    // Damper has NO connector - use symmetric clearance
    // Width: Base + Other (50mm) + Other (50mm) = Base + 100mm
    // Height: Base + Other (50mm) + Other (50mm) = Base + 100mm
    (double finalW, double finalH, _) = _sizingService.CalculateFinalDimensionsFromClashZone(
        damperWidth, damperHeight, 0, clashZone, otherClearance);
    
    return (XYZ.Zero, finalW, finalH);
}
```

**Remove Logic:**
- ❌ Remove check for `IsMSFDDamper` flag
- ❌ Remove check for `damperType == "MSFD"` or `isNonStandard`
- ✅ Use ONLY: `HasMepConnector` and `DamperConnectorSide`

---

## 3. Implementation Steps

### Step 1: Database Schema Migration
1. Add migration in `SleeveDbContext.EnsureSchemaUpgraded()`
2. Add `HasMepConnector` column (INTEGER DEFAULT 0)
3. Add `DamperConnectorSide` column (TEXT DEFAULT '')
4. Test migration on existing database

### Step 2: Update ClashZone Model
1. Add `HasMepConnector` property
2. Mark `IsMSFDDamper` as `[Obsolete]` (keep for backward compat)
3. Update XML serialization if needed

### Step 3: Update ClashZoneService_Legacy
1. Update `CreateClashZone()` to set `HasMepConnector` directly from `DamperConnectorInfo`
2. Remove dependency on `IsMSFDDamper` flag logic
3. Ensure connector detection works for ALL dampers (not just MSFD)

### Step 4: Update Repository Layer
1. Add `@HasMepConnector` and `@DamperConnectorSide` parameters to `AddClashZoneParameters()`
2. Update INSERT SQL statement
3. Update UPDATE SQL statement
4. Update Load/Read logic to populate properties from database

### Step 5: Update Placement Strategy
1. Replace `IsMSFDDamper` checks with `HasMepConnector` checks
2. Remove damper type-based logic
3. Implement connector-side-based clearance calculation
4. Test with dampers that have connectors on different sides

### Step 6: Testing & Validation
1. Test refresh: Verify connector info is saved to database
2. Test placement: Verify clearance is calculated correctly
3. Test with standard dampers with connectors
4. Test with dampers without connectors
5. Verify backward compatibility (old databases without new columns)

---

## 4. Clearance Logic Summary

### Logic Flow:
```
IF damper has MEP connector (HasMepConnector = true) AND side detected:
    → Get MEP clearance from UI (default: 100mm)
    → Get Other clearance from UI (default: 50mm)
    → Apply based on connector side:
        - Left/Right → Width: Base + MEP (100mm) + Other (50mm) = Base + 150mm
        - Top/Bottom → Height: Base + MEP (100mm) + Other (50mm) = Base + 150mm
        - Other side: Base + Other (50mm) + Other (50mm) = Base + 100mm

ELSE (no connector):
    → Use symmetric clearance: Other (50mm) on all 4 sides
    → Width: Base + 100mm, Height: Base + 100mm
```

### Key Points:
- ❌ **NO** checking of damper type name (MSFD, MSD, etc.)
- ✅ **ONLY** checking if connector exists and which side
- ✅ Clearance values come from UI settings (MEP clearance, Other clearance)
- ✅ Strategy determines which dimension gets MEP clearance based on connector side

---

## 5. Files to Modify

### Database:
- `Data/SleeveDbContext.cs` - Add migration for new columns
- `Data/Repositories/ClashZoneRepository.cs` - Add parameters and SQL updates

### Models:
- `Models/ClashZone.cs` - Add `HasMepConnector`, deprecate `IsMSFDDamper`

### Services:
- `Services/ClashZoneService_Legacy.cs` - Update populate logic
- `Services/Strategies/DamperPlacementStrategy.cs` - Update clearance logic

### Documentation:
- `SLEEVE_PLACEMENT_METHODOLOGY_REFACTORED.md` - Update OOP architecture section

---

## 6. Backward Compatibility

### Migration Strategy:
1. New columns have DEFAULT values (0 for HasMepConnector, '' for DamperConnectorSide)
2. Existing databases will have NULL/default values until next refresh
3. Old XML files will still work (connector info loaded from XML if present)
4. `IsMSFDDamper` kept as `[Obsolete]` for transition period

### Data Migration:
- On first refresh after migration, all dampers will be re-checked
- Connector info will be populated for dampers that have connectors
- Old `IsMSFDDamper` values in XML will be ignored

---

## 7. Success Criteria

✅ **Database:**
- `HasMepConnector` column exists and is populated correctly
- `DamperConnectorSide` column exists and stores "Left"/"Right"/"Top"/"Bottom"

✅ **Refresh:**
- All dampers checked for connector presence
- Connector side detected and saved to database
- No dependency on damper type name

✅ **Placement:**
- Strategy checks `HasMepConnector` flag (not `IsMSFDDamper`)
- Clearance calculated based on connector side
- Width = Base + MEP (100mm) + Other (50mm) = Base + 150mm when connector present
- Height = Base + Other (50mm) + Other (50mm) = Base + 100mm when connector present

---

## End of Plan

