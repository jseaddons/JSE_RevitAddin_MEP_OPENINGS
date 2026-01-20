# Insulation Detection Database Persistence - Implementation Plan

## Executive Summary

This document outlines the plan to implement database persistence for insulation detection data. Currently:
- ✅ **OOP architecture exists** - `IInsulationDetector` and `InsulationDetector` detect insulation
- ✅ **ClashZone model has properties** - `IsInsulated` and `InsulationThickness` exist
- ✅ **Detection works during refresh** - Insulation info is populated in `ClashZone` objects
- ❌ **NOT saved to database** - Missing columns and repository persistence
- ❌ **Placement can't access it** - When loading from DB, insulation info is missing

**Critical Issue:** Placement strategies need insulation data to calculate correct sleeve dimensions, but if the data isn't in the database, placement can't compute it!

---

## 1. Current State Analysis

### 1.1 What We Have (OOP Architecture)

✅ **Services Created:**
- `IInsulationDetector` - Interface for detecting insulation status and thickness
- `InsulationDetector` - Implementation that checks `InsulationThickness` parameter
- Methods:
  - `IsInsulated(Element)` - Returns true if insulation thickness > threshold (0.001ft ≈ 0.3mm)
  - `GetInsulationThickness(Element)` - Returns thickness in Revit internal units (feet)
  - `GetInsulationInfo(Element, MepElementSize)` - Returns (isInsulated, thickness) tuple
    - Prioritizes `MepElementSize` data (from strategies) over direct element detection

✅ **ClashZone Model Properties:**
- `IsInsulated` (bool) - Line 485 in `ClashZone.cs`
  - Comment: "✅ OOP METHOD: Whether this MEP element is insulated (detected via IInsulationDetector)"
  - Comment: "Pre-calculated during refresh and saved to DB for clearance calculations during placement"
- `InsulationThickness` (double) - Line 492 in `ClashZone.cs`
  - Comment: "✅ OOP METHOD: Insulation thickness in Revit internal units (feet)"
  - Comment: "Pre-calculated during refresh and saved to DB if element is insulated"
  - Comment: "Used by strategies to calculate appropriate clearance for insulated elements"
- `InsulationType` (string) - Line 479 in `ClashZone.cs`
  - Stores "Normal" or "Insulated" (legacy property, used for clearance selection)

✅ **ClashZoneService_Legacy:**
- `CreateClashZone()` method uses `InsulationDetector.GetInsulationInfo()` (Line 2100-2101):
  ```csharp
  var insulationDetector = new InsulationDetector();
  var (isInsulated, insulationThickness) = insulationDetector.GetInsulationInfo(mepElement, mepElementSize);
  var insulationType = isInsulated ? "Insulated" : "Normal";
  ```
- Populates `ClashZone` properties during refresh (Line 2189-2191):
  ```csharp
  InsulationType = insulationType,
  IsInsulated = isInsulated,
  InsulationThickness = insulationThickness,
  ```

✅ **Strategies Use Insulation:**
- `UniversalSleevePlacerService` uses `clashZone.IsInsulated` and `clashZone.InsulationThickness`
- `IInsulationAwareSizingService` calculates final dimensions with insulation contribution
- Formula: `FinalSize = BaseSize + (2 * InsulationThickness) + (2 * Clearance)`

### 1.2 What's Missing (Database Persistence)

❌ **Database Schema:**
- **NO** `IsInsulated` column in `ClashZones` table
- **NO** `InsulationThickness` column in `ClashZones` table
- Current schema only has columns created in `EnsureSchemaCreated()` (Lines 277-331)
- Schema migrations are handled in `EnsureSchemaUpgraded()` but insulation columns were never added

❌ **Repository Layer:**
- `ClashZoneRepository.InsertClashZone()` - **NO** parameters for `IsInsulated` or `InsulationThickness`
- `ClashZoneRepository.UpdateClashZone()` - **NO** parameters for insulation columns
- `ClashZoneRepository.AddClashZoneParameters()` - **NO** insulation parameters
- INSERT statement (Lines 500-553) - **NO** insulation columns
- UPDATE statement (Lines 635-690) - **NO** insulation columns

❌ **Load Logic:**
- When loading `ClashZone` from database, `IsInsulated` and `InsulationThickness` are not populated
- Default values (false, 0.0) are used, causing incorrect sleeve sizing

### 1.3 Current Flow (Broken)

```
REFRESH:
1. ClashZoneService_Legacy.CreateClashZone()
   → Calls InsulationDetector.GetInsulationInfo()
   → Returns (isInsulated, insulationThickness) tuple
2. ClashZone object created with:
   → IsInsulated = isInsulated (true/false)
   → InsulationThickness = insulationThickness (in feet)
   → InsulationType = isInsulated ? "Insulated" : "Normal"
3. Save to Database
   → ❌ IsInsulated NOT saved (no column exists!)
   → ❌ InsulationThickness NOT saved (no column exists!)
   → Only XML persistence has the values

PLACEMENT:
1. Load ClashZone from Database
   → ❌ IsInsulated = false (default, not loaded from DB)
   → ❌ InsulationThickness = 0.0 (default, not loaded from DB)
2. UniversalSleevePlacerService checks clashZone.IsInsulated
   → ❌ Always false (not in database)
   → Falls into non-insulated clearance logic
3. Final sleeve size is WRONG
   → Missing insulation contribution = BaseSize + (2 * Clearance)
   → Should be = BaseSize + (2 * InsulationThickness) + (2 * Clearance)
```

---

## 2. Required Changes

### 2.1 Database Schema Migration

**Add to ClashZones Table:**
```sql
ALTER TABLE ClashZones ADD COLUMN IsInsulated INTEGER DEFAULT 0;
ALTER TABLE ClashZones ADD COLUMN InsulationThickness REAL DEFAULT 0.0;
```

**Migration Details:**
- Column: `IsInsulated` (INTEGER: 0 = false, 1 = true)
- Column: `InsulationThickness` (REAL: thickness in Revit internal units - feet)
- Default values: `IsInsulated = 0`, `InsulationThickness = 0.0`
- Add to `EnsureSchemaUpgraded()` method in `SleeveDbContext.cs`
- Migration should check if columns exist before adding (idempotent)

**Migration Code:**
```csharp
// In SleeveDbContext.EnsureSchemaUpgraded()
private void EnsureInsulationColumns()
{
    try
    {
        // Check if IsInsulated column exists
        using (var cmd = _connection.CreateCommand())
        {
            cmd.CommandText = @"
                SELECT COUNT(*) FROM pragma_table_info('ClashZones') 
                WHERE name = 'IsInsulated'";
            var exists = Convert.ToInt32(cmd.ExecuteScalar()) > 0;
            
            if (!exists)
            {
                ExecuteCommand("ALTER TABLE ClashZones ADD COLUMN IsInsulated INTEGER DEFAULT 0", null);
                _logger("[SQLite] ✅ Added IsInsulated column to ClashZones");
            }
        }
        
        // Check if InsulationThickness column exists
        using (var cmd = _connection.CreateCommand())
        {
            cmd.CommandText = @"
                SELECT COUNT(*) FROM pragma_table_info('ClashZones') 
                WHERE name = 'InsulationThickness'";
            var exists = Convert.ToInt32(cmd.ExecuteScalar()) > 0;
            
            if (!exists)
            {
                ExecuteCommand("ALTER TABLE ClashZones ADD COLUMN InsulationThickness REAL DEFAULT 0.0", null);
                _logger("[SQLite] ✅ Added InsulationThickness column to ClashZones");
            }
        }
    }
    catch (Exception ex)
    {
        _logger($"[SQLite] ⚠️ Error adding insulation columns: {ex.Message}");
        // Don't throw - allow schema to continue if columns already exist
    }
}
```

### 2.2 Repository Layer Updates

**Add Parameters to Insert/Update:**
```csharp
// In AddClashZoneParameters() method (ClashZoneRepository.cs):
cmd.Parameters.AddWithValue("@IsInsulated", clashZone.IsInsulated ? 1 : 0);
cmd.Parameters.AddWithValue("@InsulationThickness", clashZone.InsulationThickness);
```

**Update SQL INSERT Statement:**
```sql
INSERT INTO ClashZones (
    -- ... existing columns ...
    IsInsulated,
    InsulationThickness
) VALUES (
    -- ... existing values ...
    @IsInsulated,
    @InsulationThickness
)
```

**Update SQL UPDATE Statement:**
```sql
UPDATE ClashZones SET
    -- ... existing columns ...
    IsInsulated = @IsInsulated,
    InsulationThickness = @InsulationThickness,
    UpdatedAt = CURRENT_TIMESTAMP
WHERE ClashZoneId = @ClashZoneId
```

**Update Load Logic:**
When loading `ClashZone` from database, populate insulation properties:
```csharp
// In LoadClashZoneFromReader() or similar method:
clashZone.IsInsulated = GetInt(reader, "IsInsulated", 0) == 1;
clashZone.InsulationThickness = GetDouble(reader, "InsulationThickness", 0.0);

// Also update InsulationType for backward compatibility:
clashZone.InsulationType = clashZone.IsInsulated ? "Insulated" : "Normal";
```

### 2.3 ClashZone Model (No Changes Needed)

✅ **Already Has Properties:**
- `IsInsulated` (bool) - Already exists, just needs database persistence
- `InsulationThickness` (double) - Already exists, just needs database persistence
- `InsulationType` (string) - Legacy property, can be derived from `IsInsulated`

**Optional Enhancement:**
- Could add property to derive `InsulationType` from `IsInsulated` for consistency

### 2.4 ClashZoneService_Legacy (No Changes Needed)

✅ **Already Populates Correctly:**
- Uses `InsulationDetector.GetInsulationInfo()` (Line 2100-2101)
- Sets `IsInsulated` and `InsulationThickness` in `ClashZone` (Line 2190-2191)
- No changes needed - data is already being detected and stored in model

### 2.5 Placement Services (No Changes Needed)

✅ **Already Use Correctly:**
- `UniversalSleevePlacerService` checks `clashZone.IsInsulated` and `clashZone.InsulationThickness`
- `IInsulationAwareSizingService` uses insulation data in calculations
- No changes needed - once data is loaded from database, placement will work correctly

---

## 3. Implementation Steps

### Step 1: Database Schema Migration
1. Add `EnsureInsulationColumns()` method to `SleeveDbContext.cs`
2. Call it from `EnsureSchemaUpgraded()` method
3. Test migration on existing database
4. Verify columns are created with correct defaults

### Step 2: Update Repository Insert Logic
1. Add `@IsInsulated` and `@InsulationThickness` parameters to `AddClashZoneParameters()`
2. Update INSERT SQL statement to include columns
3. Add parameter values for INSERT operation

### Step 3: Update Repository Update Logic
1. Add `@IsInsulated` and `@InsulationThickness` parameters to UPDATE statement
2. Ensure parameters are added in `AddClashZoneParameters()` for updates
3. Verify UPDATE preserves insulation values

### Step 4: Update Repository Load Logic
1. Find method that loads `ClashZone` from database reader
2. Add code to populate `IsInsulated` and `InsulationThickness` from reader
3. Update `InsulationType` based on `IsInsulated` value

### Step 5: Testing & Validation
1. Test refresh: Verify insulation info is saved to database
2. Test placement: Verify insulation data is loaded from database
3. Test with insulated elements: Verify sleeve sizes include insulation contribution
4. Test with non-insulated elements: Verify no insulation contribution added
5. Verify backward compatibility (old databases without new columns)

---

## 4. Insulation Calculation Logic Summary

### Formula:
```
IF element is insulated (IsInsulated = true):
    FinalSize = BaseSize + (2 * InsulationThickness) + (2 * Clearance)
    Example: Pipe OD 32mm + Insulation 25mm + Clearance 25mm
    = 32 + (2 * 25) + (2 * 25) = 132mm → Rounded up to 150mm

ELSE (not insulated):
    FinalSize = BaseSize + (2 * Clearance)
    Example: Pipe OD 32mm + Clearance 25mm
    = 32 + (2 * 25) = 82mm → Rounded up to 100mm
```

### Key Points:
- ✅ Insulation thickness is in Revit internal units (feet)
- ✅ Contribution is applied on both sides: `2 * InsulationThickness`
- ✅ Formula applies to all categories: Pipes, Ducts, Cable Trays
- ✅ Rounding is applied after insulation contribution (per UI settings)

### Example Calculation:
```
Pipe OD: 32mm (0.1049ft)
Insulation: 25mm (0.0820ft)
Clearance: 25mm (0.0820ft)

If insulated:
    FinalDiameter = 0.1049 + (2 * 0.0820) + (2 * 0.0820)
                  = 0.1049 + 0.1640 + 0.1640
                  = 0.4329ft
                  = 132mm
    After rounding (50mm, always up): 150mm

If not insulated:
    FinalDiameter = 0.1049 + (2 * 0.0820)
                  = 0.2689ft
                  = 82mm
    After rounding (50mm, always up): 100mm
```

---

## 5. Files to Modify

### Database:
- `Data/SleeveDbContext.cs` - Add migration for new columns
- `Data/Repositories/ClashZoneRepository.cs` - Add parameters, SQL updates, load logic

### Models:
- `Models/ClashZone.cs` - **NO CHANGES** (properties already exist)

### Services:
- `Services/ClashZoneService_Legacy.cs` - **NO CHANGES** (already detects and populates)
- `Services/UniversalSleevePlacerService.cs` - **NO CHANGES** (already uses insulation data)
- `Services/Strategies/*` - **NO CHANGES** (already use insulation data)

### Documentation:
- `SLEEVE_PLACEMENT_METHODOLOGY_REFACTORED.md` - Update OOP architecture section

---

## 6. Backward Compatibility

### Migration Strategy:
1. New columns have DEFAULT values (0 for IsInsulated, 0.0 for InsulationThickness)
2. Existing databases will have NULL/default values until next refresh
3. Old XML files will still work (insulation info loaded from XML if present)
4. During next refresh, all elements will be re-checked for insulation

### Data Migration:
- On first refresh after migration, all elements will be re-checked
- Insulation info will be populated for elements that have insulation
- Old databases without insulation columns will get default values (not insulated)
- Placement will work but may use incorrect sizes until refresh runs

### Safe Defaults:
- If `IsInsulated` is false or missing → treat as non-insulated (safe, conservative)
- If `InsulationThickness` is 0.0 or missing → no insulation contribution (safe, conservative)
- Placement will still work, just without insulation contribution until refresh

---

## 7. Success Criteria

✅ **Database:**
- `IsInsulated` column exists and is populated correctly (0 or 1)
- `InsulationThickness` column exists and stores thickness in feet (REAL)
- Columns have appropriate DEFAULT values (0, 0.0)

✅ **Refresh:**
- All MEP elements checked for insulation during refresh
- Insulation status and thickness saved to database
- Both insulated and non-insulated elements handled correctly

✅ **Persistence:**
- INSERT operation includes insulation columns
- UPDATE operation includes insulation columns
- Load operation populates `IsInsulated` and `InsulationThickness` from database

✅ **Placement:**
- Placement loads insulation data from database correctly
- Insulated elements: FinalSize = BaseSize + (2 * InsulationThickness) + (2 * Clearance)
- Non-insulated elements: FinalSize = BaseSize + (2 * Clearance)
- Rounding applied correctly after insulation contribution

✅ **Categories:**
- Works for all MEP categories: Pipes, Ducts, Cable Trays
- Insulation detection applies to all categories uniformly

---

## 8. Testing Scenarios

### Scenario 1: Insulated Pipe
1. Create clash zone with insulated pipe (OD 32mm, insulation 25mm)
2. Refresh → Verify `IsInsulated = true`, `InsulationThickness = 0.0820ft` in database
3. Place sleeve → Verify final diameter = 32 + 50 + 50 = 132mm → 150mm (rounded)

### Scenario 2: Non-Insulated Pipe
1. Create clash zone with non-insulated pipe (OD 32mm)
2. Refresh → Verify `IsInsulated = false`, `InsulationThickness = 0.0` in database
3. Place sleeve → Verify final diameter = 32 + 50 = 82mm → 100mm (rounded)

### Scenario 3: Insulated Duct
1. Create clash zone with insulated duct (600x300mm, insulation 25mm)
2. Refresh → Verify insulation data saved to database
3. Place sleeve → Verify final width = 600 + 50 + 50 = 700mm, height = 300 + 50 + 50 = 400mm

### Scenario 4: Database Migration
1. Start with existing database (no insulation columns)
2. Run refresh → Verify migration adds columns
3. Verify existing records have default values (0, 0.0)
4. Verify new records have correct insulation values

---

## 9. Related Files Reference

### Database Schema:
- `Data/SleeveDbContext.cs` - Lines 277-331 (CREATE TABLE), EnsureSchemaUpgraded() method

### Repository:
- `Data/Repositories/ClashZoneRepository.cs`
  - Line 500-553: INSERT statement
  - Line 635-690: UPDATE statement
  - Line 702+: AddClashZoneParameters() method
  - Load methods: Find where ClashZone is loaded from database reader

### Model:
- `Models/ClashZone.cs`
  - Line 479: InsulationType property
  - Line 485: IsInsulated property
  - Line 492: InsulationThickness property

### Detection:
- `Services/ClashZoneService_Legacy.cs`
  - Line 2099-2102: Insulation detection
  - Line 2189-2191: ClashZone population

### Placement:
- `Services/UniversalSleevePlacerService.cs` - Uses `clashZone.IsInsulated` and `clashZone.InsulationThickness`
- `Services/Sizing/InsulationAwareSizingService.cs` - Calculates with insulation

---

## End of Plan

