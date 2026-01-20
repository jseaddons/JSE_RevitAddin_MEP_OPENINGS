# CONDITIONS.xml Architecture

## Overview

This document describes the **CONDITIONS.xml** architecture pattern implemented for managing opening placement settings.

## Problem Statement

Previously, clearance values and opening type preferences were:
- Stored in **static properties** in `EmergencyMainDialog`
- Required the UI to remain in memory for commands to access values
- Mixed concerns: UI thread and Revit API thread
- No persistence between sessions

## Solution: CONDITIONS.xml Pattern

### Architecture Separation

```
┌─────────────────────────────────────────────────────────────┐
│                      USER CLICKS OK                          │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│  EmergencyMainDialog.SaveConditionsToXml()                  │
│  • Reads UI values (clearances, opening types)              │
│  • Creates OpeningConditions object                          │
│  • Saves to FilterName_CONDITIONS.xml                        │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│             CONDITIONS.xml Written to Disk                   │
│  Location: %APPDATA%\JSE_MEP_Openings\Projects\             │
│            Default\Filters\                                  │
│  Example: Ventilation_ducts_CONDITIONS.xml                  │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│           External Event Raised (Non-blocking)               │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│  DuctSleevePlacementCommand.Execute()                       │
│  • LoadConditionsFromXml()                                   │
│  • Reads FilterName_CONDITIONS.xml                          │
│  • Creates DuctSleevePlacerService with conditions           │
└─────────────────────────────────────────────────────────────┘
                           │
                           ▼
┌─────────────────────────────────────────────────────────────┐
│  DuctSleevePlacerService                                    │
│  • Uses _conditions.ClearanceSettings                        │
│  • Selects clearance based on:                              │
│    - ClashZone.DuctShape (Round/Rectangular)                │
│    - ClashZone.InsulationType (Normal/Insulated)            │
│  • No UI access required                                     │
└─────────────────────────────────────────────────────────────┘
```

## File Structure

### OpeningConditions Model (`Models/OpeningConditions.cs`)

```csharp
public class OpeningConditions
{
    public string FilterName { get; set; }
    public string Category { get; set; }
    public DateTime CreatedDate { get; set; }
    public DateTime LastModified { get; set; }
    public ClearanceSettings ClearanceSettings { get; set; }
    public OpeningTypePreferences OpeningTypePreferences { get; set; }
    public LevelConstraints LevelConstraints { get; set; }
    public string CreationMode { get; set; }
}

public class ClearanceSettings
{
    public double RectangularNormal { get; set; } = 50.0;      // mm
    public double RectangularInsulated { get; set; } = 25.0;   // mm
    public double RoundNormal { get; set; } = 50.0;            // mm
    public double RoundInsulated { get; set; } = 50.0;         // mm
}

public class OpeningTypePreferences
{
    public string RoundDucts { get; set; } = "Circular";
    public string Pipes { get; set; } = "Circular";
}
```

### ConditionsService (`Services/ConditionsService.cs`)

Manages saving/loading of CONDITIONS.xml files:

```csharp
public class ConditionsService
{
    // Save conditions to XML
    public bool SaveConditions(OpeningConditions conditions)
    
    // Load conditions from XML
    public OpeningConditions LoadConditions(string filterName)
    
    // Get file path for conditions
    public string GetConditionsFilePath(string filterName)
    
    // Check if conditions exist
    public bool ConditionsExist(string filterName)
}
```

## Example XML Structure

**File:** `Ventilation_ducts_CONDITIONS.xml`

```xml
<?xml version="1.0" encoding="utf-8"?>
<OpeningConditions>
  <FilterName>Ventilation_ducts</FilterName>
  <Category>Ducts</Category>
  <CreatedDate>2025-10-07T12:00:00</CreatedDate>
  <LastModified>2025-10-07T14:30:00</LastModified>
  <ClearanceSettings>
    <RectangularNormal>75.0</RectangularNormal>
    <RectangularInsulated>25.0</RectangularInsulated>
    <RoundNormal>50.0</RoundNormal>
    <RoundInsulated>50.0</RoundInsulated>
  </ClearanceSettings>
  <OpeningTypePreferences>
    <RoundDucts>Circular</RoundDucts>
    <Pipes>Circular</Pipes>
  </OpeningTypePreferences>
  <LevelConstraints>
    <HorizontalLevel>Host Level</HorizontalLevel>
    <VerticalLevel>Host Level</VerticalLevel>
  </LevelConstraints>
  <CreationMode>Opening</CreationMode>
</OpeningConditions>
```

## Data Flow

### 1. Refresh Phase (Clash Detection)

```
RefreshService
  ↓
Detects: MEP element shape (Round/Rectangular) from family name
Detects: Insulation type (Normal/Insulated) from parameters
  ↓
Stores in ClashZone:
  - DuctShape: "Round" or "Rectangular"
  - InsulationType: "Normal" or "Insulated"
  ↓
Saves to: Ventilation_ducts.xml
```

### 2. Configuration Phase (UI → XML)

```
User changes clearances/opening types in UI
  ↓
Clicks OK
  ↓
EmergencyMainDialog.SaveConditionsToXml()
  ↓
For each selected filter + category combination:
  - Creates combinedKey = "FilterName_Category"
  - Reads UI values (clearances, opening types)
  - Creates OpeningConditions object
  ↓
ConditionsService.SaveConditions(conditions, combinedKey)
  ↓
Saves to: FilterName_Category_CONDITIONS.xml
  Example: Ventilation_Pipes_CONDITIONS.xml
```

### 3. Placement Phase (XML → Revit)

```
UniversalSleevePlacementCommand.Execute()
  ↓
LoadConditionsFromXml()
  ↓
Creates combinedKey = "FilterName_Category"
  ↓
ConditionsService.LoadConditions(combinedKey)
  ↓
Loads: FilterName_Category_CONDITIONS.xml
  ↓
Creates UniversalSleevePlacerService(_doc, _conditions)
  ↓
For each ClashZone:
  - Read: MEP category (from ClashZone)
  - Read: Clearance from _conditions.ClearanceSettings
  - Read: Opening type from _conditions.OpeningTypePreferences
  ↓
GetClearanceFromConditions(category, mepSize)
  ↓
SelectUniversalFamily(clashZone, mepSize)
  ↓
Uses opening type preferences:
  - Pipes: _conditions.OpeningTypePreferences.Pipes
  - Round Ducts: _conditions.OpeningTypePreferences.RoundDucts
```

## Benefits

### 1. **Clean Separation of Concerns**
- **REFRESH**: Detects element properties (shape, insulation) → ClashZone.xml
- **CONDITIONS**: Stores user preferences (clearances, types) → CONDITIONS.xml
- **PLACEMENT**: Combines both to place sleeves

### 2. **No UI Dependencies**
- Commands don't access UI directly
- Works even if UI is closed
- Thread-safe

### 3. **Persistence**
- Settings saved to disk
- Survive Revit restarts
- Can be version-controlled

### 4. **Testability**
- Can load conditions from test XML files
- Mock conditions for unit tests
- Independent of UI state

### 5. **Scalability**
- Easy to add new settings (e.g., LevelConstraints)
- Per-filter customization
- Future: Global vs per-project settings

## File Naming Convention

```
FilterName_Category_CONDITIONS.xml
```

Examples:
- `Ventilation_Ducts_CONDITIONS.xml`
- `Ventilation_Pipes_CONDITIONS.xml`
- `Ventilation_DuctAccessories_CONDITIONS.xml`
- `Plumbing_Pipes_CONDITIONS.xml`
- `Electrical_CableTrays_CONDITIONS.xml`

**🛡️ ARCHITECTURE UPDATE:** Uses BOTH filter name AND category to allow different clearance/opening types per category within the same filter.

Stored alongside clash zone files:
```
%APPDATA%\JSE_MEP_Openings\Projects\Default\Filters\
├── Ventilation_Ducts.xml                  (CLASH ZONES - data)
├── Ventilation_Ducts_CONDITIONS.xml       (CONDITIONS - settings)
├── Ventilation_Pipes.xml
├── Ventilation_Pipes_CONDITIONS.xml
├── Ventilation_DuctAccessories.xml
└── Ventilation_DuctAccessories_CONDITIONS.xml
```

## Critical Methods

### ⚠️ DO NOT REMOVE ⚠️

1. **EmergencyMainDialog.SaveConditionsToXml()**
   - Saves UI values to CONDITIONS.xml using combined key
   - Called on OK click before external event
   - Creates separate XML files per category

2. **UniversalSleevePlacementCommand.LoadConditionsFromXml()**
   - Loads CONDITIONS.xml using combined key (FilterName_Category)
   - Called in constructor
   - Gets category-specific settings

3. **UniversalSleevePlacerService.GetClearanceFromConditions()**
   - Selects clearance from _conditions based on category + insulation
   - Called during sleeve placement
   - Uses CONDITIONS service for all clearance types

4. **UniversalSleevePlacerService.SelectUniversalFamily()**
   - Uses opening type preferences from CONDITIONS
   - Pipes: _conditions.OpeningTypePreferences.Pipes
   - Round Ducts: _conditions.OpeningTypePreferences.RoundDucts

5. **ClashZoneService.DetectNewClashZones()**
   - Stores raw dimensions (no pre-calculated clearance)
   - All clearance handled by CONDITIONS service

## Migration Notes

### What Changed

**Before:**
```csharp
// Static properties in UI
public static double CurrentDuctNormalClearance { get; set; }

// Commands accessed UI directly
var clearance = EmergencyMainDialog.CurrentDuctNormalClearance;
```

**After:**
```csharp
// Saved to XML
var conditions = new OpeningConditions();
conditionsService.SaveConditions(conditions);

// Commands load from XML
_conditions = conditionsService.LoadConditions("Ventilation_ducts");
var clearance = _conditions.ClearanceSettings.RectangularNormal;
```

### What Was Removed

- ❌ Static clearance properties in `EmergencyMainDialog`
- ❌ `UpdateClearanceValuesFromUI()` method
- ❌ Direct UI access from commands

### What Was Added

- ✅ `OpeningConditions` model
- ✅ `ConditionsService` for XML management
- ✅ `SaveConditionsToXml()` on OK click
- ✅ `LoadConditionsFromXml()` in commands
- ✅ XML persistence layer

## Future Enhancements

1. **Global Settings**
   - `GLOBAL_CONDITIONS.xml` for defaults
   - Per-filter overrides

2. **Version Management**
   - Conditions versioning
   - Migration scripts

3. **UI Enhancements**
   - Load/Save presets
   - Import/Export settings

4. **Advanced Conditions**
   - Rule-based clearances (e.g., "if duct > 400mm, use 100mm clearance")
   - Material-specific settings
   - Level-based overrides

## Troubleshooting

### Issue: Clearances not updating

**Check:**
1. Is CONDITIONS.xml being saved? (Check log for `[ConditionsService] Saved conditions`)
2. Is CONDITIONS.xml being loaded? (Check log for `[ConditionsService] Loaded conditions`)
3. Are clearances being read correctly? (Check log for `[GetClearanceFromUI] Using clearance`)

### Issue: Wrong clearance used

**Check:**
1. ClashZone.DuctShape is correct (should be "Round" or "Rectangular")
2. ClashZone.InsulationType is correct (should be "Normal" or "Insulated")
3. CONDITIONS.xml has correct values for all 4 clearance types

### Issue: Default clearances always used

**Check:**
1. Filter name matches between UI and command (e.g., "Ventilation_ducts")
2. CONDITIONS.xml file exists in correct location
3. XML is valid (no corruption)

## Summary

The CONDITIONS.xml architecture provides:
- ✅ Clean separation: DATA (ClashZones) vs SETTINGS (Conditions)
- ✅ Thread-safe: No UI dependencies
- ✅ Persistent: Settings survive restarts
- ✅ Scalable: Easy to add new settings
- ✅ Testable: Can mock conditions

This pattern follows industry best practices for Revit add-in development and ensures robust, maintainable code.


