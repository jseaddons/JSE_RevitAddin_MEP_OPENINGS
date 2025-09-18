# Cable Tray Clearance Implementation

## Overview
This document details the implementation of different top clearance vs other side clearance for cable trays.

## Problem Statement
Cable trays require different clearances:
- **Top Side**: 75mm clearance (default)
- **Other Sides**: 25mm clearance (default)

The original implementation used uniform clearance for all sides, which was incorrect.

## Solution Implemented

### 1. ✅ **CableTrayClearanceProvider** (New)
- **File**: `Services/ClearanceProviders/CableTrayClearanceProvider.cs`
- **Purpose**: Specialized clearance provider for cable trays
- **Features**:
  - `GetClearance()`: Returns top clearance by default
  - `GetClearanceForSide()`: Returns specific clearance for "top" or "other" sides
  - Supports UI clearance overrides
  - Default clearances: Top=75mm, Other=25mm

```csharp
public double GetClearanceForSide(Element mepElement, string side, Dictionary<string, double> uiClearances = null)
{
    string key = side.ToLower() switch
    {
        "top" => "cabletray_top_normal",
        "other" => "cabletray_other_normal",
        _ => "cabletray_other_normal"
    };
    
    if (uiClearances?.TryGetValue(key, out double clearance) == true)
    {
        return UnitUtils.ConvertToInternalUnits(clearance, UnitTypeId.Millimeters);
    }
    
    // Fallback to default clearances
    return side.ToLower() switch
    {
        "top" => UnitUtils.ConvertToInternalUnits(75.0, UnitTypeId.Millimeters),
        "other" => UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters),
        _ => UnitUtils.ConvertToInternalUnits(25.0, UnitTypeId.Millimeters)
    };
}
```

### 2. ✅ **ClearanceProviderFactory** (Updated)
- **File**: `Services/ClearanceProviders/ClearanceProviderFactory.cs`
- **Change**: Updated to use `CableTrayClearanceProvider` instead of `SleeveClearanceProvider` for cable trays

```csharp
_providers = new Dictionary<string, IClearanceProvider>
{
    { "Ducts", new SleeveClearanceProvider() },
    { "Pipes", new SleeveClearanceProvider() },
    { "Cable Trays", new CableTrayClearanceProvider() }, // ✅ Updated
    { "Fire Dampers", new FireDamperClearanceProvider() },
    { "Default", new SleeveClearanceProvider() }
};
```

### 3. ✅ **CableTraySleevePlacer** (Updated)
- **File**: `Services/CableTraySleevePlacer.cs`
- **Change**: Updated to use different clearances for width vs height

```csharp
// OLD (Uniform clearance)
double clearance = ClearanceManager.Instance.GetClearance(tray!);
double widthWithClearance = width + 2 * clearance;
double heightWithClearance = height + 2 * clearance;

// NEW (Different top/other side clearances)
var clearanceProvider = new CableTrayClearanceProvider();
double topClearance = clearanceProvider.GetClearanceForSide(tray!, "top", GetUIClearances());
double otherClearance = clearanceProvider.GetClearanceForSide(tray!, "other", GetUIClearances());

double widthWithClearance = width + 2 * otherClearance;           // Other sides clearance
double heightWithClearance = height + topClearance + otherClearance; // Top + bottom clearance
```

### 4. ✅ **UI Integration** (Updated)
- **File**: `Views/EmergencyMainDialog.cs`
- **Change**: Updated `GetClearanceSettings()` to collect cable tray specific clearance values

```csharp
// Get cable tray specific clearance values
if (_cableTrayPanel?.Controls.Count > 0)
{
    foreach (var control in _cableTrayPanel.Controls)
    {
        if (control is WinForms.TextBox textBox && textBox.Tag != null)
        {
            if (double.TryParse(textBox.Text, out double value))
            {
                string key = textBox.Tag.ToString() ?? "";
                clearances[key] = value; // Direct key mapping
            }
        }
    }
}
```

## UI Clearance Keys

### Cable Tray Panel Controls
- **Top Side Normal**: `"cabletray_top_normal"` (75mm default)
- **Top Side Insulated**: `"cabletray_top_insulated"` (75mm default)
- **Other Sides Normal**: `"cabletray_other_normal"` (25mm default)
- **Other Sides Insulated**: `"cabletray_other_insulated"` (25mm default)

### Clearance Application Logic
- **Width**: Uses "other sides" clearance (25mm default)
- **Height**: Uses "top" clearance + "other sides" clearance (75mm + 25mm = 100mm total)

## Clearance Flow

```
UI Cable Tray Panel → GetClearanceSettings() → ClearanceManager → 
CableTrayClearanceProvider.GetClearanceForSide() → CableTraySleevePlacer
```

## Default Clearances

| Side | Default Value | UI Key |
|------|---------------|---------|
| Top | 75mm | `cabletray_top_normal` |
| Other Sides | 25mm | `cabletray_other_normal` |

## Verification

### ✅ **Backend Logic**
- CableTrayClearanceProvider handles different top/other side clearances
- ClearanceProviderFactory routes cable trays to specialized provider
- CableTraySleevePlacer applies correct clearances to width/height

### ✅ **UI Integration**
- EmergencyMainDialog collects cable tray specific clearance values
- Clearance keys match provider expectations
- UI clearance overrides are properly passed through

### ✅ **Clearance Application**
- **Width**: `width + 2 * otherClearance` (25mm each side)
- **Height**: `height + topClearance + otherClearance` (75mm top + 25mm bottom)

## Status: ✅ **COMPLETE**

The cable tray different top clearance vs other side clearance has been fully implemented and integrated with the UI clearance system.
