# Parameter Service Dialog Enhancements - Summary

## Overview
Enhanced the Parameter Service Dialog with three key improvements:
1. **Default parameters always visible** in dropdowns
2. **Shared parameters prioritized** in search
3. **MEP reference parameters** included in search

---

## 1. Default Parameters Always Visible

### Problem
Default parameter mappings (MEP Size, MEP System Type, etc.) weren't showing in the opening parameter dropdown if they didn't exist in the opening families.

### Solution
Modified `AddParameterRow` to ensure the specified opening parameter is always added to the dropdown, even if it's not found in the opening families.

**Code Location:** `Views/ParameterServiceDialogV2.cs` (lines ~920-935)

```csharp
// ✅ FIX: Ensure the specified opening parameter is always in the dropdown
// This handles cases where default parameters (MEP Size, MEP System Type, etc.) 
// aren't in the opening family but should still be selectable
if (!string.IsNullOrEmpty(openingParam) && !openingCombo.Items.Contains(openingParam))
{
    openingCombo.Items.Add(openingParam);
}
```

**Result:** Default mappings now always appear:
- `Size → MEP Size` ✅
- `System Type → MEP System Type` ✅
- `System Abbreviation → MEP System Abbreviation` ✅
- `Reference Level → Level` ✅

---

## 2. Shared Parameters Prioritized

### Problem
Users want to see **shared parameters** first in the search dialog, as these are the main parameters they work with.

### Solution
Modified `GetAllOpeningParameters` to track which parameters are shared and sort them to the top.

**New Method:** `GetAllOpeningParametersWithType()`

**Code Location:** `Views/ParameterServiceDialogV2.cs` (lines ~2543-2597)

```csharp
// Check if it's a shared parameter
bool isShared = param.IsShared;

allParams.Add((paramName, isShared));

// ✅ SORT: Shared parameters first (alphabetically), then non-shared (alphabetically)
allParams = allParams
    .OrderByDescending(p => p.isShared) // Shared first
    .ThenBy(p => p.name) // Then alphabetically
    .ToList();
```

**Visual Indicators:**
- 🔗 **Shared Parameter** (shown first)
- 📄 **Opening Parameter** (instance/type)
- 🔧 **MEP Reference Parameter** (from linked files)

---

## 3. MEP Reference Parameters Included

### Problem
Search dialog only showed opening family parameters, not MEP parameters from reference elements (linked files).

### Solution
Extended search to include both opening parameters AND MEP parameters from reference elements.

**Code Location:** `Views/ParameterServiceDialogV2.cs` (ShowParameterSearchDialog method)

```csharp
// Get ALL parameters from opening families (shared parameters prioritized)
var openingParameters = GetAllOpeningParametersWithType();

// Get MEP parameters from reference elements (linked files)
var mepParameters = GetMepParametersForCategory(category);

// Combine into a single list with source indication
var allParameters = new List<(string display, string actualName, string source)>();

// Add opening parameters (with shared indicator)
foreach (var param in openingParameters)
{
    string prefix = param.isShared ? "🔗 " : "📄 ";
    string display = $"{prefix}{param.name}";
    allParameters.Add((display, param.name, param.isShared ? "Shared" : "Opening"));
}

// Add MEP parameters from reference elements
foreach (var mepParam in mepParameters)
{
    // Avoid duplicates (case-insensitive)
    if (!openingParameters.Any(op => string.Equals(op.name, mepParam, StringComparison.OrdinalIgnoreCase)))
    {
        string display = $"🔧 {mepParam}";
        allParameters.Add((display, mepParam, "MEP"));
    }
}
```

**Dialog Title:** "Search Parameters (Opening + MEP Reference)"

**Info Bar:** Shows legend for parameter types

---

## Search Dialog Features

### Visual Layout
```
┌─────────────────────────────────────────────────────┐
│ Search Parameters (Opening + MEP Reference)         │
├─────────────────────────────────────────────────────┤
│ 🔗 Shared Parameter | 📄 Opening | 🔧 MEP Reference │
│                                                      │
│ Search for parameter:                                │
│ [_______________________________________________]    │
│                                                      │
│ Found 45 parameters:                                 │
│ ┌──────────────────────────────────────────────┐   │
│ │ 🔗 MEP Size                                   │   │
│ │ 🔗 MEP System Type                            │   │
│ │ 🔗 MEP System Abbreviation                    │   │
│ │ 📄 Width                                       │   │
│ │ 📄 Height                                      │   │
│ │ 🔧 System Classification                      │   │
│ │ 🔧 System Name                                 │   │
│ └──────────────────────────────────────────────┘   │
│                                                      │
│                              [Select]  [Cancel]      │
└─────────────────────────────────────────────────────┘
```

### Features
1. **Real-time search** - Filter as you type
2. **Double-click to select** - Quick selection
3. **Visual indicators** - Know the parameter source
4. **Sorted by priority** - Shared parameters first
5. **No duplicates** - Avoids showing same parameter twice
6. **Auto-add to dropdown** - Selected parameter added if not present

---

## Parameter Sources

### 1. Shared Parameters (🔗)
- **Source:** Opening family shared parameters
- **Priority:** Highest (shown first)
- **Example:** MEP Size, MEP System Type
- **Use Case:** Main parameters for coordination

### 2. Opening Parameters (📄)
- **Source:** Opening family instance/type parameters
- **Priority:** Medium
- **Example:** Width, Height, Depth, Comments
- **Use Case:** Family-specific parameters

### 3. MEP Reference Parameters (🔧)
- **Source:** MEP elements from linked files
- **Priority:** Lower (shown after opening parameters)
- **Example:** System Classification, System Name, Nominal Diameter
- **Use Case:** Parameters from MEP elements that aren't in opening families

---

## Benefits

### For Users
✅ **Always see default parameters** - No more missing MEP Size, MEP System Type
✅ **Find shared parameters easily** - Highlighted with 🔗 icon
✅ **Access MEP parameters** - Can map any parameter from reference elements
✅ **Visual clarity** - Icons show parameter source
✅ **Fast search** - Real-time filtering

### For Developers
✅ **Robust parameter handling** - Always adds default parameters
✅ **Flexible search** - Combines multiple parameter sources
✅ **Clear code structure** - Separate methods for each source
✅ **Good logging** - Tracks shared vs non-shared parameters

---

## Example Usage

### Scenario 1: Default Mapping
**User Action:** Opens Parameter Service Dialog
**Result:** Sees default mappings:
- Size → MEP Size ✅
- System Type → MEP System Type ✅
- System Abbreviation → MEP System Abbreviation ✅
- Reference Level → Level ✅

### Scenario 2: Search for Shared Parameter
**User Action:** Clicks 🔍 search button
**Result:** Dialog shows:
- 🔗 MEP Size (shared, shown first)
- 🔗 MEP System Type (shared, shown first)
- 📄 Width (opening parameter)
- 🔧 System Classification (MEP reference)

### Scenario 3: Search for MEP Parameter
**User Action:** Types "classification" in search
**Result:** Filters to:
- 🔧 System Classification (from MEP reference)

---

## Files Modified

1. **Views/ParameterServiceDialogV2.cs**
   - `AddParameterRow` - Ensures default parameters always added
   - `GetAllOpeningParametersWithType` - Tracks shared parameters
   - `ShowParameterSearchDialog` - Shows opening + MEP parameters

---

## Testing Checklist

- [ ] Default parameters appear in dropdown on startup
- [ ] Search dialog shows shared parameters with 🔗 icon
- [ ] Search dialog shows opening parameters with 📄 icon
- [ ] Search dialog shows MEP parameters with 🔧 icon
- [ ] Shared parameters appear first in search results
- [ ] Real-time search filtering works
- [ ] Double-click selects parameter
- [ ] Select button adds parameter to dropdown
- [ ] No duplicate parameters in search results
- [ ] Dialog title shows "Opening + MEP Reference"
- [ ] Info bar shows legend

---

## Known Limitations

### 1. **Shared Parameter Detection**
- Uses `param.IsShared` property
- Only works for parameters in opening families
- MEP reference parameters are not checked for shared status

### 2. **Duplicate Handling**
- Case-insensitive comparison
- If MEP parameter has same name as opening parameter, MEP version is hidden
- Opening parameter takes precedence

---

## Future Enhancements

1. **Color coding** - Different colors for each parameter type
2. **Grouping** - Group by parameter source in listbox
3. **Favorites** - Star frequently used parameters
4. **Recent** - Show recently selected parameters
5. **Tooltips** - Show parameter details on hover

---

## Summary

✅ **Default parameters always visible** - Fixed missing MEP Size, MEP System Type
✅ **Shared parameters prioritized** - 🔗 icon, shown first
✅ **MEP reference parameters included** - 🔧 icon, from linked files
✅ **Visual clarity** - Icons show parameter source
✅ **Robust search** - Combines multiple sources, no duplicates
