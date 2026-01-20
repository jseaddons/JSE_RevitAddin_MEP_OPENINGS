# Parameter Search Feature - Implementation Summary

## Overview
Added a **🔍 Search button** next to the opening parameter dropdown in the Parameter Transfer dialog. This allows users to search for and select ANY parameter from opening families, not just those in the whitelist.

## What Was Added

### 1. Search Button (🔍)
**Location:** Next to the opening parameter dropdown in each parameter mapping row

**Visual:**
```
[MEP Parameter ▼] → [Opening Parameter ▼] [🔍] [×]
```

**Features:**
- Blue background color for visibility
- Magnifying glass emoji (🔍) icon
- Tooltip: "Search for parameters not in whitelist (scans all opening family parameters)"
- Hand cursor on hover

### 2. Parameter Search Dialog
**Triggered by:** Clicking the 🔍 button

**Dialog Features:**
- **Title:** "Search Opening Parameters"
- **Resizable:** Yes (minimum 400x300)
- **Search Textbox:** Real-time filtering as you type
- **Results Listbox:** Shows all matching parameters (alphabetically sorted)
- **Results Count:** Dynamic label showing "Found X parameters matching 'search term'"
- **Select Button:** Adds selected parameter to dropdown and selects it
- **Cancel Button:** Closes dialog without changes
- **Double-Click:** Quick selection by double-clicking a parameter

## User Experience Flow

### Step 1: Click Search Button
```
User clicks 🔍 button next to opening parameter dropdown
↓
Search dialog opens with ALL opening family parameters
```

### Step 2: Search for Parameter
```
User types in search box (e.g., "width")
↓
Results filter in real-time to show matching parameters:
- "Width"
- "Opening Width"
- "Sleeve Width"
- "Max Width"
etc.
```

### Step 3: Select Parameter
```
User double-clicks parameter OR selects and clicks "Select" button
↓
Parameter is added to dropdown (if not already there)
↓
Parameter is automatically selected
↓
Dialog closes
```

## Technical Implementation

### Files Modified:
- `Views/ParameterServiceDialogV2.cs`

### New Methods Added:

#### 1. `ShowParameterSearchDialog(ComboBox targetComboBox, string category)`
- Creates and displays the search dialog
- Handles search filtering
- Manages parameter selection
- Updates the target dropdown

#### 2. `GetAllOpeningParameters()`
- Scans ALL parameters from all opening families
- Returns `HashSet<string>` with case-insensitive comparison
- Used to populate the search dialog

### Search Logic:
```csharp
// Real-time filtering (case-insensitive)
var filtered = allParameters
    .Where(p => p.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
    .OrderBy(p => p)
    .ToList();
```

### Selection Logic:
```csharp
// Add to combo box if not already there
if (!targetComboBox.Items.Contains(selectedParam))
{
    targetComboBox.Items.Add(selectedParam);
}

// Select the parameter
targetComboBox.SelectedItem = selectedParam;
```

## Example Use Cases

### Use Case 1: Finding Custom Parameters
**Scenario:** User has a custom parameter "Fire Rating Required" in their opening family

**Before:** Not in whitelist → Cannot select
**After:** 
1. Click 🔍
2. Type "fire"
3. See "Fire Rating Required" in results
4. Double-click to select
5. ✅ Parameter now available in dropdown

### Use Case 2: Finding Dimension Parameters
**Scenario:** User wants to map to "Opening Width" instead of "MEP Size"

**Before:** Only "MEP Size" in whitelist
**After:**
1. Click 🔍
2. Type "width"
3. See all width-related parameters:
   - "Width"
   - "Opening Width"
   - "Sleeve Width"
   - "Max Width"
4. Select "Opening Width"
5. ✅ Parameter mapped successfully

### Use Case 3: Exploring Available Parameters
**Scenario:** User doesn't know what parameters are available

**Before:** Limited to whitelist, can't explore
**After:**
1. Click 🔍
2. Leave search box empty
3. Scroll through ALL parameters alphabetically
4. Discover available parameters
5. Select the one needed

## UI Layout Changes

### Before:
```
[MEP Parameter ▼] → [Opening Parameter (wider) ▼] [×]
```

### After:
```
[MEP Parameter ▼] → [Opening Parameter ▼] [🔍] [×]
```

**Width Adjustment:**
- Opening parameter dropdown: Reduced by 30px to make room for search button
- Search button: 24px wide
- Spacing: 2px between dropdown and search button

## Search Dialog Layout

```
┌─────────────────────────────────────────────┐
│ Search Opening Parameters                   │
├─────────────────────────────────────────────┤
│ Search for parameter:                       │
│ [Type to filter parameters...            ] │
│                                             │
│ Found 45 parameters:                        │
│ ┌─────────────────────────────────────────┐ │
│ │ Level                                   │ │
│ │ MEP Mark                                │ │
│ │ MEP Size                                │ │
│ │ MEP System                              │ │
│ │ MEP System Abbreviation                 │ │
│ │ MEP System Type                         │ │
│ │ Opening Height                          │ │
│ │ Opening Width                           │ │
│ │ ... (scrollable)                        │ │
│ └─────────────────────────────────────────┘ │
│                                             │
│                        [Select]   [Cancel] │
└─────────────────────────────────────────────┘
```

## Features Summary

✅ **Real-Time Search:** Filter as you type
✅ **Case-Insensitive:** Finds "SIZE", "size", "Size", etc.
✅ **Alphabetically Sorted:** Easy to browse
✅ **Double-Click Selection:** Quick selection
✅ **Dynamic Results Count:** Shows how many parameters match
✅ **Resizable Dialog:** Adjust size as needed
✅ **Auto-Focus:** Search box focused on open
✅ **Persistent Selection:** Parameter stays in dropdown after selection

## Benefits

### 🎯 Flexibility
- Access ANY parameter from opening families
- Not limited to whitelist
- Discover parameters you didn't know existed

### 🔍 Discoverability
- Browse all available parameters
- Search by partial name
- See exactly what's available in your families

### ⚡ Efficiency
- Real-time filtering saves time
- Double-click for quick selection
- No need to remember exact parameter names

### 🛡️ Robustness
- Handles missing families gracefully
- Error handling for edge cases
- Clear user feedback

## Testing Checklist

- [ ] Click 🔍 button opens search dialog
- [ ] Search dialog shows all opening family parameters
- [ ] Typing in search box filters results in real-time
- [ ] Results count updates correctly
- [ ] Double-clicking a parameter selects it and closes dialog
- [ ] "Select" button adds parameter to dropdown
- [ ] "Cancel" button closes dialog without changes
- [ ] Selected parameter appears in dropdown
- [ ] Selected parameter is automatically selected
- [ ] Dialog is resizable
- [ ] Search is case-insensitive
- [ ] Results are alphabetically sorted
- [ ] Works with opening families that have many parameters
- [ ] Handles case when no opening families are found

## Edge Cases Handled

1. **No opening families found:** Shows message "No opening families found in the active document"
2. **Document is null:** Shows error message
3. **No parameter selected:** Shows message "Please select a parameter from the list"
4. **Parameter already in dropdown:** Doesn't add duplicate, just selects it
5. **Empty search:** Shows all parameters
6. **No matches:** Shows empty list with "Found 0 parameters matching..."

## Future Enhancements (Optional)

- Add category filter to search only parameters from specific MEP categories
- Add parameter type filter (text, number, yes/no, etc.)
- Add "Recently Used" section at top of results
- Add "Add to Whitelist" button to permanently add parameter to whitelist
- Support searching parameters from linked files (not just active document)
