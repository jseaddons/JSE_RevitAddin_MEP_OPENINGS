# Opening Parameter Whitelist - Flexible Filter Implementation

## Overview
Modified the opening parameter dropdown in the Parameter Transfer dialog to use a **flexible whitelist filter** instead of hardcoded parameter names. The dropdown now automatically includes any parameter containing "size" or "system" (case-insensitive).

## What Changed

### Before (Hardcoded Whitelist)
```csharp
var essentialParams = new[] { 
    "MEP System Type", 
    "MEP Size", 
    "MEP System Abbreviation", 
    "Level" 
};
```

**Problem:** Only showed exactly 4 parameters, missing variations like:
- "System Size"
- "MEP System"
- "Opening Size"
- "System Classification"
- Any custom size/system parameters

### After (Flexible Whitelist)
```csharp
// ✅ FLEXIBLE WHITELIST: Collect all parameters containing "size" or "system" (case-insensitive)
var whitelistedParams = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

// Always include Level
whitelistedParams.Add("Level");

foreach (var familySymbol in openingFamilies)
{
    foreach (Parameter param in familySymbol.Parameters)
    {
        var paramName = param.Definition?.Name;
        if (!string.IsNullOrEmpty(paramName))
        {
            // Include if parameter name contains "size" OR "system" (case-insensitive)
            if (paramName.IndexOf("size", StringComparison.OrdinalIgnoreCase) >= 0 ||
                paramName.IndexOf("system", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                whitelistedParams.Add(paramName);
            }
        }
    }
}

// Sort alphabetically for better UX
var sortedParams = whitelistedParams.OrderBy(p => p).ToArray();
```

**Benefits:** Automatically detects ALL size and system-related parameters!

## Examples of Parameters Now Included

### Size-Related Parameters (case-insensitive):
- ✅ "MEP Size"
- ✅ "System Size"
- ✅ "Opening Size"
- ✅ "Sleeve Size"
- ✅ "Size"
- ✅ "Duct Size"
- ✅ "Pipe Size"
- ✅ "Cable Tray Size"
- ✅ Any custom parameter with "size" in the name

### System-Related Parameters (case-insensitive):
- ✅ "MEP System Type"
- ✅ "MEP System"
- ✅ "System Type"
- ✅ "System Classification"
- ✅ "System Abbreviation"
- ✅ "MEP System Abbreviation"
- ✅ "System Name"
- ✅ Any custom parameter with "system" in the name

### Always Included:
- ✅ "Level" (commonly used for reference level)

## Technical Details

### File Modified:
- `Views/ParameterServiceDialogV2.cs`
- Method: `GetOpeningParametersForCategory(string category)`

### Filter Logic:
```csharp
// Case-insensitive check for "size" OR "system"
if (paramName.IndexOf("size", StringComparison.OrdinalIgnoreCase) >= 0 ||
    paramName.IndexOf("system", StringComparison.OrdinalIgnoreCase) >= 0)
{
    whitelistedParams.Add(paramName);
}
```

### Sorting:
- Parameters are sorted alphabetically for better user experience
- Makes it easier to find specific parameters in the dropdown

### Performance:
- Scans all opening families in the active document
- Uses `HashSet<string>` with `StringComparer.OrdinalIgnoreCase` to avoid duplicates
- Efficient case-insensitive comparison

## User Experience

### Before:
1. User opens Parameter Transfer dialog
2. Opening parameter dropdown shows only 4 hardcoded options
3. User cannot select "System Size" even if it exists in their opening family
4. User frustrated 😞

### After:
1. User opens Parameter Transfer dialog
2. Opening parameter dropdown automatically shows ALL size/system parameters
3. User can select any variation: "MEP Size", "System Size", "Opening Size", etc.
4. User happy 😊

## Example Dropdown Contents

**Typical Opening Family might show:**
```
Level
MEP Size
MEP System
MEP System Abbreviation
MEP System Type
Opening Size
System Classification
System Size
System Type
```

**All sorted alphabetically!**

## Edge Cases Handled

1. **No opening families found:** Returns basic fallback `["MEP System Type", "MEP Size", "Level"]`
2. **Document is null:** Returns basic fallback
3. **Exception during scan:** Returns basic fallback with error logging
4. **Duplicate parameters:** `HashSet` with case-insensitive comparer prevents duplicates
5. **Empty parameter names:** Skipped with `!string.IsNullOrEmpty(paramName)` check

## Logging

Debug logs show:
```
[ParameterServiceDialogV2] Scanning 4 opening families for size/system parameters
[ParameterServiceDialogV2] Found 8 whitelisted opening parameters (containing 'size' or 'system')
[ParameterServiceDialogV2] Whitelisted opening parameters: Level, MEP Size, MEP System, MEP System Abbreviation, MEP System Type, Opening Size, System Size, System Type
```

## Testing Checklist

- [ ] Verify dropdown shows all parameters containing "size" (case-insensitive)
- [ ] Verify dropdown shows all parameters containing "system" (case-insensitive)
- [ ] Verify "Level" is always included
- [ ] Verify parameters are sorted alphabetically
- [ ] Verify no duplicates in dropdown
- [ ] Test with opening families that have custom size/system parameters
- [ ] Test with opening families that have no size/system parameters (should show fallback)
- [ ] Verify parameter transfer works with newly visible parameters

## Benefits Summary

✅ **Flexibility:** Automatically adapts to any opening family parameter schema
✅ **Completeness:** Shows ALL size and system-related parameters, not just hardcoded ones
✅ **User-Friendly:** Alphabetically sorted for easy navigation
✅ **Robust:** Handles edge cases with fallback values
✅ **Future-Proof:** Works with custom parameters without code changes
