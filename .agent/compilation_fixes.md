# Compilation Error Fixes

## Errors Fixed

### 1. CS0841: Cannot use local variable 'searchBtn' before it is declared
**Location:** Line 935
**Problem:** The `searchBtn` variable was referenced in the `row.Resize` event handler before it was declared.

**Fix:** Moved the `searchBtn` declaration BEFORE the resize event handler.

**Before:**
```csharp
row.Resize += (s, e) => {
    searchBtn.Location = new Point(...); // ERROR: searchBtn not declared yet
};

var searchBtn = new WinForms.Button { ... };
```

**After:**
```csharp
var searchBtn = new WinForms.Button { ... }; // Declare FIRST

row.Resize += (s, e) => {
    searchBtn.Location = new Point(...); // OK: searchBtn is now declared
};
```

---

### 2. CS0117: 'TextBox' does not contain a definition for 'PlaceholderText'
**Location:** Line 2383
**Problem:** `PlaceholderText` property is not available in WinForms `TextBox` (only in WPF).

**Fix:** Removed the `PlaceholderText` property and added a comment explaining why.

**Before:**
```csharp
var searchTextBox = new WinForms.TextBox
{
    PlaceholderText = "Type to filter parameters..." // ERROR: Not available in WinForms
};
```

**After:**
```csharp
var searchTextBox = new WinForms.TextBox
{
    // Note: PlaceholderText not available in WinForms TextBox (only in WPF)
};
```

**Alternative (if placeholder is needed):**
- Use `GotFocus` and `LostFocus` events to simulate placeholder text
- Or use a label overlay
- Or use a third-party control

---

### 3. CS0136: A local or parameter named 'transferDebugLogPath' cannot be declared in this scope
**Location:** Line 1363
**Problem:** Variable `transferDebugLogPath` was already declared in an outer scope, causing a naming conflict.

**Fix:** Renamed the inner variable to `activeViewLogPath` to avoid the conflict.

**Before:**
```csharp
// Outer scope
string transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");

// ... later in code ...

if (_activeViewOnlyCheckBox.Checked)
{
    string transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log"); // ERROR: Already declared
}
```

**After:**
```csharp
// Outer scope
string transferDebugLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log");

// ... later in code ...

if (_activeViewOnlyCheckBox.Checked)
{
    string activeViewLogPath = SafeFileLogger.GetLogFilePath("transfer_debug.log"); // OK: Different name
}
```

---

## Summary

All three compilation errors have been resolved:
- ✅ Fixed variable declaration order (searchBtn)
- ✅ Removed unsupported property (PlaceholderText)
- ✅ Renamed duplicate variable (transferDebugLogPath → activeViewLogPath)

The code should now compile successfully!
