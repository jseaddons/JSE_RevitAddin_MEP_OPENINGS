# Revit API Compatibility Guidelines (2023-2026)

This project targets multiple versions of the Revit API which undergo significant architectural changes, specifically the transition of `ElementId` and `BuiltInParameter` from `int` to `long`.

## 🚫 Forbidden Patterns
Do NOT use the following patterns anywhere in the codebase (except within `ElementIdCompat.cs`):

- ❌ `new ElementId(int value)` or `new ElementId(long value)`
- ❌ `elementId.Value` (R24+)
- ❌ `elementId.IntegerValue` (R23-)
- ❌ Direct casts to `BuiltInParameter` from `int` or `long`.

## ✅ Mandatory Patterns
Always use the `JSE_RevitAddin_MEP_OPENINGS.Helpers.ElementIdCompat` class:

### 1. Creating ElementIds
```csharp
// Instead of: new ElementId(idValue)
var id = ElementIdCompat.FromValue(idValue);
```

### 2. Getting Numeric Values
```csharp
// Instead of: id.Value or id.IntegerValue
long val = id.GetIdValue();
```

### 3. Handling Parameters
```csharp
// Instead of: (BuiltInParameter)paramId
BuiltInParameter bip = ElementIdCompat.GetBip(paramId);
```

## Why this is necessary?
Starting with Revit 2024, Autodesk changed `ElementId` to use 64-bit integers. Code using `IntegerValue` or `int` constructors will break in R24+. Code using `.Value` will break in R23.

Using `ElementIdCompat` ensures the code builds and runs correctly across all supported versions without manual intervention.
