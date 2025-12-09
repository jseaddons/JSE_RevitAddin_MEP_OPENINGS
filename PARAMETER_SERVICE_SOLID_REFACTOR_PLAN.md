# Parameter Service SOLID Refactoring Plan

## Problem Statement

**Current State:**
- `ParameterSnapshotService` violates SRP by mixing:
  1. Policy (whitelist management, what to capture)
  2. Mechanics (how to read parameters from Revit)
  3. Storage (learned keys persistence)
- `ParameterCaptureService` duplicates parameter reading logic
- Both services have similar fallback logic, category-specific handling, etc.

**Goal:**
- Split responsibilities according to SOLID principles
- Share common capture mechanics between both services
- Allow different policies (full snapshot vs. minimal refresh)
- Preserve all 28 features from comprehensive architecture

## Proposed Architecture

### 1. Interfaces (DIP Compliance)

```csharp
// Mechanics: HOW to read parameters
public interface IParameterCapture
{
    Parameter? LookupParameter(Element element, string parameterName);
    string ConvertParameterToString(Element element, Parameter parameter);
    List<SerializableKeyValue> CaptureParameters(Element element, IParameterPolicy policy);
}

// Policy: WHAT to capture
public interface IParameterPolicy
{
    HashSet<string> GetWhitelist();
    HashSet<string> GetMustCaptureKeys();
    bool ShouldCapture(string parameterName, Element element);
    string MapParameterName(string requestedName, Element element); // e.g., "System Type" -> "Service Type" for Cable Trays
}

// Storage: Learned keys persistence
public interface IParameterKeyStore
{
    HashSet<string> LoadLearnedKeys();
    void SaveLearnedKeys(HashSet<string> keys);
    void AddLearnedKey(string key);
}
```

### 2. Concrete Implementations

#### `RevitParameterCapture` (Mechanics)
- **Single Responsibility:** Read parameters from Revit elements
- **Responsibilities:**
  - `LookupParameter()` with fallbacks (built-in params, name variations)
  - `ConvertParameterToString()` (handles ElementId, Double, String, etc.)
  - `CaptureParameters()` (orchestrates capture using policy)
  - Category-specific fallbacks (Cable Tray, Duct, Pipe, etc.)
  - Diagnostic logging

#### `SnapshotParameterPolicy` (Full Policy)
- **Single Responsibility:** Define what parameters to capture for snapshots
- **Responsibilities:**
  - Full whitelist (ESSENTIAL_PARAMETERS + learned keys)
  - Must-capture keys
  - Category-specific mappings (System Type -> Service Type for Cable Trays)
  - Integration with `IParameterKeyStore` for learned keys

#### `MinimalParameterPolicy` (Refresh Policy)
- **Single Responsibility:** Define minimal parameters for refresh
- **Responsibilities:**
  - Minimal whitelist (10-15 parameters)
  - No learned keys
  - Same category mappings as snapshot policy

#### `FileParameterKeyStore` (Storage)
- **Single Responsibility:** Persist learned parameter keys
- **Responsibilities:**
  - Load/save learned keys from/to file
  - Thread-safe operations
  - File path management

### 3. Facade/Orchestrator

#### `ParameterSnapshotService` (Facade - Backward Compatible)
- **Single Responsibility:** Public API for snapshot capture
- **Responsibilities:**
  - Compose `IParameterCapture`, `IParameterPolicy`, `IParameterKeyStore`
  - Maintain backward-compatible public methods
  - Delegate to composed services

#### `ParameterCaptureService` (Refresh - Updated)
- **Single Responsibility:** Public API for refresh capture
- **Responsibilities:**
  - Compose `IParameterCapture` + `MinimalParameterPolicy`
  - No learned keys (refresh doesn't need persistence)
  - Delegate to composed services

## Implementation Steps

1. ✅ Create interfaces (`IParameterCapture`, `IParameterPolicy`, `IParameterKeyStore`)
2. ✅ Implement `RevitParameterCapture` (extract mechanics from `ParameterSnapshotService`)
3. ✅ Implement `SnapshotParameterPolicy` (extract policy from `ParameterSnapshotService`)
4. ✅ Implement `MinimalParameterPolicy` (extract from `ParameterCaptureService`)
5. ✅ Implement `FileParameterKeyStore` (extract learned keys logic)
6. ✅ Refactor `ParameterSnapshotService` to use composed services
7. ✅ Refactor `ParameterCaptureService` to use composed services
8. ✅ Update all callers (if needed)
9. ✅ Preserve all 28 features (diagnostic logging, deployment mode, etc.)

## File Structure

```
Services/
├── ParameterCapture/
│   ├── IParameterCapture.cs
│   ├── RevitParameterCapture.cs
│   ├── IParameterPolicy.cs
│   ├── SnapshotParameterPolicy.cs
│   ├── MinimalParameterPolicy.cs
│   ├── IParameterKeyStore.cs
│   └── FileParameterKeyStore.cs
├── ParameterSnapshotService.cs (Facade - backward compatible)
└── Refresh/
    └── ParameterCaptureService.cs (Updated to use shared mechanics)
```

## Benefits

1. **SRP Compliance:** Each class has one reason to change
2. **DRY:** No duplicate parameter reading logic
3. **Testability:** Each component can be tested independently
4. **Extensibility:** Easy to add new policies (e.g., `CustomParameterPolicy`)
5. **Maintainability:** Changes to capture mechanics don't affect policy
6. **Backward Compatibility:** Facade maintains existing public API

## Migration Strategy

1. Create new interfaces and implementations alongside existing code
2. Wire new services into facade (backward compatible)
3. Test thoroughly
4. Remove old code from facade (now just delegates)
5. Update `ParameterCaptureService` to use shared mechanics




