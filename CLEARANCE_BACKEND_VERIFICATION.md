# Clearance Backend Verification

## Current Status Analysis

### ✅ **What's Working Correctly**

1. **ClearanceManager**: Singleton properly implemented
2. **ClearanceProviderFactory**: Correctly maps MEP categories to providers
3. **SleeveClearanceProvider**: Handles Ducts, Pipes, Cable Trays with insulation detection
4. **FireDamperClearanceProvider**: Handles Fire Dampers with MSFD logic
5. **Service Integration**: All sleeve placer services use `ClearanceManager.Instance.GetClearance()`

### ❌ **Issue Identified: Clearance Key Mismatch**

#### UI Clearance Keys (EmergencyMainDialog.cs)
```csharp
Tag = "normal_clearance"      // Generic key
Tag = "insulated_clearance"    // Generic key
```

#### Provider Clearance Keys (SleeveClearanceProvider.cs)
```csharp
return $"{category.ToLower().Replace(" ", "_")}_{insulationType}_clearance";
// Examples:
// "ducts_normal_clearance"
// "pipes_insulated_clearance" 
// "cable_trays_normal_clearance"
```

#### Fire Damper Keys (FireDamperClearanceProvider.cs)
```csharp
return isMSFD ? "fire_damper_msfd_clearance" : "fire_damper_standard_clearance";
```

## **Problem**: UI keys don't match provider keys!

### Current Flow (Broken)
1. UI sets: `clearances["normal_clearance"] = 50`
2. Provider looks for: `clearances["ducts_normal_clearance"]` → **NOT FOUND**
3. Falls back to existing logic

### Required Flow (Fixed)
1. UI sets: `clearances["ducts_normal_clearance"] = 50`
2. Provider finds: `clearances["ducts_normal_clearance"]` → **FOUND**
3. Uses UI override

## **Solution Required**

The UI needs to generate category-specific clearance keys that match what the providers expect.

### Option 1: Update UI to generate specific keys
- Modify `GetClearanceSettings()` to generate category-specific keys
- Use current MEP category to create proper keys

### Option 2: Update providers to handle generic keys
- Modify providers to also check generic keys as fallback
- Keep existing specific keys for precision

## **Recommendation**: Option 1 (Update UI)

The UI should generate proper clearance keys based on the current MEP category selection, ensuring exact matching with provider expectations.

## **MEP Categories Coverage**

### ✅ **Ducts**
- Provider: `SleeveClearanceProvider`
- Keys: `"ducts_normal_clearance"`, `"ducts_insulated_clearance"`
- Service: `DuctSleevePlacerService` ✅ Uses `ClearanceManager.Instance.GetClearance(duct)`

### ✅ **Pipes** 
- Provider: `SleeveClearanceProvider`
- Keys: `"pipes_normal_clearance"`, `"pipes_insulated_clearance"`
- Service: `PipeSleevePlacerService` ✅ Uses `ClearanceManager.Instance.GetClearance(pipe)`

### ✅ **Cable Trays**
- Provider: `SleeveClearanceProvider` 
- Keys: `"cable_trays_normal_clearance"`, `"cable_trays_insulated_clearance"`
- Service: `CableTraySleevePlacer` ✅ Uses `ClearanceManager.Instance.GetClearance(tray)`

### ✅ **Fire Dampers**
- Provider: `FireDamperClearanceProvider`
- Keys: `"fire_damper_msfd_clearance"`, `"fire_damper_standard_clearance"`
- Service: `FireDamperSleevePlacerService` ✅ Uses `ClearanceManager.Instance.GetClearance(accessory)`

## **Status**: ✅ **FIXED** - Backend logic is correct and UI clearance key generation now matches provider expectations!

### **Fix Applied**
- Updated `GetClearanceSettings()` to generate category-specific clearance keys
- Added `ConvertToSpecificClearanceKey()` method to map generic keys to specific keys
- Added `GetCurrentMepCategory()` method to get current MEP category from UI selection
- UI now generates proper keys that match what providers expect

### **Key Mapping Examples**
- UI: `"normal_clearance"` + Category: `"Ducts"` → Provider Key: `"ducts_normal_clearance"`
- UI: `"insulated_clearance"` + Category: `"Pipes"` → Provider Key: `"pipes_insulated_clearance"`
- UI: `"normal_clearance"` + Category: `"Duct Accessories"` → Provider Key: `"fire_damper_standard_clearance"`
