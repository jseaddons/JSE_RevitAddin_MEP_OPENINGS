# Filter XML Overwrite Issues - Identification and Fix Plan

## 🔴 CRITICAL ISSUES IDENTIFIED

### Issue 1: Line 1332 - 3-Point Validation Overwrites Existing Clash Zones

**Location:** `Services/RefreshService.cs` line 1332

**Problem:**
```csharp
// Replace with validated clash zones only
existingClashZones.ClashZones = validClashZones;
```

**Issue:**
- This REPLACES the entire `existingClashZones.ClashZones` list with only `validClashZones`
- Invalid zones are removed EVEN if "Adopt Document" is disabled
- Should ONLY remove zones if `enableThreePointValidation == true` AND validation failed

**Fix Required:**
- Only remove invalid zones if `enableThreePointValidation == true`
- If disabled, keep ALL existing zones (append-only behavior)

---

### Issue 2: Line 3324 - Filter XML Storage Overwritten

**Location:** `Services/RefreshService.cs` line 3324

**Problem:**
```csharp
targetFilter.ClashZoneStorage = clashZoneStorage;
```

**Issue:**
- This OVERWRITES the entire `ClashZoneStorage` with a new instance
- Even if `allClashZones` includes existing zones, this replaces the reference
- Should append to existing storage, not replace

**Fix Required:**
- Check if `targetFilter.ClashZoneStorage` exists
- If exists, merge `allClashZones` with existing zones (append-only)
- If not exists, create new storage

---

### Issue 3: Line 3708 - Category-Specific Filter XML Overwritten

**Location:** `Services/RefreshService.cs` line 3708

**Problem:**
```csharp
_filterManagementService.SaveFilterToXmlFile(categoryFilter, categoryFilePath);
```

**Issue:**
- This saves `categoryFilter` which may have been created from scratch
- Should load existing category-specific XML and merge with new zones
- Currently overwrites entire category XML file

**Fix Required:**
- Load existing category-specific XML first
- Merge new zones with existing zones
- Save merged list (append-only)

---

### Issue 4: Lines 3329-3330 - UI State Saved to Filter XML

**Location:** `Services/RefreshService.cs` lines 3329-3330

**Problem:**
```csharp
targetFilter.SelectedReferenceFiles = selectedReferenceFiles;
targetFilter.SelectedHostFiles = selectedHostFiles;
```

**Issue:**
- UI state (selected files) is being saved to Filter XML
- Should ONLY be saved to Global XML (`ProcessedFileCombos`)
- Filter XML should contain ONLY placement data

**Fix Required:**
- Remove UI state saving from Filter XML
- Only save UI state to Global XML at line 3741 (`MarkFileCombosAsProcessed`)

---

### Issue 5: Line 1172 - Section Box Filter Overwrites Existing Zones

**Location:** `Services/RefreshService.cs` line 1172

**Problem:**
```csharp
existingClashZones.ClashZones = kept;
```

**Issue:**
- Section box filtering REPLACES entire list
- Should only filter for display/processing, not permanently remove from storage
- Zones outside section box should remain in Filter XML

**Fix Required:**
- Use filtered list for processing only
- Keep original list for saving to Filter XML
- Append-only means zones stay in XML even if outside section box

---

## ✅ CORRECT BEHAVIOR (Append-Only)

### Filter XML Should:
1. ✅ Load existing clash zones from XML
2. ✅ Merge new zones with existing zones (avoid duplicates by GUID)
3. ✅ Only remove zones if "Adopt Document" enabled AND 3-point validation failed
4. ✅ Save merged list (append-only, never overwrite)

### Filter XML Should NOT:
1. ❌ Overwrite entire storage
2. ❌ Remove zones unless validation explicitly failed
3. ❌ Save UI state (only placement data)
4. ❌ Clear zones on rerun

---

## 🔧 FIX PLAN

### Step 1: Fix 3-Point Validation Logic
- Only remove zones if `enableThreePointValidation == true` AND validation failed
- If disabled, keep ALL existing zones

### Step 2: Implement Append-Only Merging
- Load existing Filter XML clash zones
- Merge with new zones (avoid duplicates by GUID)
- Save merged list (never overwrite)

### Step 3: Remove UI State from Filter XML
- Remove lines 3329-3330 (UI state saving)
- Ensure UI state only saved to Global XML at line 3741

### Step 4: Fix Category-Specific XML Saving
- Load existing category XML before saving
- Merge new zones with existing zones
- Save merged list (append-only)

### Step 5: Add IsProcessed Flag to ProcessedFileCombo
- Add `IsProcessed` boolean property to `ProcessedFileCombo` class
- Set to `false` for new combos, `true` for existing combos

---

## 📋 CODE SECTIONS TO MODIFY

1. **`Services/RefreshService.cs`**:
   - Line 1190-1337: 3-point validation logic
   - Line 3175-3324: Merging and saving logic
   - Line 3329-3330: Remove UI state saving
   - Line 3708: Category-specific XML saving

2. **`Services/GlobalIndexService.cs`**:
   - `ProcessedFileCombo` class: Add `IsProcessed` property

3. **`Services/FilterManagementService.cs`**:
   - `SaveFilterToXmlFile`: May need to support merge mode

---

## ✅ VERIFICATION CHECKLIST

After fixes:
- [ ] Filter XML never overwrites existing zones
- [ ] New zones are appended to existing zones
- [ ] Zones only removed if "Adopt Document" enabled AND validation failed
- [ ] UI state not saved to Filter XML
- [ ] UI state only saved to Global XML ProcessedFileCombos
- [ ] IsProcessed flag added to ProcessedFileCombo
- [ ] FilterName set after intersection detection (before sleeve placement)

