# Bounding Box Save Issue - Debug Checklist

## Log Files Location
**EXACT Path**: `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\`

**EXACT Log File Names**:
- **`cabletraysleeveplacer.log`** - Contains ALL DebugLogger entries:
  - `[BOUNDING_BOX_AFTER_PLACEMENT]` (also in cluster_debug.log)
  - `[BOUNDING_BOX_BEFORE_REGEN]`
  - `[BOUNDING_BOX_AFTER_REGEN]`
  - `[BOUNDING_BOX_BEFORE_SET]`
  - `[BOUNDING_BOX_AFTER_SET]`
  - `[BOUNDING_BOX_BEFORE_XML_SAVE]`
  - `[BOUNDING_BOX_AFTER_XML_SAVE]`
  - `[UpdateSleeveCoordinatesInXml] CALLED`
  - `[DIRECT-ID-MATCH]`
  - `[UPDATE_COORD_DEBUG]`
  - ✅ **NOW OVERWRITES** (cleared at each session start)
  
- **`cluster_debug.log`** - Contains:
  - `[BOUNDING_BOX_AFTER_PLACEMENT]` (duplicate of cabletraysleeveplacer.log)
  - ✅ **APPENDS** (keeps all historical entries)

---

## STEP 1: Check After Placement (Before Regeneration)

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log` OR `cluster_debug.log`:
- [ ] Search for `[BOUNDING_BOX_AFTER_PLACEMENT]`
- [ ] **Check**: Do these entries exist?
- [ ] **Check**: Are coordinates non-zero? (should be actual values like `MinX=1234.567890`)
- [ ] **Check**: Does `SleeveInstanceId` match the sleeve IDs you're looking for?
- [ ] **Note**: Format: `Sleeve {ID}: Min=(...), Max=(...)` OR `MinX=..., MinY=..., MinZ=..., MaxX=..., MaxY=..., MaxZ=...`

**What to look for:**
- ✅ If coordinates are non-zero → bounding box was captured during placement
- ❌ If coordinates are zero → bounding box was NOT captured during placement

---

## STEP 2: Check Before Regeneration

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `[BOUNDING_BOX_BEFORE_REGEN]`
- [ ] **Check**: Do these entries exist?
- [ ] **Check**: Are bounding boxes NULL or non-zero?
- [ ] **Check**: Count of sleeves found matches expected?
- [ ] **Check**: How many `[BOUNDING_BOX_BEFORE_REGEN]` entries total?

**What to look for:**
- ✅ If bboxes are non-zero → Revit has bounding boxes before regeneration
- ❌ If bboxes are NULL or zero → Revit doesn't have bounding boxes yet (needs regeneration)

---

## STEP 3: Check After Regeneration

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `[BOUNDING_BOX_AFTER_REGEN]`
- [ ] **Check**: Do these entries exist?
- [ ] **Check**: Are bounding boxes non-zero after regeneration?
- [ ] **Check**: Do values match the `[BOUNDING_BOX_AFTER_PLACEMENT]` values?
- [ ] **Check**: How many `[BOUNDING_BOX_AFTER_REGEN]` entries total?

**What to look for:**
- ✅ If bboxes are non-zero → Regeneration worked, bounding boxes available
- ❌ If bboxes are still NULL → Regeneration didn't help

---

## STEP 4: Check UpdateSleeveCoordinates Call

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `[UpdateSleeveCoordinatesInXml] CALLED`
- [ ] **Check**: Was this method called?
- [ ] **Check**: `Loaded {X} clash zones from XML` → How many loaded?
- [ ] **Check**: `{X} out of {Y} clash zones have SleeveInstanceId > 0` → Are sleeve IDs present in XML?
- [ ] **Check**: `Looking for sleeves with IDs: [...]` → Do these match your new sleeve IDs?

**What to look for:**
- ✅ If sleeve IDs are found → XML has SleeveInstanceId saved
- ❌ If sleeve IDs are missing → SleeveInstanceId not saved to XML during placement

---

## STEP 5: Check Matching in UpdateSleeveCoordinates

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `[DIRECT-ID-MATCH]`
- [ ] **Check**: `Found individual sleeve {ID} by ID` → Sleeves found by ID?
- [ ] **Check**: `Sleeve {ID} not found in document` → Any sleeves not found?
- [ ] Search for `[UPDATE_COORD_DEBUG]`
- [ ] **Check**: `Matching IDs: {X} out of {Y}` → How many matched?

**What to look for:**
- ✅ If all sleeves matched → ID lookup working
- ❌ If sleeves not found → IDs don't exist in Revit or don't match

---

## STEP 6: Check Before SetSleeveBoundingBox

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `[BOUNDING_BOX_BEFORE_SET]`
- [ ] **Check**: `ClashZone has: MinX=...` → What values does ClashZone have BEFORE update?
- [ ] **Check**: `Revit bbox from sleeve {ID}` → What values does Revit have?

**What to look for:**
- ✅ If ClashZone has zeros but Revit has values → Will be updated correctly
- ❌ If ClashZone has non-zero but Revit has zeros → Problem with Revit read
- ❌ If both have zeros → Problem with Revit bounding box availability

---

## STEP 7: Check After SetSleeveBoundingBox

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `[BOUNDING_BOX_AFTER_SET]`
- [ ] **Check**: `ClashZone now has: MinX=...` → What values does ClashZone have AFTER update?
- [ ] **Check**: Do values match Revit values from BEFORE_SET?

**What to look for:**
- ✅ If values match Revit → SetSleeveBoundingBox worked correctly
- ❌ If values still zeros → SetSleeveBoundingBox didn't work

---

## STEP 8: Check Before XML Save

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `[BOUNDING_BOX_BEFORE_XML_SAVE]`
- [ ] **Check**: What values are being saved?
- [ ] **Check**: Are these the same as `[BOUNDING_BOX_AFTER_SET]`?

**What to look for:**
- ✅ If values match AFTER_SET → Data preserved correctly
- ❌ If values reverted to zeros → Something overwrote them

---

## STEP 9: Check After XML Save

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `[BOUNDING_BOX_AFTER_XML_SAVE]`
- [ ] **Check**: `XML now has - MinX=...` → What values are in XML after save?
- [ ] **Check**: Do values match `[BOUNDING_BOX_BEFORE_XML_SAVE]`?

**What to look for:**
- ✅ If values match → XML save worked correctly
- ❌ If values are zeros or different → XML save didn't work or file was overwritten

---

## STEP 10: Check UpdateSleeveCoordinatesInXml Completion

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `After UpdateSleeveCoordinates: {X} out of {Y} clash zones with SleeveInstanceId now have bounding boxes`
- [ ] **Check**: How many got bounding boxes?
- [ ] **Check**: Does this match number of sleeves placed?

**What to look for:**
- ✅ If all sleeves got bboxes → Process succeeded
- ❌ If zero sleeves got bboxes → Process failed

---

## STEP 11: Check XML File Directly

### File Location:
- `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Projects\<ProjectName>\Filters\Plumbing_pipes.xml` (or relevant filter XML)

### For each sleeve ID you're checking:
- [ ] Open XML file
- [ ] Search for `<SleeveInstanceId>{YOUR_SLEEVE_ID}</SleeveInstanceId>`
- [ ] **Check**: Does `SleeveInstanceId` match your sleeve ID?
- [ ] **Check**: `<SleeveBoundingBoxMinX>` → Is it zero or non-zero?
- [ ] **Check**: `<SleeveBoundingBoxMinY>` → Is it zero or non-zero?
- [ ] **Check**: `<SleeveBoundingBoxMinZ>` → Is it zero or non-zero?
- [ ] **Check**: `<SleeveBoundingBoxMaxX>` → Is it zero or non-zero?
- [ ] **Check**: `<SleeveBoundingBoxMaxY>` → Is it zero or non-zero?
- [ ] **Check**: `<SleeveBoundingBoxMaxZ>` → Is it zero or non-zero?

**What to look for:**
- ✅ If all 6 coordinates are non-zero → XML save worked
- ❌ If any coordinate is zero → XML save didn't work

---

## STEP 12: Check Timing Sequence

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
Look for this sequence (should appear in order):

1. [ ] `[BOUNDING_BOX_AFTER_PLACEMENT]` - First (after placement)
2. [ ] `[BOUNDING_BOX_BEFORE_REGEN]` - Second (before regeneration)
3. [ ] `[BOUNDING_BOX_AFTER_REGEN]` - Third (after regeneration)
4. [ ] `[UpdateSleeveCoordinatesInXml] CALLED` - Fourth
5. [ ] `[BOUNDING_BOX_BEFORE_SET]` - Fifth (during update)
6. [ ] `[BOUNDING_BOX_AFTER_SET]` - Sixth (after update)
7. [ ] `[BOUNDING_BOX_BEFORE_XML_SAVE]` - Seventh (before save)
8. [ ] `[BOUNDING_BOX_AFTER_XML_SAVE]` - Eighth (after save)

**What to look for:**
- ✅ If all steps appear in order → Flow is correct
- ❌ If steps are missing or out of order → Flow is broken

---

## STEP 13: Check for Errors/Warnings

### In `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs\cabletraysleeveplacer.log`:
- [ ] Search for `ERROR`
- [ ] Search for `WARNING`
- [ ] Search for `Exception`
- [ ] Search for `not found`
- [ ] Search for `NULL`

**What to look for:**
- Any errors that might prevent bounding box save
- Warnings about missing sleeves or null bounding boxes

---

## Summary Questions to Answer

1. **After placement**: Are bounding boxes logged? (Y/N) Values: _______
2. **After regeneration**: Are bounding boxes available? (Y/N) Values: _______
3. **UpdateSleeveCoordinatesInXml**: Was it called? (Y/N) How many sleeves processed? _______
4. **Matching**: How many sleeves matched by ID? _______ out of _______
5. **Before SetSleeveBoundingBox**: ClashZone has _______, Revit has _______
6. **After SetSleeveBoundingBox**: ClashZone has _______
7. **Before XML save**: ClashZone has _______
8. **After XML save**: XML has _______
9. **Final XML file**: All coordinates non-zero? (Y/N)
10. **Where did it break?** (At which step above did values become zeros?)

---

## Expected Flow Summary

```
Place Sleeve → BBOX in memory (AFTER_PLACEMENT) → cabletraysleeveplacer.log + cluster_debug.log
    ↓
Regenerate → BBOX available in Revit (BEFORE_REGEN → AFTER_REGEN) → cabletraysleeveplacer.log
    ↓
UpdateSleeveCoordinatesInXml → Read from Revit (BEFORE_SET → AFTER_SET) → cabletraysleeveplacer.log
    ↓
SaveClashZonesToXml → Write to XML (BEFORE_XML_SAVE → AFTER_XML_SAVE) → cabletraysleeveplacer.log
    ↓
XML file has correct values
```

**The break point is where values go from non-zero to zero.**
