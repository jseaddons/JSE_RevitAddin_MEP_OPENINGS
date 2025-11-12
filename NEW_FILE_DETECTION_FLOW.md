# Logic Flow: Adding New Linked File to Existing Filter

## Overview
This document explains the complete logic flow when a user adds a new linked file to an existing filter and triggers clash detection.

## Scenario
- **Existing Filter**: "Plumbing Check" (already saved with files: `MEP_Model_A.rvt`, `Structural_Model_A.rvt`)
- **User Action**: Adds new linked file `MEP_Model_B.rvt` to the filter
- **Goal**: Detect new clash zones from the newly added linked file

---

## Step-by-Step Flow

### STEP 1: User Adds New Linked File in UI
**Location**: UI file selection controls
- User selects `MEP_Model_B.rvt` in Reference Files dropdown
- UI state updates: `selectedReferenceFiles` now contains `["MEP_Model_A.rvt", "MEP_Model_B.rvt"]`

### STEP 2: User Saves Filter
**Location**: `Services/FilterManagementService.cs` → `SaveFilter()` or `SaveFilterToXmlFile()`
- Filter XML file is saved: `Plumbing Check.xml`
- **Saved to XML**:
  ```xml
  <OpeningFilter>
    <SelectedReferenceFiles>
      <string>MEP_Model_A.rvt</string>
      <string>MEP_Model_B.rvt</string>  <!-- ✅ NEW FILE SAVED -->
    </SelectedReferenceFiles>
    <SelectedHostFiles>...</SelectedHostFiles>
    <ClashZoneStorage>
      <!-- Existing clash zones from previous runs -->
    </ClashZoneStorage>
  </OpeningFilter>
  ```
- **Key Point**: The new file is now **persisted** in filter XML

### STEP 3: User Clicks "Process Clash Zones" (Refresh)
**Location**: `Services/RefreshService.cs` → `ExecuteRefreshInternal()`

#### 3.1: Load Filter from XML
**Lines 737-761**: `LoadFilterAuto()` loads filter XML
```csharp
var filter = _filterManagementService.LoadFilterAuto(filterName);
filtersToProcess.Add(filter); // Filter now contains SelectedReferenceFiles from XML
```
- **Result**: `filtersToProcess[0].SelectedReferenceFiles` = `["MEP_Model_A.rvt", "MEP_Model_B.rvt"]` (from XML)

#### 3.2: Load Existing Clash Zones
**Lines 854-864**: Load existing clash zones from Filter XML
```csharp
var existingClashZones = LoadExistingClashZonesFromFilterXml(selectedFilterItems, selectedMepCategories);
```
- **Result**: Existing clash zones loaded (from previous runs, no clashes from `MEP_Model_B.rvt` yet)

### STEP 4: UI State Comparison (NEW FILE DETECTION)
**Location**: `Services/RefreshService.cs` → Lines 1696-1766

#### 4.1: Get Saved Files from Filter XML
```csharp
// Get saved files from filtersToProcess (loaded from XML)
foreach (var filter in filtersToProcess)
{
    if (filter.SelectedReferenceFiles != null)
    {
        foreach (var f in filter.SelectedReferenceFiles.Select(norm))
            savedRefFiles.Add(f);  // ["mep model a", "mep model b"] ✅
    }
}
```
- **savedRefFiles**: `["mep model a", "mep model b"]` (normalized, from XML)

#### 4.2: Get Current UI Files
```csharp
// Get current UI files (from UI state)
var currentRefFiles = new HashSet<string>(
    (selectedReferenceFiles ?? new List<string>()).Select(f => norm(f)),
    StringComparer.OrdinalIgnoreCase);
```
- **currentRefFiles**: `["mep model a", "mep model b"]` (normalized, from UI)

#### 4.3: Compare and Detect New Files
```csharp
// Check if new files were added
bool hasNewRefFiles = savedRefFiles.Count == 0 && currentRefFiles.Count > 0 || 
                     currentRefFiles.Any(f => !savedRefFiles.Contains(f));
```
- **Logic**: 
  - If saved files empty AND current files exist → NEW FILTER (treat as new files)
  - If current files contains any file NOT in saved files → NEW FILE ADDED
- **Result**: `hasNewFilesAdded = true` ✅

#### 4.4: Log Detection Result
```csharp
if (hasNewFilesAdded)
{
    DebugLogger.Info($"[UI-STATE-CHECK] ✅ NEW FILES ADDED: {currentRefFiles.Count} ref files (saved: {savedRefFiles.Count})");
    DebugLogger.Info($"[UI-STATE-CHECK] Will process ALL intersections (skip resolved filtering)");
}
```

### STEP 5: Skip Resolved Filtering
**Location**: `Services/RefreshService.cs` → Lines 1768-1819

```csharp
if (!hasNewFilesAdded)
{
    // Build resolved index from Global XML
    // Skip intersections that are already resolved
}
else
{
    // ✅ SKIP resolved filtering - process ALL intersections
    DebugLogger.Info($"[REFRESH-GLOBAL-FILTER] Skipping resolved filtering - new files added");
}
```
- **Result**: `resolvedIntersectionPointsFromGlobalXml` remains **empty**
- **Why**: New files might have clashes that aren't in Global XML yet, so we need to process ALL intersections

### STEP 6: Intersection Detection
**Location**: `Services/RefreshService.cs` → Lines 1541-1549

```csharp
currentIntersections = intersectionService.FindIntersections(
    _document,
    view3D,
    selectedMepCategories,
    selectedReferenceFiles,  // ✅ Contains NEW file: ["MEP_Model_A.rvt", "MEP_Model_B.rvt"]
    selectedHostFiles,
    allowedHostTypesUI.ToList(),
    optimizationService);
```
- **Key Point**: `selectedReferenceFiles` comes from **UI state** (current selection), NOT from saved filter XML
- **Result**: `currentIntersections` contains:
  - Existing intersections from `MEP_Model_A.rvt` (already processed)
  - **NEW intersections from `MEP_Model_B.rvt`** ✅

### STEP 7: Filter Intersections
**Location**: `Services/RefreshService.cs` → Lines 1958-1998

```csharp
if (knownIntersectionsMap.Count > 0 || resolvedIntersectionPointsFromGlobalXml.Count > 0)
{
    foreach (var intersection in currentIntersections)
    {
        // Skip if resolved in Global XML (but resolvedIntersectionPointsFromGlobalXml is empty!)
        if (resolvedIntersectionPointsFromGlobalXml.ContainsKey(key))
        {
            skippedResolvedCount++;
            continue; // Skip
        }
        
        // Skip if already known (from existing clash zones)
        if (knownIntersectionsMap.ContainsKey(key))
        {
            skippedCount++;
            continue; // Skip
        }
        
        // ✅ NEW intersection - keep it
        filteredIntersections.Add(intersection);
    }
}
```
- **Result**: 
  - Intersections from `MEP_Model_A.rvt` → Filtered out (already in `knownIntersectionsMap`)
  - **Intersections from `MEP_Model_B.rvt`** → **Kept** (not in known map) ✅

### STEP 8: Create New Clash Zones
**Location**: `Services/RefreshService.cs` → Lines 2113-2142

```csharp
newClashZones = _clashZoneService.DetectNewClashZones(
    filteredIntersections,  // ✅ Contains NEW intersections from MEP_Model_B.rvt
    _document,
    clearanceSettings,
    selectedMepCategories);
```
- **Result**: New clash zones created for intersections from `MEP_Model_B.rvt`

### STEP 9: Save to Filter XML
**Location**: `Services/RefreshService.cs` → Lines 2846-2847, 2977-2984

```csharp
// Update filter with new clash zones
targetFilter.ClashZoneStorage.ClashZones = allClashZones; // Existing + New
targetFilter.SelectedReferenceFiles = selectedReferenceFiles; // ✅ Includes new file
targetFilter.SelectedHostFiles = selectedHostFiles;

// Save to XML
_filterManagementService.SaveFilterToXmlFile(targetFilter, mainFilePath);
```
- **Result**: Filter XML updated with:
  - New clash zones from `MEP_Model_B.rvt`
  - Updated `SelectedReferenceFiles` including new file

### STEP 10: Save Category-Specific XML Files
**Location**: `Services/RefreshService.cs` → Lines 2986-3214

```csharp
foreach (var category in selectedCategories)
{
    var categoryClashZones = FilterClashZonesByCategory(...);
    
    var categoryFilter = new OpeningFilter
    {
        Name = "plumbing check_pipes",  // Category-specific filename
        SelectedReferenceFiles = targetFilter.SelectedReferenceFiles, // ✅ Includes new file
        ClashZoneStorage = new ClashZoneStorage
        {
            ClashZones = categoryClashZones  // Existing + New clash zones
        }
    };
    
    _filterManagementService.SaveFilterToXmlFile(categoryFilter, categoryFilePath);
}
```
- **Result**: Category-specific XML files updated with new clash zones

---

## Key Points

### ✅ Why New File Detection Works

1. **Filter XML Stores Files**: When filter is saved, `SelectedReferenceFiles` is saved to XML
2. **Refresh Loads Filter XML**: Gets saved files from XML
3. **UI State Comparison**: Compares saved files (from XML) with current UI files
4. **Detection Logic**: 
   - If current UI has files NOT in saved XML → NEW FILE ADDED
   - If saved files empty AND current files exist → NEW FILTER
5. **Skip Resolved Filtering**: When new files detected, skip Global XML resolved filtering
6. **Process All Intersections**: Intersection detection uses **current UI files** (includes new file)
7. **Filter Known Intersections**: Only filter out intersections already in `knownIntersectionsMap` (existing clash zones)
8. **Create New Clash Zones**: New intersections from new file → New clash zones created

### ⚠️ Important Notes

- **UI State is Source of Truth**: Intersection detection uses `selectedReferenceFiles` from **UI**, not from saved filter XML
- **Filter XML is Updated**: After refresh, filter XML is updated with new clash zones AND new files
- **Next Refresh**: Next time refresh runs, saved files will match current UI files → No new files detected → Normal filtering resumes

---

## Flow Diagram

```
User Adds New File → Save Filter
         ↓
Filter XML Updated: SelectedReferenceFiles = [Old Files + New File]
         ↓
User Clicks "Process Clash Zones"
         ↓
Load Filter XML → Get savedRefFiles = [Old Files + New File]
         ↓
Compare with UI → currentRefFiles = [Old Files + New File]
         ↓
hasNewFilesAdded = true ✅ (if new file not in savedRefFiles)
         ↓
Skip Resolved Filtering (resolvedIntersectionPointsFromGlobalXml = empty)
         ↓
Run Intersection Detection → Uses UI Files (includes new file)
         ↓
currentIntersections = [All intersections from all files]
         ↓
Filter Intersections:
  - Skip if in knownIntersectionsMap (existing clash zones)
  - Keep NEW intersections (from new file) ✅
         ↓
Create New Clash Zones → newClashZones = [Clash zones from new file]
         ↓
Save to Filter XML:
  - Update ClashZoneStorage with new clash zones
  - Update SelectedReferenceFiles (already has new file)
         ↓
Save Category-Specific XML Files → Updated with new clash zones
```

---

## Code References

### UI State Comparison
```1696:1766:Services/RefreshService.cs
// ✅ CRITICAL FIX: Check if UI state has changed (new files added)
// If new files added → Don't filter by resolved status (process all intersections for new files)
// Then filter out resolved zones from Global XML for existing zones
// ✅ FIX: Check UI state even if no existing clash zones (for new filters)
bool hasNewFilesAdded = false;
if (filtersToProcess != null && filtersToProcess.Count > 0)
{
    // Compare current UI files with saved files from filters
    Func<string, string> norm = s =>
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        var trimmed = s;
        var idxParen = trimmed.IndexOf('(');
        if (idxParen >= 0) trimmed = trimmed.Substring(0, idxParen);
        trimmed = System.IO.Path.GetFileNameWithoutExtension(trimmed);
        trimmed = trimmed.ToLowerInvariant().Replace("_detached", "");
        trimmed = trimmed.Replace('_', ' ').Replace('-', ' ');
        trimmed = System.Text.RegularExpressions.Regex.Replace(trimmed, "\\s+", " ");
        return trimmed.Trim();
    };
    
    // Get saved files from filter XML (if available)
    var savedRefFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var savedHostFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    
    // Try to get saved files from filtersToProcess (loaded earlier)
    foreach (var filter in filtersToProcess)
    {
        if (filter.SelectedReferenceFiles != null)
        {
            foreach (var f in filter.SelectedReferenceFiles.Select(norm))
                savedRefFiles.Add(f);
        }
        if (filter.SelectedHostFiles != null)
        {
            foreach (var f in filter.SelectedHostFiles.Select(norm))
                savedHostFiles.Add(f);
        }
    }
    
    // Get current UI files
    var currentRefFiles = new HashSet<string>(
        (selectedReferenceFiles ?? new List<string>()).Select(f => norm(f)),
        StringComparer.OrdinalIgnoreCase);
    var currentHostFiles = new HashSet<string>(
        (selectedHostFiles ?? new List<string>()).Select(f => norm(f)),
        StringComparer.OrdinalIgnoreCase);
    
    // Check if new files were added
    // ✅ FIX: For new filters, savedRefFiles/savedHostFiles might be empty, so compare with current UI
    bool hasNewRefFiles = savedRefFiles.Count == 0 && currentRefFiles.Count > 0 || 
                         currentRefFiles.Any(f => !savedRefFiles.Contains(f));
    bool hasNewHostFiles = savedHostFiles.Count == 0 && currentHostFiles.Count > 0 || 
                          currentHostFiles.Any(f => !savedHostFiles.Contains(f));
    hasNewFilesAdded = hasNewRefFiles || hasNewHostFiles;
    
    if (hasNewFilesAdded)
    {
        if (!DeploymentConfiguration.DeploymentMode)
        {
            DebugLogger.Info($"[UI-STATE-CHECK] ✅ NEW FILES ADDED: {currentRefFiles.Count} ref files (saved: {savedRefFiles.Count}), {currentHostFiles.Count} host files (saved: {savedHostFiles.Count})");
            DebugLogger.Info($"[UI-STATE-CHECK] Will process ALL intersections (skip resolved filtering) to detect new clashes from new files");
        }
        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [UI-STATE-CHECK] ✅ NEW FILES ADDED: Processing all intersections (skip resolved filtering)\n");
    }
    else
    {
        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[UI-STATE-CHECK] No new files added - will filter resolved zones from Global XML");
    }
}
```

### Skip Resolved Filtering When New Files Added
```1768:1819:Services/RefreshService.cs
try
{
    // ✅ ONLY build resolved index if NO new files added
    // If new files added → Process all intersections (don't filter by resolved status)
    if (!hasNewFilesAdded)
    {
        // ✅ Build index from ALL Global XML entries with resolved flags (not just ones matching Filter XML)
        // This handles case where Filter B detects clash zone that Filter A already placed sleeves for
        foreach (var category in selectedMepCategories ?? new List<string>())
        {
            try
            {
                var globalIndex = GlobalIndexService.BuildMepHostPointIndex(_document, category);
                
                // ✅ SIMPLE LOGIC: Only include entries with resolved flags = true
                // If flags are false, entry won't be in resolvedIntersectionPointsFromGlobalXml, so clash zone will be processed
                // ResetFlagsForDeletedSleeves already checks sleeve existence and sets flags to false when sleeves are deleted
                foreach (var kvp in globalIndex)
                {
                    var entry = kvp.Value;
                    
                    // ✅ CRITICAL LOGIC: Skip ONLY if BOTH flags are true
                    // Process if ANY flag is false (IsResolved == false OR IsClusterResolved == false)
                    if (entry.IsResolved && entry.IsClusterResolved)
                    {
                        resolvedIntersectionPointsFromGlobalXml[kvp.Key] = entry; // Add to skip list
                    }
                    // If any flag is false → don't add to skip list → process clash zone
                }
            }
            catch (Exception categoryEx)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Warning($"[REFRESH-GLOBAL-FILTER] Error building index for category '{category}': {categoryEx.Message}");
                }
            }
        }
        
        if (resolvedIntersectionPointsFromGlobalXml.Count > 0 && !DeploymentConfiguration.DeploymentMode)
        {
            DebugLogger.Info($"[REFRESH-GLOBAL-FILTER] Built resolved intersections index: {resolvedIntersectionPointsFromGlobalXml.Count} (MEP+Host+Point) entries already resolved in Global XML");
            SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-FILTER] Built resolved index: {resolvedIntersectionPointsFromGlobalXml.Count} entries\n");
        }
    }
    else
    {
        if (!DeploymentConfiguration.DeploymentMode)
            DebugLogger.Info($"[REFRESH-GLOBAL-FILTER] Skipping resolved filtering - new files added, will process all intersections");
        SafeFileLogger.SafeAppendText(refreshLogName, $"[{DateTime.Now}] [REFRESH-GLOBAL-FILTER] Skipping resolved filtering - new files added\n");
    }
}
```

### Intersection Detection Uses UI Files
```1541:1549:Services/RefreshService.cs
currentIntersections = intersectionService.FindIntersections(
    _document,
    view3D,
    selectedMepCategories,
    selectedReferenceFiles,  // ✅ From UI state (includes new file)
    selectedHostFiles,
    allowedHostTypesUI.ToList(),
    optimizationService);
```

### Filter Intersections (Skip Known Only)
```1964:1994:Services/RefreshService.cs
foreach (var intersection in currentIntersections)
{
    var (mepElement, structuralElement, bbox, point) = intersection;
    int mepId = mepElement.Id.IntegerValue;
    int hostId = structuralElement.Id.IntegerValue;
    
    // ✅ CRITICAL FIX: Check if intersection point matches resolved clash zone in Global XML
    // This handles case where Filter B detects clash zone that Filter A already placed sleeves for
    // Uses O(1) Dictionary lookup by MEP+Host+Point (not GUID) - enables cross-filter matching
    double tolerance = 0.1;
    string pointKey = $"{Math.Round(point.X / tolerance) * tolerance:F1}," +
                     $"{Math.Round(point.Y / tolerance) * tolerance:F1}," +
                     $"{Math.Round(point.Z / tolerance) * tolerance:F1}";
    var key = (mepId, hostId, pointKey);
    
    if (resolvedIntersectionPointsFromGlobalXml.ContainsKey(key))
    {
        // Intersection point matches resolved clash zone in Global XML - skip creating new clash zone
        skippedResolvedCount++;
        continue; // Skip this intersection - already resolved in Global XML with matching point
    }
    
    // Then check known intersections map (performance optimization)
    if (knownIntersectionsMap.ContainsKey(key))
    {
        skippedCount++;
        continue; // Skip this intersection - already known (performance optimization)
    }
    
    filteredIntersections.Add(intersection); // NEW intersection - keep it for clash zone creation
}
```

---

## Summary

**The Trigger Mechanism:**
1. Filter XML is saved with new file → `SelectedReferenceFiles` updated in XML
2. Refresh loads filter XML → Gets saved files
3. Compares saved files (from XML) with current UI files
4. If mismatch detected → `hasNewFilesAdded = true`
5. Skip resolved filtering → Process ALL intersections
6. Intersection detection uses **current UI files** (includes new file)
7. New intersections from new file → New clash zones created
8. Filter XML updated with new clash zones

**Key Insight**: The comparison happens **BEFORE** intersection detection, so when new files are detected, the system knows to process ALL intersections instead of filtering by resolved status.

