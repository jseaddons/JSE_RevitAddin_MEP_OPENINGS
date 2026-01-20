# Complete Clustering Implementation Guide

## Overview
This document provides a complete guide to the sleeve clustering system, including data flow, file management, cache operations, and critical methods.

## 1. Data Flow Architecture

### 1.1 Two-Phase Process
```
Phase 1: Refresh Button
├── Load clash zones from XML
├── Detect intersections
├── Place individual sleeves (UniversalSleevePlacerService)
├── Update coordinates (SleeveCoordinateService)
└── Save sleeve data to _CLUSTER.xml (SleeveCoordinateService)

Phase 2: OK Button  
├── Load sleeve data from _CLUSTER.xml (UniversalClusterService)
├── Build cache with MepElementId keys
├── Perform proximity clustering
├── Place cluster sleeves
└── Delete individual sleeves
```

### 1.2 File Types and Their Purpose
- **Regular XML** (`Ventilation_ducts.xml`): Contains clash zone data with intersection points
- **_CLUSTER.xml** (`Ventilation_Duct Accessories_CLUSTER.xml`): Contains actual sleeve placement coordinates
- **CONDITIONS.xml**: Contains filter conditions

## 2. Critical File Locations

### 2.1 XML File Storage
```
C:\Users\{username}\AppData\Roaming\JSE_MEP_Openings\Projects\Default\Filters\
├── Ventilation.xml (main filter)
├── Ventilation_ducts.xml (clash zones)
├── Ventilation_duct_accessories.xml (clash zones)
├── Ventilation_Duct Accessories_CLUSTER.xml (sleeve coordinates)
├── Ventilation_ducts_CONDITIONS.xml
└── Ventilation_ductaccessories_CONDITIONS.xml
```

### 2.2 Log Files
```
C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\
├── cluster_debug.log (clustering process)
├── placement_debug.log (sleeve placement)
├── coordinate_update.log (coordinate updates)
└── all_sleeve_coordinates.log (sleeve coordinates)
```

## 3. Phase 1: Refresh Button Process

### 3.1 Sleeve Placement Service
**File**: `Services/UniversalSleevePlacerService.cs`

**Key Method**: `PlaceSleevesForClashZones()`
```csharp
public void PlaceSleevesForClashZones(List<ClashZone> clashZones, string filterName, string categoryName)
{
    // 1. Place individual sleeves
    // 2. Get actual coordinates from placed sleeves
    // 3. Save to _CLUSTER.xml
}
```

**Critical Steps**:
1. Place individual sleeves using Revit API
2. Get bounding box coordinates from placed sleeve instances
3. Create `SleeveData` objects with 4 corner coordinates
4. Save to `*_{category}_CLUSTER.xml` file

### 3.2 SleeveData Structure
```csharp
public class SleeveData
{
    public int SleeveInstanceId { get; set; }
    public XYZ Corner1 { get; set; }  // Min corner
    public XYZ Corner2 { get; set; }  // Min corner (different axis)
    public XYZ Corner3 { get; set; }  // Max corner
    public XYZ Corner4 { get; set; }  // Max corner (different axis)
    public double Width { get; set; }
    public double Height { get; set; }
    public double Depth { get; set; }
    public string HostType { get; set; }
    public string Orientation { get; set; }
    public string Category { get; set; }
    public DateTime CreatedAt { get; set; }
}
```

### 3.3 Coordinate Update Service
**File**: `Services/SleeveCoordinateService.cs`

**Key Method**: `RegenerateClusterXmlAfterPlacement()`
```csharp
public void RegenerateClusterXmlAfterPlacement(string filterName = null)
{
    // 1. Wait for Revit to update the model (500ms)
    System.Threading.Thread.Sleep(500);
    
    // 2. Collect ALL sleeves with actual Revit coordinates
    var allSleeves = new FilteredElementCollector(_doc)
        .OfClass(typeof(FamilyInstance))
        .Cast<FamilyInstance>()
        .Where(s => s.Symbol.FamilyName.Contains("Opening"))
        .ToList();
    
    // 3. Group sleeves by category for separate XML files
    var groupedSleeves = allSleeves.GroupBy(s => GetSleeveCategory(s)).ToList();
    
    // 4. Create SleeveDataList with actual Revit coordinates
    foreach (var group in groupedSleeves)
    {
        var sleeveDataList = new List<SleeveData>();
        foreach (var sleeve in group)
        {
            var bbox = sleeve.get_BoundingBox(null);
            if (bbox != null)
            {
                var sleeveData = new SleeveData
                {
                    SleeveInstanceId = sleeve.Id.IntegerValue,
                    Corner1 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                    Corner2 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                    Corner3 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
                    Corner4 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
                    Width = UnitUtils.ConvertFromInternalUnits(bbox.Max.X - bbox.Min.X, UnitTypeId.Meters),
                    Height = UnitUtils.ConvertFromInternalUnits(bbox.Max.Y - bbox.Min.Y, UnitTypeId.Meters),
                    Depth = UnitUtils.ConvertFromInternalUnits(bbox.Max.Z - bbox.Min.Z, UnitTypeId.Meters),
                    HostType = GetHostType(sleeve),
                    Orientation = GetSleeveOrientation(sleeve),
                    Category = group.Key,
                    CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                };
                sleeveDataList.Add(sleeveData);
            }
        }
        
        // 5. Save to _CLUSTER.xml with actual coordinates
        SaveSleeveDataToClusterXml(sleeveDataList, group.Key, filterName);
    }
}
```

**Critical Features**:
- **Timing Fix**: Waits 500ms for Revit to update model before reading coordinates
- **Direct Revit Reading**: Reads actual coordinates from placed sleeve instances
- **Category Detection**: Determines category from MEP element or family name
- **Dual File Creation**: Creates both singular and plural category files for compatibility

### 3.4 Saving to _CLUSTER.xml
**File**: `Services/SleeveCoordinateService.cs`

**Key Method**: `SaveSleeveDataToClusterXml()`
```csharp
private void SaveSleeveDataToClusterXml(List<SleeveData> sleeveDataList, string category, string filterName = null)
{
    // Create filename that matches clustering service expectations
    var actualFilterName = filterName ?? "Ventilation";
    var fileName = $"{actualFilterName}_{category.Replace(" ", "_").ToLower()}_CLUSTER.xml";
    
    // Also create the plural version that clustering service expects
    var pluralCategory = category switch
    {
        "Pipe Accessories" => "Pipes",
        "Duct Accessories" => "Ducts", 
        "Cable Tray Accessories" => "Cable Trays",
        _ => category
    };
    var pluralFileName = $"{actualFilterName}_{pluralCategory.Replace(" ", "_").ToLower()}_CLUSTER.xml";
    
    // Create XML document with SleeveData
    var xmlDoc = new System.Xml.XmlDocument();
    var root = xmlDoc.CreateElement("SleeveDataList");
    xmlDoc.AppendChild(root);
    
    foreach (var sleeveData in sleeveDataList)
    {
        var sleeveElement = xmlDoc.CreateElement("SleeveData");
        AddXmlElement(xmlDoc, sleeveElement, "SleeveInstanceId", sleeveData.SleeveInstanceId.ToString());
        AddXmlElement(xmlDoc, sleeveElement, "Corner1X", sleeveData.Corner1.X.ToString("F6"));
        // ... add all corner coordinates, dimensions, and metadata
        root.AppendChild(sleeveElement);
    }
    
    // Save both singular and plural versions
    xmlDoc.Save(filePath);
    xmlDoc.Save(pluralFilePath);
}
```

## 4. Phase 2: OK Button Process

### 4.1 Clustering Service Entry Point
**File**: `Services/UniversalClusterService.cs`

**Key Method**: `ClusterSleeves()`
```csharp
public (int placedCount, int deletedCount) ClusterSleeves(Document doc, string targetCategory, UIDocument uiDoc = null, string xmlFilePath = null)
{
    // 1. Load cache from _CLUSTER.xml
    // 2. Filter sleeves by category
    // 3. Group by host type, system type, orientation
    // 4. Perform proximity clustering
    // 5. Place cluster sleeves
    // 6. Delete individual sleeves
}
```

### 4.2 Cache Loading Process
**Key Method**: `LoadClashZoneCache()`

**Critical Implementation**:
```csharp
private void LoadClashZoneCache(string xmlFilePath, string targetCategory = null, Document doc = null)
{
    _clashZoneCache.Clear();
    
    // 1. Look for _CLUSTER.xml files
    var pattern1 = $"*_{targetCategory}_CLUSTER.xml";
    var pattern2 = $"*_{targetCategory.Replace("_", " ")}_CLUSTER.xml";
    var clusterFiles = Directory.GetFiles(filtersDirectory, pattern1)
        .Concat(Directory.GetFiles(filtersDirectory, pattern2))
        .Distinct()
        .ToArray();
    
    if (clusterFiles.Length > 0)
    {
        // 2. Load SleeveData from _CLUSTER.xml
        var sleeveDataService = new SleeveDataService(doc);
        var sleeveDataList = sleeveDataService.LoadSleeveDataFromClusterXml("Ventilation", targetCategory);
        
        // 3. Convert SleeveData to ClashZone
        foreach (var sleeveData in sleeveDataList)
        {
            // Get MepElementId from sleeve instance
            var sleeveElement = doc.GetElement(new ElementId(sleeveData.SleeveInstanceId));
            var mepElementIdParam = sleeveElement.LookupParameter("MEP_ElementId");
            long mepElementId = mepElementIdParam.AsInteger();
            
            // Calculate bounding box from 4 corners
            var minX = Math.Min(Math.Min(sleeveData.Corner1.X, sleeveData.Corner2.X), Math.Min(sleeveData.Corner3.X, sleeveData.Corner4.X));
            var maxX = Math.Max(Math.Max(sleeveData.Corner1.X, sleeveData.Corner2.X), Math.Max(sleeveData.Corner3.X, sleeveData.Corner4.X));
            // ... similar for Y and Z
            
            var clashZone = new Models.ClashZone
            {
                SleeveInstanceId = sleeveData.SleeveInstanceId,
                SleeveBoundingBoxMinX = minX,
                SleeveBoundingBoxMinY = minY,
                SleeveBoundingBoxMinZ = minZ,
                SleeveBoundingBoxMaxX = maxX,
                SleeveBoundingBoxMaxY = maxY,
                SleeveBoundingBoxMaxZ = maxZ,
                MepElementCategory = sleeveData.Category,
                StructuralElementType = sleeveData.HostType,
                MepElementOrientationDirection = sleeveData.Orientation, // ✅ CORRECT FIELD NAME
                SleeveWidth = sleeveData.Width,
                SleeveHeight = sleeveData.Height,
                SleeveDiameter = sleeveData.Depth,
                IsCurrentClash = true
            };
            
            // CRITICAL: Use MepElementId as cache key
            _clashZoneCache[mepElementId] = clashZone;
        }
    }
}
```

### 4.3 Cache Key Management
**CRITICAL ISSUE**: Cache key mismatch

**Problem**:
- Cache loaded with `SleeveInstanceId` as key
- Cache lookup uses `MepElementId` as key
- Result: Cache has data but lookup fails

**Solution**:
```csharp
// WRONG: Using SleeveInstanceId as key
_clashZoneCache[sleeveData.SleeveInstanceId] = clashZone;

// CORRECT: Using MepElementId as key
_clashZoneCache[mepElementId] = clashZone;
```

### 4.4 Sleeve Filtering Process
**Key Method**: Inside `ClusterSleeves()`

```csharp
// Get all sleeves in document
var allSleeves = new FilteredElementCollector(doc)
    .OfClass(typeof(FamilyInstance))
    .Cast<FamilyInstance>()
    .Where(f => f.Symbol.FamilyName.Contains("Opening"))
    .ToList();

// Filter sleeves that are in cache
var cacheSleeves = allSleeves.Where(s =>
{
    var mepElementIdParam = s.LookupParameter("MEP_ElementId");
    if (mepElementIdParam != null)
    {
        long mepElementId = mepElementIdParam.AsInteger();
        if (_clashZoneCache != null && _clashZoneCache.ContainsKey(mepElementId))
        {
            var clashZone = _clashZoneCache[mepElementId];
            return clashZone.IsCurrentClash; // Only process current clashes
        }
    }
    return false;
}).ToList();
```

### 4.5 Proximity Clustering Algorithm
**Key Method**: `CheckBoundingBoxOverlapWithOrientation()`

**Implementation**:
```csharp
private bool CheckBoundingBoxOverlapWithOrientation(ClashZone rect1, ClashZone rect2, double toleranceDist, string orientation)
{
    // Get 2D coordinates based on host type and orientation
    double x1Min, y1Min, x1Max, y1Max;
    double x2Min, y2Min, x2Max, y2Max;
    
    if (orientation == "Y") // Y-oriented walls
    {
        // Use Y,Z coordinates
        x1Min = rect1.SleeveBoundingBoxMinY;
        y1Min = rect1.SleeveBoundingBoxMinZ;
        x1Max = rect1.SleeveBoundingBoxMaxY;
        y1Max = rect1.SleeveBoundingBoxMaxZ;
        
        x2Min = rect2.SleeveBoundingBoxMinY;
        y2Min = rect2.SleeveBoundingBoxMinZ;
        x2Max = rect2.SleeveBoundingBoxMaxY;
        y2Max = rect2.SleeveBoundingBoxMaxZ;
    }
    else if (orientation == "X") // X-oriented walls
    {
        // Use X,Z coordinates
        x1Min = rect1.SleeveBoundingBoxMinX;
        y1Min = rect1.SleeveBoundingBoxMinZ;
        x1Max = rect1.SleeveBoundingBoxMaxX;
        y1Max = rect1.SleeveBoundingBoxMaxZ;
        
        x2Min = rect2.SleeveBoundingBoxMinX;
        y2Min = rect2.SleeveBoundingBoxMinZ;
        x2Max = rect2.SleeveBoundingBoxMaxX;
        y2Max = rect2.SleeveBoundingBoxMaxZ;
    }
    else // Floors
    {
        // Use X,Y coordinates
        x1Min = rect1.SleeveBoundingBoxMinX;
        y1Min = rect1.SleeveBoundingBoxMinY;
        x1Max = rect1.SleeveBoundingBoxMaxX;
        y1Max = rect1.SleeveBoundingBoxMaxY;
        
        x2Min = rect2.SleeveBoundingBoxMinX;
        y2Min = rect2.SleeveBoundingBoxMinY;
        x2Max = rect2.SleeveBoundingBoxMaxX;
        y2Max = rect2.SleeveBoundingBoxMaxY;
    }
    
    // Calculate distance between rectangles
    double dx = Math.Max(0, Math.Max(x1Min - x2Max, x2Min - x1Max));
    double dy = Math.Max(0, Math.Max(y1Min - y2Max, y2Min - y1Max));
    double distance = Math.Sqrt(dx * dx + dy * dy);
    
    return distance <= toleranceDist;
}
```

## 5. Orientation Field Names and Coordinate Systems

### 5.1 Orientation Field Names
**CRITICAL**: There are multiple orientation-related fields in the system. Use the correct one:

| Field Name | Location | Purpose | Values |
|------------|----------|---------|---------|
| `MepElementOrientationDirection` | ClashZone | ✅ **CORRECT** - Used for clustering | "X", "Y", "" (empty for floors) |
| `HostOrientation` | ClashZone | ❌ **WRONG** - Not used in clustering | Various values |
| `Orientation` | SleeveData | Source data from _CLUSTER.xml | "X", "Y", "" (empty for floors) |

### 5.2 Coordinate System Usage
Based on host type and orientation:

| Host Type | Orientation | Coordinate System | Usage |
|-----------|-------------|-------------------|-------|
| Floor | "" (empty) | X,Y | Ignore Z coordinate |
| Wall | "X" | Y,Z | Ignore X coordinate |
| Wall | "Y" | X,Z | Ignore Y coordinate |
| Unknown | Any | X,Y,Z | Use 3D distance |

### 5.3 Orientation Detection Method
**File**: `Services/UniversalClusterService.cs`

**Key Method**: `GetOrientationFromClashZone()`
```csharp
private string GetOrientationFromClashZone(FamilyInstance sleeve)
{
    try
    {
        var mepElementIdParam = sleeve.LookupParameter("MEP_ElementId");
        if (mepElementIdParam != null)
        {
            long mepElementId = mepElementIdParam.AsInteger();
            
            // Find clash zone for this MEP element
            var clashZone = GetClashZoneByMepElementId(mepElementId);
            if (clashZone != null)
            {
                // ✅ CORRECT: Use MepElementOrientationDirection
                return clashZone.MepElementOrientationDirection ?? "";
                
                // ❌ WRONG: Don't use HostOrientation
                // return clashZone.HostOrientation ?? "";
            }
        }
        
        return "";
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[UniversalClusterService] Error getting orientation: {ex.Message}");
        return "";
    }
}
```

### 5.4 Distance Calculation Based on Orientation
**File**: `Services/UniversalClusterService.cs`

**Key Method**: `CheckBoundingBoxOverlapWithOrientation()`
```csharp
private bool CheckBoundingBoxOverlapWithOrientation(ClashZone rect1, ClashZone rect2, double toleranceDist, string orientation)
{
    double x1Min, y1Min, x1Max, y1Max;
    double x2Min, y2Min, x2Max, y2Max;
    
    if (orientation == "Y") // Y-oriented walls
    {
        // Use Y,Z coordinates (ignore X)
        x1Min = rect1.SleeveBoundingBoxMinY;
        y1Min = rect1.SleeveBoundingBoxMinZ;
        x1Max = rect1.SleeveBoundingBoxMaxY;
        y1Max = rect1.SleeveBoundingBoxMaxZ;
        
        x2Min = rect2.SleeveBoundingBoxMinY;
        y2Min = rect2.SleeveBoundingBoxMinZ;
        x2Max = rect2.SleeveBoundingBoxMaxY;
        y2Max = rect2.SleeveBoundingBoxMaxZ;
    }
    else if (orientation == "X") // X-oriented walls
    {
        // Use X,Z coordinates (ignore Y)
        x1Min = rect1.SleeveBoundingBoxMinX;
        y1Min = rect1.SleeveBoundingBoxMinZ;
        x1Max = rect1.SleeveBoundingBoxMaxX;
        y1Max = rect1.SleeveBoundingBoxMaxZ;
        
        x2Min = rect2.SleeveBoundingBoxMinX;
        y2Min = rect2.SleeveBoundingBoxMinZ;
        x2Max = rect2.SleeveBoundingBoxMaxX;
        y2Max = rect2.SleeveBoundingBoxMaxZ;
    }
    else // Floors (empty orientation)
    {
        // Use X,Y coordinates (ignore Z)
        x1Min = rect1.SleeveBoundingBoxMinX;
        y1Min = rect1.SleeveBoundingBoxMinY;
        x1Max = rect1.SleeveBoundingBoxMaxX;
        y1Max = rect1.SleeveBoundingBoxMaxY;
        
        x2Min = rect2.SleeveBoundingBoxMinX;
        y2Min = rect2.SleeveBoundingBoxMinY;
        x2Max = rect2.SleeveBoundingBoxMaxX;
        y2Max = rect2.SleeveBoundingBoxMaxY;
    }
    
    // Calculate 2D distance
    double dx = Math.Max(0, Math.Max(x1Min - x2Max, x2Min - x1Max));
    double dy = Math.Max(0, Math.Max(y1Min - y2Max, y2Min - y1Max));
    double distance = Math.Sqrt(dx * dx + dy * dy);
    
    return distance <= toleranceDist;
}
```

## 6. Critical Methods Checklist

### 5.1 Phase 1 Methods (Refresh Button)
- [ ] `UniversalSleevePlacerService.PlaceSleevesForClashZones()`
- [ ] `SleeveCoordinateService.RegenerateClusterXmlAfterPlacement()`
- [ ] `SleeveCoordinateService.SaveSleeveDataToClusterXml()`

### 5.2 Phase 2 Methods (OK Button)
- [ ] `UniversalClusterService.ClusterSleeves()`
- [ ] `UniversalClusterService.LoadClashZoneCache()`
- [ ] `UniversalClusterService.CheckBoundingBoxOverlapWithOrientation()`
- [ ] `UniversalClusterService.PlaceClusterSleeve()`
- [ ] `UniversalClusterService.UpdateClashZoneFlagsForCluster()`

### 5.3 Data Conversion Methods
- [ ] `SleeveDataService.LoadSleeveDataFromClusterXml()`
- [ ] `SleeveData` to `ClashZone` conversion
- [ ] Bounding box calculation from 4 corners

## 6. Debugging Checklist

### 6.1 Verify _CLUSTER.xml Creation
```bash
# Check if _CLUSTER.xml files exist
dir "C:\Users\{username}\AppData\Roaming\JSE_MEP_Openings\Projects\Default\Filters" | findstr "_CLUSTER.xml"
```

### 6.2 Verify Cache Loading
Check `cluster_debug.log` for:
```
[CACHE] Loaded X sleeves from _CLUSTER.xml (filtered for {category})
```

### 6.3 Verify Cache Key Matching
Check `cluster_debug.log` for:
```
Found X sleeves in current cache (out of Y total)
```

### 6.4 Verify Coordinate Values
Check `cluster_debug.log` for:
```
[DISTANCE-DEBUG] Rect1: Min=(X.XXX, Y.YYY), Max=(X.XXX, Y.YYY)
```
Should NOT be `(0.000, 0.000)`

## 7. Common Issues and Solutions

### 7.1 Issue: Pipe clustering not working (SleeveInstanceId = -1)
**Cause**: Pipe strategy executed but fell through to "NO STRATEGY MATCHED" fallback
**Solution**: Fixed logic flow in `UniversalSleevePlacerService.cs` to use proper `else if` statements
```csharp
// BEFORE (WRONG): Pipe strategy executed but continued to other strategies
if (isPipesCategory) { /* pipe logic */ }
else if (_strategy is DamperPlacementStrategy) { /* damper logic */ }

// AFTER (CORRECT): Only one strategy executes per clash zone
if (isPipesCategory) { /* pipe logic */ }
else if (_strategy is DamperPlacementStrategy) { /* damper logic */ }
```

### 7.2 Issue: Cache shows 0 sleeves despite loading data
**Cause**: Cache key mismatch (SleeveInstanceId vs MepElementId)
**Solution**: Use MepElementId as cache key consistently

### 7.3 Issue: Coordinates show (0.000, 0.000)
**Cause**: Reading from regular XML instead of _CLUSTER.xml
**Solution**: Ensure LoadClashZoneCache reads from _CLUSTER.xml files

### 7.4 Issue: No clustering occurs
**Cause**: Sleeves not found in cache
**Solution**: Verify cache key matching and MepElementId parameter

### 7.5 Issue: Wrong file pattern matching
**Cause**: Category name mismatch (spaces vs underscores)
**Solution**: Handle both patterns: `*_{category}_CLUSTER.xml` and `*_{category with spaces}_CLUSTER.xml`

### 7.6 Issue: Orientation shows empty in logs
**Cause**: Wrong field name used for orientation lookup
**Solution**: Use `MepElementOrientationDirection` instead of `HostOrientation`

### 7.9 Issue: Orientation showing "Z" instead of correct wall orientation
**Cause**: `SleeveCoordinateService` recalculating orientation instead of using pre-calculated clash zone data
**Solution**: Fixed to read orientation from clash zone data using MEP_ElementId lookup
```csharp
// BEFORE (WRONG): Recalculating orientation
var wallDirectionParam = mepElement.LookupParameter("Wall Direction Type");
var orientation = wallDirectionType switch { ... }; // ❌ WRONG: Recalculating

// AFTER (CORRECT): Using clash zone data
var clashZone = FindClashZoneByMepElementId(mepElementIdValue);
return clashZone?.MepElementOrientationDirection ?? ""; // ✅ CORRECT: Use pre-calculated data
```

## 8. Testing Procedure

### 8.1 Phase 1 Testing
1. Click Refresh button
2. Verify individual sleeves are placed
3. Check `_CLUSTER.xml` file is created with correct coordinates
4. Verify `all_sleeve_coordinates.log` shows actual coordinates

### 8.2 Phase 2 Testing
1. Click OK button
2. Check `cluster_debug.log` for cache loading
3. Verify sleeves are found in cache
4. Check distance calculations show real coordinates
5. Verify clustering occurs and cluster sleeves are placed

## 9. File Naming Conventions

### 9.1 XML File Patterns
- Regular clash zones: `{FilterName}_{Category}.xml`
- Sleeve coordinates: `{FilterName}_{Category}_CLUSTER.xml`
- Conditions: `{FilterName}_{Category}_CONDITIONS.xml`

### 9.2 Category Name Handling
- Input category: `duct_accessories`
- File pattern 1: `*_duct_accessories_CLUSTER.xml`
- File pattern 2: `*_Duct Accessories_CLUSTER.xml` (with spaces)

## 10. Performance Considerations

### 10.1 Cache Usage
- Load cache once at start of clustering
- Use O(1) lookup by MepElementId
- Filter by category during loading, not after

### 10.2 Coordinate Calculation
- Calculate bounding box from 4 corners once
- Store in cache for reuse
- Avoid repeated Revit API calls during clustering

### 10.3 File I/O
- Read _CLUSTER.xml once per clustering session
- Write cluster flags to XML after clustering
- Use most recently modified file if multiple exist

---

**CRITICAL SUCCESS FACTORS**:
1. ✅ Use MepElementId as cache key consistently
2. ✅ Read from _CLUSTER.xml files, not regular XML
3. ✅ Calculate bounding box from 4 corner coordinates
4. ✅ Handle both underscore and space patterns in filenames
5. ✅ Verify coordinates are not (0.000, 0.000) in debug logs
6. ✅ Use `MepElementOrientationDirection` for orientation detection (NOT `HostOrientation`)
7. ✅ Apply correct coordinate system based on host type and orientation
8. ✅ Verify orientation shows "X", "Y", or "" in debug logs (not empty)
9. ✅ **NEW**: Ensure SleeveCoordinateService creates _CLUSTER.xml files after sleeve placement
10. ✅ **NEW**: Fix pipe strategy logic flow to prevent fallthrough to "NO STRATEGY MATCHED"
11. ✅ **NEW**: Fix SleeveCoordinateService category mapping to return "Pipes" instead of "Pipe Accessories"
12. ✅ **NEW**: Fix SleeveCoordinateService orientation detection to use clash zone data instead of recalculating

## 11. Complete Flow Summary (Updated 2025-01-24)

### 11.1 Refresh Button Flow
1. **ClashZoneService** → Detects intersections and creates clash zones
2. **UniversalSleevePlacerService** → Places individual sleeves in Revit
   - **Pipe Strategy Fix**: Proper `else if` logic prevents fallthrough
3. **SleeveCoordinateService** → Reads sleeves from Revit and creates _CLUSTER.xml
   - Waits 500ms for Revit to update
   - Groups sleeves by category
   - Creates both singular and plural category files
4. **Result**: _CLUSTER.xml files ready for clustering

### 11.2 OK Button Flow  
1. **UniversalClusterService** → Loads sleeve data from _CLUSTER.xml files
2. **Cache Building** → Uses MepElementId as keys
3. **Proximity Clustering** → Groups sleeves by host type, system type, orientation
4. **Cluster Placement** → Places cluster sleeves and deletes individual sleeves
5. **Result**: Clustered sleeves replace individual sleeves

### 11.3 Key Services Integration
- **SleevePlacementExternalEvent** → Calls SleeveCoordinateService after placement
- **SleeveCoordinateService** → Creates _CLUSTER.xml files with actual Revit coordinates  
- **UniversalClusterService** → Reads from _CLUSTER.xml files for clustering
- **UniversalSleevePlacerService** → Fixed pipe strategy logic flow
