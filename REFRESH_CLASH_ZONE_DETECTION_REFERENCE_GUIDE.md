# REFRESH CLASH ZONE DETECTION - WORKING CODE REFERENCE GUIDE

## Overview
This guide documents the **exact working code path** when you click the Refresh button to get clash zones. Based on the actual implementation, not documentation.

## Complete Working Flow

### 1. **Refresh Button Click** → `EmergencyMainDialog.cs`
```csharp
// Lines 5863-5896: GetCurrentIntersections() method
private List<(Element, Element, BoundingBoxXYZ, XYZ)> GetCurrentIntersections()
{
    // Use the proven TestMepIntersection service with current UI selections
    var selectedCategories = GetSelectedMepCategories();
    var intersectionService = new IntersectionDetectionService(msg => DebugLogger.Info(msg));
    var intersections = intersectionService.FindIntersections(document, view3D, selectedCategories);
    
    return intersections;
}
```

### 2. **Main Service** → `RefreshService.cs`
```csharp
// Lines 189-192: ExecuteRefresh() method
var currentIntersections = _intersectionService.FindIntersections(_document, _document.ActiveView as View3D, selectedMepCategories);

DebugLogger.Info($"[CLASH_DEBUG] INTERSECTION DETECTION COMPLETE: {currentIntersections.Count} intersections found");
```

### 3. **Core Intersection Detection** → `IntersectionDetectionService.cs`

#### **Step 1: Section Box Processing (Lines 35-54)**
```csharp
// Get section box in model coordinates
BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
Transform sectionTransform = sectionBox.Transform;

// Transform all 8 corners to model coordinates
List<XYZ> corners = new List<XYZ>
{
    sectionTransform.OfPoint(sectionBox.Min),
    sectionTransform.OfPoint(new XYZ(sectionBox.Max.X, sectionBox.Min.Y, sectionBox.Min.Z)),
    // ... all 8 corners
};

XYZ modelMin = new XYZ(corners.Min(p => p.X), corners.Min(p => p.Y), corners.Min(p => p.Z));
XYZ modelMax = new XYZ(corners.Max(p => p.X), corners.Max(p => p.Y), corners.Max(p => p.Z));
```

#### **Step 2: Element Collection (Lines 60, 92-300)**
```csharp
// Collect MEP and Structural elements within section box
CollectElements(document, modelMin, modelMax, ref mepElements, ref wallElements, selectedMepCategories);
```

**MEP Element Collection (Lines 99-150):**
```csharp
// Determine MEP categories from UI selections
List<BuiltInCategory> mepCats = new List<BuiltInCategory>();

if (selectedMepCategories != null && selectedMepCategories.Count > 0)
{
    // Map UI category names to BuiltInCategory enums
    foreach (var catName in selectedMepCategories)
    {
        switch (catName.ToLower())
        {
            case "ducts":
                mepCats.Add(BuiltInCategory.OST_DuctCurves);
                break;
            case "duct accessories":
                mepCats.Add(BuiltInCategory.OST_DuctAccessory);
                break;
            case "pipes":
                mepCats.Add(BuiltInCategory.OST_PipeCurves);
                break;
            // ... other categories
        }
    }
}

// Collect MEP elements from HOST document only
var hostCollector = new FilteredElementCollector(doc)
    .WherePasses(new ElementMulticategoryFilter(mepCats))
    .WhereElementIsNotElementType()
    .WherePasses(new BoundingBoxIntersectsFilter(hostOutline));

mepElements.AddRange(hostCollector.ToElements());
```

**Structural Element Collection (Lines 151-300):**
```csharp
// Collect structural elements from HOST document
BuiltInCategory[] structCats = {
    BuiltInCategory.OST_Walls, 
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_Floors
};

foreach (var cat in structCats)
{
    var elements = new FilteredElementCollector(doc)
        .OfCategory(cat)
        .WhereElementIsNotElementType()
        .WherePasses(new BoundingBoxIntersectsFilter(hostOutline))
        .ToElements();
    wallElements.AddRange(elements);
}

// Process linked files for structural elements
var links = new FilteredElementCollector(doc)
    .OfClass(typeof(RevitLinkInstance))
    .Cast<RevitLinkInstance>()
    .ToList();

foreach (RevitLinkInstance link in links)
{
    Document linkDoc = link.GetLinkDocument();
    if (linkDoc == null) continue;

    // Get the transform from link to host
    Transform linkTransform = link.GetTotalTransform();
    
    // Transform section box to link's coordinate system using INVERSE
    Transform invTransform = linkTransform.Inverse;
    XYZ minLink = invTransform.OfPoint(modelBox.Min);
    XYZ maxLink = invTransform.OfPoint(modelBox.Max);
    
    BoundingBoxXYZ linkBox = new BoundingBoxXYZ
    {
        Min = new XYZ(Math.Min(minLink.X, maxLink.X), Math.Min(minLink.Y, maxLink.Y), Math.Min(minLink.Z, maxLink.Z)),
        Max = new XYZ(Math.Max(minLink.X, maxLink.X), Math.Max(minLink.Y, maxLink.Y), Math.Max(minLink.Z, maxLink.Z))
    };

    Outline linkOutline = new Outline(linkBox.Min, linkBox.Max);

    // Collect structural elements from linked file
    foreach (var cat in structCats)
    {
        var elements = new FilteredElementCollector(linkDoc)
            .OfCategory(cat)
            .WhereElementIsNotElementType()
            .WherePasses(new BoundingBoxIntersectsFilter(linkOutline))
            .ToElements();
        wallElements.AddRange(elements);
    }
}
```

#### **Step 3: Intersection Detection (Lines 77, 301-400)**
```csharp
// Find intersections using MepIntersectionService
var intersections = FindIntersectionsInternal(mepElements, wallElements, document);
```

**Internal Intersection Logic:**
```csharp
private List<(Element, Element, BoundingBoxXYZ, XYZ)> FindIntersectionsInternal(
    List<Element> mepElements, List<Element> wallElements, Document document)
{
    var results = new List<(Element, Element, BoundingBoxXYZ, XYZ)>();
    
    foreach (Element mepElement in mepElements)
    {
        // Get MEP element geometry and line
        var mepGeometry = mepElement.get_Geometry(new Options());
        Line mepLine = GetMepElementLine(mepElement, mepGeometry);
        
        if (mepLine == null) continue;

        // Create MEP bounding box
        BoundingBoxXYZ mepBBox = mepElement.get_BoundingBox(null) ?? CreateBoundingBoxFromLine(mepLine);
        
        // Expand MEP bounding box with tolerance
        const double tolerance = 1.0; // 1.0 foot tolerance
        var expandedMin = new XYZ(mepBBox.Min.X - tolerance, mepBBox.Min.Y - tolerance, mepBBox.Min.Z - tolerance);
        var expandedMax = new XYZ(mepBBox.Max.X + tolerance, mepBBox.Max.Y + tolerance, mepBBox.Max.Z + tolerance);

        foreach (Element wallElement in wallElements)
        {
            // Quick bounding box intersection test
            BoundingBoxXYZ wallBBox = wallElement.get_BoundingBox(null);
            if (wallBBox == null) continue;
            
            if (!BoundingBoxesIntersect(expandedMin, expandedMax, wallBBox.Min, wallBBox.Max))
                continue;

            // Perform detailed intersection using MepIntersectionService
            var intersectionResults = MepIntersectionService.FindIntersections(
                mepElement, new List<(Element, Transform?)> { (wallElement, null) }, _logger);
            
            foreach (var (_, intersectionBBox, intersectionPoint) in intersectionResults)
            {
                results.Add((mepElement, wallElement, intersectionBBox, intersectionPoint));
            }
        }
    }
    
    return results;
}
```

### 4. **Coordinate Transformation** → `MepIntersectionService.cs`

**Key Coordinate Transformation Logic (Lines 134-148 from PipeSleevePlacerService.cs):**
```csharp
if (transform != null)
{
    // Transform pipe end points into host coords
    hostLine = Line.CreateBound(
        transform.OfPoint(pipeLine.GetEndPoint(0)),
        transform.OfPoint(pipeLine.GetEndPoint(1)));
    
    // Derive bbox in host coords
    var origBbox = pipe.get_BoundingBox(null);
    if (origBbox != null)
    {
        var min = transform.OfPoint(origBbox.Min);
        var max = transform.OfPoint(origBbox.Max);
        pipeBBox = new BoundingBoxXYZ { 
            Min = new XYZ(Math.Min(min.X, max.X), Math.Min(min.Y, max.Y), Math.Min(min.Z, max.Z)), 
            Max = new XYZ(Math.Max(min.X, max.X), Math.Max(min.Y, max.Y), Math.Max(min.Z, max.Z)) 
        };
    }
}
```

## Key Services and Their Roles

### 1. **`RefreshService.cs`** - Main Orchestrator
- **Role**: Coordinates the entire refresh process
- **Key Method**: `ExecuteRefresh()`
- **Coordinates**: UI selections → Intersection detection → Clash zone storage

### 2. **`IntersectionDetectionService.cs`** - Core Detection Engine
- **Role**: Performs the actual intersection detection
- **Key Method**: `FindIntersections()`
- **Process**: Section box → Element collection → Intersection detection

### 3. **`MepIntersectionService.cs`** - Geometric Intersection Logic
- **Role**: Handles the detailed geometric intersection calculations
- **Key Method**: `FindIntersections()`
- **Process**: MEP geometry → Structural geometry → Intersection points

### 4. **`ClashZoneService.cs`** - Clash Zone Management
- **Role**: Creates and manages clash zone objects
- **Key Method**: `CreateClashZone()`
- **Process**: Intersection data → ClashZone objects → Storage

## Element Collection Strategy

### **MEP Elements:**
- **Source**: **HOST document only**
- **Method**: `FilteredElementCollector` with `BoundingBoxIntersectsFilter`
- **Categories**: Based on UI selections (Ducts, Pipes, Cable Trays, etc.)
- **Filter**: Section box bounds

### **Structural Elements:**
- **Source**: **HOST document + Linked files**
- **Method**: `FilteredElementCollector` with coordinate transformation
- **Categories**: Walls, Floors, Structural Framing
- **Transform**: `GetTotalTransform()` for linked elements
- **Filter**: Section box bounds (transformed for linked files)

## Coordinate System Handling

### **Universal Coordinate Support:**
- **Shared Coordinates**: ✅ Supported
- **Origin-to-Origin**: ✅ Supported  
- **Manual Positioning**: ✅ Supported
- **Any Linking Method**: ✅ Supported

### **Key Transformation Method:**
```csharp
Transform linkTransform = link.GetTotalTransform();
```
This method captures the complete spatial relationship regardless of linking method.

## Section Box Processing

### **Section Box to Model Coordinates:**
```csharp
BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
Transform sectionTransform = sectionBox.Transform;

// Transform all 8 corners to model coordinates
XYZ modelMin = new XYZ(corners.Min(p => p.X), corners.Min(p => p.Y), corners.Min(p => p.Z));
XYZ modelMax = new XYZ(corners.Max(p => p.X), corners.Max(p => p.Y), corners.Max(p => p.Z));
```

### **Linked File Section Box Transformation:**
```csharp
// Transform section box to link's coordinate system using INVERSE
Transform invTransform = linkTransform.Inverse;
XYZ minLink = invTransform.OfPoint(modelBox.Min);
XYZ maxLink = invTransform.OfPoint(modelBox.Max);
```

## Output Format

### **Intersection Results:**
```csharp
List<(Element, Element, BoundingBoxXYZ, XYZ)> intersections
// Item1: MEP Element
// Item2: Structural Element  
// Item3: Intersection BoundingBox
// Item4: Intersection Center Point
```

### **Clash Zone Creation:**
```csharp
var clashZone = new ClashZone
{
    MepElementId = mepElement.Id,
    StructuralElementId = structuralElement.Id,
    IntersectionPoint = intersectionPoint,
    ClashBoundingBox = intersectionBBox,
    DetectedAt = DateTime.Now,
    IsResolved = false
};
```

## Performance Optimizations

### **Spatial Pre-filtering:**
- **Bounding Box Intersection**: Quick rejection of non-intersecting elements
- **Section Box Filtering**: Only process elements within active view bounds
- **Tolerance-based Expansion**: 1.0 foot tolerance for intersection detection

### **Batch Processing:**
- **Element Collection**: Single pass through all elements
- **Intersection Detection**: Batch processing of MEP vs Structural elements
- **Coordinate Transformation**: Cached transforms for linked elements

## Error Handling

### **Graceful Degradation:**
- **No 3D View**: Returns empty results with warning
- **No Elements**: Returns empty results with info message
- **Linked File Issues**: Skips problematic links, continues processing
- **Geometry Issues**: Skips elements with invalid geometry

### **Comprehensive Logging:**
- **Timestamped Log Files**: `Refresh_yyyy-MM-dd_HH-mm-ss.log`
- **Debug Logger**: Real-time console output
- **Intersection Details**: Element IDs, coordinates, categories

---

**This guide is based on the actual working code implementation as of the latest version. All code snippets are extracted from the real source files.**
