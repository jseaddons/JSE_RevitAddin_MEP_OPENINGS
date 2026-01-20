# 🎯 CLUSTER ALGORITHM - CONVOID METHODOLOGY ALIGNED

## 📚 References Reviewed

1. ✅ **SIMPLIFIED_FAMILY_STRATEGY_CONVOID_APPROACH.md** - Universal family approach (2 families for all categories)
2. ✅ **OpeningClusterModification.md** - Original pipe clustering logic (proximity-based)
3. ✅ **ClusterSleeveMergePlan.md** - Manual merge of 2 clusters (not applicable to auto-clustering)
4. ✅ **RectangularSleeveClusterCommandV2.cs** - Working BFS clustering implementation

---

## 🔍 CONVOID Methodology Key Principles

### **1. Universal Families (CRITICAL)**
```
❌ OLD (Category-Specific):
- DuctOpeningOnWall
- PipeOpeningOnWall  
- ClusterOpeningOnWallX (special cluster family)

✅ CONVOID (Universal):
- OpeningOnWall     ← For ALL categories (individual AND cluster)
- OpeningOnSlab     ← For ALL categories (individual AND cluster)
```

**Differentiation:** By `MEP_Category` parameter, NOT family name!

### **2. Parameter-Based Identification**
```csharp
// Mark parameter for scheduling
sleeve.LookupParameter("Mark")?.Set("DS-001");  // Duct Sleeve #1
sleeve.LookupParameter("Mark")?.Set("PS-002");  // Pipe Sleeve #2

// Category parameter for filtering
sleeve.LookupParameter("MEP_Category")?.Set("Ducts");
sleeve.LookupParameter("MEP_Category")?.Set("Pipes");

// Cluster flag
sleeve.LookupParameter("IsCluster")?.Set(1); // Yes/No parameter
```

### **3. Simple Family Selection**
```csharp
// CONVOID approach - ONE LINE!
string familyName = (hostType == "Wall" || hostType == "Framing")
    ? "OpeningOnWall"
    : "OpeningOnSlab";

// No switch statements, no category checks!
```

---

## 🏗️ Current Implementation vs CONVOID

### **Current (4 Universal Families)**
```
✓ RectangularOpeningOnWall
✓ CircularOpeningOnWall
✓ RectangularOpeningOnSlab
✓ CircularOpeningOnSlab
```

**Status:** ⚠️ **Almost CONVOID-aligned!**
- ✅ Uses universal families (not category-specific)
- ✅ Parameters identify category (`MEP_Category`)
- ⚠️ Still distinguishes Rectangular vs Circular in family name

### **Pure CONVOID (2 Universal Families)**
```
✓ OpeningOnWall     ← Handles BOTH rectangular AND circular
✓ OpeningOnSlab     ← Handles BOTH rectangular AND circular
```

**How?** Shape parameter or nested families for circular vs rectangular geometry.

---

## 🔧 Cluster Algorithm - CONVOID-Aligned

### **Phase 1: Collection & Category Filtering**

```csharp
// Step 1: Collect ALL universal families
var allSleeves = new FilteredElementCollector(doc)
    .OfClass(typeof(FamilyInstance))
    .Where(fi => 
    {
        var famName = fi.Symbol?.Family?.Name ?? "";
        
        // Current (4 families):
        return famName == "RectangularOpeningOnWall" ||
               famName == "CircularOpeningOnWall" ||
               famName == "RectangularOpeningOnSlab" ||
               famName == "CircularOpeningOnSlab";
        
        // Future CONVOID (2 families):
        // return famName == "OpeningOnWall" || famName == "OpeningOnSlab";
    })
    .ToList();

// Step 2: Filter by MEP_Category (CRITICAL for preventing cross-category clustering)
var categorySleeves = allSleeves.Where(sleeve => 
{
    var categoryParam = sleeve.LookupParameter("MEP_Category");
    return categoryParam?.AsString() == targetCategory; // "Ducts", "Pipes", etc.
}).ToList();

DebugLogger.Info($"[Cluster] Found {allSleeves.Count} total sleeves, {categorySleeves.Count} for '{targetCategory}'");
```

**Why Category Filter is CRITICAL:**
- ❌ Without filter: Ducts cluster with Pipes (same family name!)
- ✅ With filter: Only ducts cluster with ducts

---

### **Phase 2: Grouping (Same Host, Orientation, Category)**

```csharp
// Group sleeves by common properties
var groups = categorySleeves.GroupBy(sleeve => new 
{
    hostType = GetHostType(sleeve),         // "Wall", "Floor", "Structural Framing"
    orientation = GetOrientation(sleeve),   // "X", "Y", "Z" (from HostOrientation parameter)
    // systemType already filtered by targetCategory
});

private string GetHostType(FamilyInstance sleeve)
{
    var famName = sleeve.Symbol.Family.Name.ToLower();
    return famName.Contains("onwall") ? "Wall" 
         : famName.Contains("onslab") ? "Floor" 
         : "Unknown";
}

private string GetOrientation(FamilyInstance sleeve)
{
    var orientParam = sleeve.LookupParameter("HostOrientation");
    return orientParam?.AsString() ?? "";
}
```

**Result:** Sleeves are grouped by:
1. ✅ Same structural host type (Wall vs Floor)
2. ✅ Same wall orientation (X vs Y walls)
3. ✅ Same MEP category (Ducts vs Pipes) - via pre-filter

---

### **Phase 3: Proximity Clustering (BFS + Bounding Box)**

#### **Step 3.1: Precompute Data**
```csharp
// Cache bounding boxes (expensive API call - do once per sleeve)
var bboxes = new Dictionary<FamilyInstance, BoundingBoxXYZ>();
foreach (var sleeve in groupSleeves)
{
    var bbox = sleeve.get_BoundingBox(null);
    if (bbox != null) bboxes[sleeve] = bbox;
}
```

#### **Step 3.2: Build Spatial Hash Grid (Performance)**
```csharp
// Divide 3D space into grid cells (size = JoinOpeningsDistance)
double cellSize = JoinOpeningsDistance; // e.g., 200mm from config
var grid = new Dictionary<(int, int, int), List<FamilyInstance>>();

foreach (var sleeve in groupSleeves)
{
    var bbox = bboxes[sleeve];
    
    // Hash sleeve into all grid cells it overlaps
    int min_ix = (int)Math.Floor(bbox.Min.X / cellSize);
    int max_ix = (int)Math.Floor(bbox.Max.X / cellSize);
    int min_iy = (int)Math.Floor(bbox.Min.Y / cellSize);
    int max_iy = (int)Math.Floor(bbox.Max.Y / cellSize);
    int min_iz = (int)Math.Floor(bbox.Min.Z / cellSize);
    int max_iz = (int)Math.Floor(bbox.Max.Z / cellSize);
    
    for (int gx = min_ix; gx <= max_ix; gx++)
        for (int gy = min_iy; gy <= max_iy; gy++)
            for (int gz = min_iz; gz <= max_iz; gz++)
            {
                var key = (gx, gy, gz);
                if (!grid.ContainsKey(key)) grid[key] = new List<FamilyInstance>();
                grid[key].Add(sleeve);
            }
}
```

**Performance Impact:**
```
Without Grid:
  100 sleeves → 4,950 comparisons (O(n²))
  1000 sleeves → 499,500 comparisons 💥

With Grid (cell size = 200mm):
  100 sleeves → ~800 comparisons (O(n × k), k ≈ 8)
  1000 sleeves → ~8,000 comparisons ✅ 62x faster!
```

#### **Step 3.3: BFS Graph Clustering**
```csharp
var unprocessed = new HashSet<FamilyInstance>(groupSleeves);
var clusters = new List<List<FamilyInstance>>();

while (unprocessed.Count > 0)
{
    var start = unprocessed.First();
    var queue = new Queue<FamilyInstance>();
    var cluster = new List<FamilyInstance>();
    
    queue.Enqueue(start);
    unprocessed.Remove(start);
    
    // Breadth-First Search to find all connected sleeves
    while (queue.Count > 0)
    {
        var current = queue.Dequeue();
        cluster.Add(current);
        
        // Find neighbors using spatial grid
        var neighbors = FindNeighborsViaGrid(current, grid, bboxes, JoinOpeningsDistance);
        
        foreach (var neighbor in neighbors)
        {
            if (unprocessed.Remove(neighbor)) // Not yet visited
            {
                queue.Enqueue(neighbor);
            }
        }
    }
    
    if (cluster.Count > 1) // Only save clusters with 2+ sleeves
        clusters.Add(cluster);
}
```

**Why BFS?**
```
Example: 3 sleeves in a row

[A]--150mm--[B]--150mm--[C]

A ↔ B: 150mm (< 200mm threshold) ✅
B ↔ C: 150mm (< 200mm threshold) ✅
A ↔ C: 300mm (> 200mm threshold) ❌

BFS Result: A, B, C all in SAME cluster (transitive connection)
  A connects to B
  B connects to C
  Therefore A, B, C are all connected!

Without BFS (pairwise only):
  Cluster 1: {A, B}
  Cluster 2: {B, C}  ← B appears twice! ❌
```

#### **Step 3.4: Neighbor Detection (Bounding Box Overlap)**
```csharp
private List<FamilyInstance> FindNeighborsViaGrid(
    FamilyInstance sleeve, 
    Dictionary<(int, int, int), List<FamilyInstance>> grid,
    Dictionary<FamilyInstance, BoundingBoxXYZ> bboxes,
    double tolerance)
{
    var bbox1 = bboxes[sleeve];
    var candidates = new HashSet<FamilyInstance>();
    
    // Get all grid cells within tolerance of this sleeve's bbox
    double cellSize = tolerance;
    int min_ix = (int)Math.Floor((bbox1.Min.X - tolerance) / cellSize);
    int max_ix = (int)Math.Floor((bbox1.Max.X + tolerance) / cellSize);
    int min_iy = (int)Math.Floor((bbox1.Min.Y - tolerance) / cellSize);
    int max_iy = (int)Math.Floor((bbox1.Max.Y + tolerance) / cellSize);
    int min_iz = (int)Math.Floor((bbox1.Min.Z - tolerance) / cellSize);
    int max_iz = (int)Math.Floor((bbox1.Max.Z + tolerance) / cellSize);
    
    for (int gx = min_ix; gx <= max_ix; gx++)
        for (int gy = min_iy; gy <= max_iy; gy++)
            for (int gz = min_iz; gz <= max_iz; gz++)
            {
                if (grid.TryGetValue((gx, gy, gz), out var bucket))
                    candidates.UnionWith(bucket);
            }
    
    // Check actual bounding box overlap for candidates
    var neighbors = new List<FamilyInstance>();
    foreach (var candidate in candidates)
    {
        if (candidate == sleeve) continue;
        
        var bbox2 = bboxes[candidate];
        
        // 3D bounding box overlap check
        bool xOverlap = bbox1.Max.X >= bbox2.Min.X - tolerance && 
                        bbox1.Min.X <= bbox2.Max.X + tolerance;
        bool yOverlap = bbox1.Max.Y >= bbox2.Min.Y - tolerance && 
                        bbox1.Min.Y <= bbox2.Max.Y + tolerance;
        bool zOverlap = bbox1.Max.Z >= bbox2.Min.Z - tolerance && 
                        bbox1.Min.Z <= bbox2.Max.Z + tolerance;
        
        if (xOverlap && yOverlap && zOverlap)
            neighbors.Add(candidate);
    }
    
    return neighbors;
}
```

**Why Bounding Box vs Center-to-Center?**
```
Scenario: Two 400×200mm rectangular sleeves

Sleeve A at (0, 0, 0)     → BBox: X[-200,+200], Y[-100,+100]
Sleeve B at (250, 0, 0)   → BBox: X[+50,+450], Y[-100,+100]

Center-to-Center Distance: 250mm (> 200mm) ❌ Would NOT cluster

Bounding Box Overlap:
  A.BBox.Max.X = 200mm
  B.BBox.Min.X = 50mm
  Gap = 50mm - 200mm = -150mm (overlapping!) ✅ WILL cluster

Bounding box is edge-to-edge, which is what we want for rectangular sleeves!
```

---

### **Phase 4: Cluster Sleeve Placement**

#### **Step 4.1: Calculate Combined Bounding Box**
```csharp
// ClusterBoundingBoxServices.GetClusterBoundingBox(cluster)
BoundingBoxXYZ combinedBbox = null;

foreach (var sleeve in cluster)
{
    var bbox = sleeve.get_BoundingBox(null);
    
    if (combinedBbox == null)
    {
        combinedBbox = new BoundingBoxXYZ 
        { 
            Min = bbox.Min, 
            Max = bbox.Max 
        };
    }
    else
    {
        // Union: expand to include this sleeve
        combinedBbox.Min = new XYZ(
            Math.Min(combinedBbox.Min.X, bbox.Min.X),
            Math.Min(combinedBbox.Min.Y, bbox.Min.Y),
            Math.Min(combinedBbox.Min.Z, bbox.Min.Z)
        );
        combinedBbox.Max = new XYZ(
            Math.Max(combinedBbox.Max.X, bbox.Max.X),
            Math.Max(combinedBbox.Max.Y, bbox.Max.Y),
            Math.Max(combinedBbox.Max.Z, bbox.Max.Z)
        );
    }
}

double width = combinedBbox.Max.X - combinedBbox.Min.X;
double height = combinedBbox.Max.Y - combinedBbox.Min.Y;
double depth = combinedBbox.Max.Z - combinedBbox.Min.Z;
XYZ centroid = (combinedBbox.Min + combinedBbox.Max) / 2.0;
```

**Result:** Cluster sleeve is sized to EXACTLY cover all individual sleeves

#### **Step 4.2: Select Universal Family (CONVOID-Aligned)**
```csharp
// CURRENT (4 families):
bool isCircular = cluster.All(s => s.Symbol.Family.Name.Contains("Circular"));
string familyName = (hostType == "Wall" || hostType == "Framing")
    ? (isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall")
    : (isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab");

// FUTURE CONVOID (2 families):
string familyName = (hostType == "Wall" || hostType == "Framing")
    ? "OpeningOnWall"   // ← ONE family for walls (handles both shapes)
    : "OpeningOnSlab";  // ← ONE family for slabs (handles both shapes)

// Set shape via parameter instead:
clusterSleeve.LookupParameter("Shape")?.Set(isCircular ? "Circular" : "Rectangular");
```

#### **Step 4.3: Place Cluster Sleeve**
```csharp
var familySymbol = FindFamilySymbol(doc, familyName);
if (!familySymbol.IsActive) familySymbol.Activate();

var level = doc.GetElement(cluster.First().LevelId) as Level;

FamilyInstance clusterSleeve = doc.Create.NewFamilyInstance(
    centroid,           // At center of combined bounding box
    familySymbol,       // Universal family
    level,
    StructuralType.NonStructural
);
```

#### **Step 4.4: Set Rotation (Walls only)**
```csharp
if (hostType == "Wall")
{
    double rotationAngle = (orientation == "Y") ? Math.PI / 2 : 0.0;
    
    if (rotationAngle != 0.0)
    {
        Line rotationAxis = Line.CreateBound(centroid, centroid + XYZ.BasisZ);
        ElementTransformUtils.RotateElement(doc, clusterSleeve.Id, rotationAxis, rotationAngle);
    }
}
```

#### **Step 4.5: Set Parameters (CONVOID-Aligned)**
```csharp
// Size parameters
if (hostType == "Wall" || hostType == "Framing")
{
    // Wall families: Width=Y, Height=Z, Depth=X (created in "Left" view)
    clusterSleeve.LookupParameter("Width")?.Set(height);
    clusterSleeve.LookupParameter("Height")?.Set(depth);
    
    double hostThickness = GetHostThickness(cluster.First().Host);
    clusterSleeve.LookupParameter("Depth")?.Set(hostThickness);
}
else // Floor
{
    clusterSleeve.LookupParameter("Width")?.Set(width);
    clusterSleeve.LookupParameter("Height")?.Set(height);
    clusterSleeve.LookupParameter("Depth")?.Set(depth);
}

// ✅ CRITICAL: Set MEP_Category (CONVOID methodology)
clusterSleeve.LookupParameter("MEP_Category")?.Set(targetCategory); // "Ducts", "Pipes", etc.

// ✅ CRITICAL: Mark as cluster
clusterSleeve.LookupParameter("IsCluster")?.Set(1); // Yes/No parameter

// Mark parameter (CONVOID scheduling)
string mark = GenerateClusterMark(targetCategory, clusterIndex); // e.g., "DC-001" (Duct Cluster #1)
clusterSleeve.LookupParameter("Mark")?.Set(mark);

DebugLogger.Info($"[Cluster] Placed {targetCategory} cluster '{mark}' - Size: {width:F0}×{height:F0}×{depth:F0}mm");
```

#### **Step 4.6: Delete Individual Sleeves**
```csharp
foreach (var sleeve in cluster)
{
    doc.Delete(sleeve.Id);
}

DebugLogger.Info($"[Cluster] Deleted {cluster.Count} individual sleeves");
```

---

## 🎯 CONVOID Mark Prefixes

```csharp
private string GenerateClusterMark(string category, int index)
{
    string prefix = category switch
    {
        "Ducts" => "DC",               // Duct Cluster
        "Pipes" => "PC",               // Pipe Cluster
        "Cable Trays" => "CC",         // Cable Tray Cluster
        "Duct Accessories" => "DAC",   // Damper Cluster
        _ => "OC"                      // Generic Opening Cluster
    };
    
    return $"{prefix}-{index:D3}";  // e.g., "DC-001", "PC-002"
}
```

**Individual sleeve marks:**
- DS-001 (Duct Sleeve #1)
- PS-002 (Pipe Sleeve #2)
- CS-003 (Cable Tray Sleeve #3)
- DAS-004 (Damper Sleeve #4)

**Cluster sleeve marks:**
- DC-001 (Duct Cluster #1)
- PC-002 (Pipe Cluster #2)
- CC-003 (Cable Tray Cluster #3)
- DAC-004 (Damper Cluster #4)

---

## ✅ CONVOID Alignment Checklist

### **Implemented ✅**
1. ✅ Universal families (not category-specific)
2. ✅ MEP_Category parameter identifies category
3. ✅ Category filtering before clustering
4. ✅ Parameter-based family selection
5. ✅ IsCluster flag for schedules
6. ✅ Mark parameter for scheduling
7. ✅ Proximity-based clustering (JoinOpeningsDistance from config)
8. ✅ BFS graph clustering (transitive connections)
9. ✅ Bounding box overlap detection (edge-to-edge)
10. ✅ Combined bounding box for cluster size

### **Future CONVOID Improvements 🔜**
1. 🔜 Reduce from 4 families to 2 (OpeningOnWall, OpeningOnSlab)
2. 🔜 Shape parameter instead of separate Circular/Rectangular families
3. 🔜 Nested families for circular geometry
4. 🔜 Mark prefix configuration in settings
5. 🔜 Auto-numbering of marks based on placement order

---

## 📊 Expected Results

### **Before Clustering:**
```
24 Individual Sleeves:
  22× Ducts (MEP_Category="Ducts", Mark="DS-001" to "DS-022")
  2× Dampers (MEP_Category="Duct Accessories", Mark="DAS-001", "DAS-002")
```

### **After Clustering (Ducts only, targetCategory="Ducts"):**
```
3 Duct Clusters:
  DC-001: 8 ducts merged → RectangularOpeningOnWall (Size: 2000×800×200mm)
  DC-002: 10 ducts merged → RectangularOpeningOnWall (Size: 2400×1000×200mm)
  DC-003: 4 ducts merged → CircularOpeningOnWall (Size: Ø800mm)

2 Dampers UNCHANGED:
  DAS-001: RectangularOpeningOnWall (MEP_Category="Duct Accessories")
  DAS-002: RectangularOpeningOnWall (MEP_Category="Duct Accessories")
  
  (Dampers not clustered because targetCategory="Ducts" filter excluded them)
```

### **Schedule View (CONVOID-Style):**
```
Mark      | MEP_Category      | Width  | Height | IsCluster
----------|-------------------|--------|--------|----------
DC-001    | Ducts             | 2000mm | 800mm  | Yes
DC-002    | Ducts             | 2400mm | 1000mm | Yes
DC-003    | Ducts             | Ø800mm | -      | Yes
DAS-001   | Duct Accessories  | 400mm  | 400mm  | No
DAS-002   | Duct Accessories  | 600mm  | 600mm  | No
```

**Filtering in Schedule:**
```
Show Duct Clusters: Filter by MEP_Category = "Ducts" AND IsCluster = "Yes"
Show Individual Dampers: Filter by MEP_Category = "Duct Accessories" AND IsCluster = "No"
```

---

## 🚀 Implementation Status

### **Phase 1: Collection & Filtering** ✅ COMPLETE
- ✅ Collect 4 universal families
- ✅ Filter by MEP_Category parameter
- ✅ Section box filtering

### **Phase 2: Grouping** ✅ COMPLETE
- ✅ Group by (HostType, Orientation, Category)
- ✅ Read systemType from MEP_Category parameter

### **Phase 3: Clustering** ✅ COMPLETE
- ✅ Bounding box precomputation
- ✅ Spatial hash grid
- ✅ BFS graph clustering
- ✅ Bounding box overlap detection

### **Phase 4: Placement** ✅ COMPLETE
- ✅ Combined bounding box calculation
- ✅ Universal family selection
- ✅ Cluster sleeve placement
- ✅ Parameter setting (MEP_Category, IsCluster, Mark)
- ✅ Individual sleeve deletion

### **Phase 5: Integration** ⏳ PENDING
- ⏳ Trigger cluster command after each category placement
- ⏳ Update SleevePlacementExternalEvent
- ⏳ Update ClashZone XML with cluster metadata
- ⏳ Test with all categories (Ducts, Pipes, Cable Trays, Dampers)

---

## 📝 Next Steps

1. ✅ **Planning Complete** - This document
2. ⏭️ **Revert Premature Changes** - Undo direct edits to RectangularSleeveClusterCommandV2.cs
3. ⏭️ **Code Review** - Verify current implementation against this plan
4. ⏭️ **Testing** - Test cluster command with `targetCategory="Ducts"`
5. ⏭️ **Integration** - Add cluster trigger to SleevePlacementExternalEvent
6. ⏭️ **Validation** - Verify no cross-category clustering

---

## 📚 Key Takeaways

1. ✅ **CONVOID = Universal Families** - Minimize family count, maximize flexibility
2. ✅ **Parameters over Family Names** - MEP_Category, Mark, IsCluster for identification
3. ✅ **Proximity Clustering** - BFS + Bounding Box for accurate merging
4. ✅ **Category Filtering is CRITICAL** - Prevents cross-category clustering
5. ✅ **Spatial Hash Grid** - Essential for performance with 100+ sleeves
6. ✅ **Transitive Connections** - BFS ensures all connected sleeves cluster together

**This is the right architectural approach - aligned with CONVOID best practices!** 🎯

