# 🔍 CLUSTER ALGORITHM ANALYSIS - For Universal Families Adaptation

## 📋 Current Working Algorithm (RectangularSleeveClusterCommandV2)

### **Phase 1: Collection & Filtering**

#### **Step 1.1: Collect All Sleeves**
```csharp
// CURRENT (Old Category-Specific Families):
var allSleeves = new FilteredElementCollector(doc)
    .OfClass(typeof(FamilyInstance))
    .Where(fi => fi.Symbol.Family.Name.EndsWith("OpeningOnWall") ||
                 fi.Symbol.Family.Name.EndsWith("OpeningOnSlab"))
    .ToList();

// NEEDED (Universal Families):
var allSleeves = new FilteredElementCollector(doc)
    .OfClass(typeof(FamilyInstance))
    .Where(fi => 
    {
        var famName = fi.Symbol?.Family?.Name ?? "";
        return famName == "RectangularOpeningOnWall" ||
               famName == "CircularOpeningOnWall" ||
               famName == "RectangularOpeningOnSlab" ||
               famName == "CircularOpeningOnSlab";
    })
    .ToList();
```

#### **Step 1.2: Filter by Category (CRITICAL NEW STEP)**
```csharp
// CURRENT: No category filtering (family name was enough)

// NEEDED: Filter by MEP_Category parameter to prevent cross-category clustering
var categorySleeves = string.IsNullOrEmpty(targetCategory)
    ? allSleeves
    : allSleeves.Where(sleeve => 
    {
        var categoryParam = sleeve.LookupParameter("MEP_Category");
        return categoryParam?.AsString() == targetCategory; // "Ducts", "Pipes", etc.
    })
    .ToList();
```

**Why Critical?** 
- Old system: DuctOpeningOnWall ≠ PipeOpeningOnWall (different families)
- New system: Both use RectangularOpeningOnWall (same family!)
- **Without category filter, ducts WILL cluster with pipes!**

#### **Step 1.3: Section Box Filtering**
```csharp
// Filter to only sleeves visible in active 3D section box
var filtered = SectionBoxHelper.FilterElementsBySectionBox(uiDoc, categorySleeves);

// Fallback to all sleeves if section box filter returns 0
if (filtered.Count == 0 && categorySleeves.Count > 0)
    sleeves = categorySleeves;
else
    sleeves = filtered;
```

---

### **Phase 2: Grouping**

#### **Step 2.1: Group by (Host Type, MEP Category, Orientation)**
```csharp
// CURRENT: Extract systemType from family name
string systemType = famName.Contains("duct") ? "Duct" 
                  : famName.Contains("pipe") ? "Pipe"
                  : famName.Contains("cabletray") ? "CableTray"
                  : "Unknown";

// NEEDED: Extract systemType from MEP_Category parameter
var categoryParam = sleeve.LookupParameter("MEP_Category");
string systemType = categoryParam?.AsString() ?? "Unknown";

// Grouping Key:
var groups = sleeves.GroupBy(sleeve => new 
{
    hostType = GetHostType(sleeve),         // "Wall", "Floor", "Structural Framing"
    systemType = GetSystemType(sleeve),     // From MEP_Category parameter
    orientation = GetOrientation(sleeve)    // From HostOrientation parameter
});
```

**Grouping Ensures:**
- ✅ Ducts on X-wall don't cluster with ducts on Y-wall
- ✅ Ducts on walls don't cluster with ducts on slabs
- ✅ Ducts don't cluster with pipes (systemType different)
- ✅ Normal ducts don't cluster with insulated ducts (could add insulation flag)

---

### **Phase 3: Cluster Formation (Graph-Based BFS)**

#### **Step 3.1: Precompute Bounding Boxes**
```csharp
var bboxes = new Dictionary<FamilyInstance, BoundingBoxXYZ>();
foreach (var sleeve in groupSleeves)
{
    var bbox = sleeve.get_BoundingBox(null);
    if (bbox != null) bboxes[sleeve] = bbox;
}
```

**Why Precompute?**
- ✅ `get_BoundingBox()` is expensive - call once per sleeve
- ✅ Reused multiple times during neighbor search

#### **Step 3.2: Build Spatial Hash Grid (Performance Optimization)**
```csharp
double cellSize = JoinOpeningsDistance; // e.g., 200mm
var grid = new Dictionary<(int, int, int), List<FamilyInstance>>();

// Hash each sleeve into grid cells based on bbox extents
foreach (var sleeve in groupSleeves)
{
    var bbox = bboxes[sleeve];
    int min_ix = (int)Math.Floor(bbox.Min.X / cellSize);
    int max_ix = (int)Math.Floor(bbox.Max.X / cellSize);
    // ... same for Y, Z
    
    for (int gx = min_ix; gx <= max_ix; gx++)
        for (int gy = min_iy; gy <= max_iy; gy++)
            for (int gz = min_iz; gz <= max_iz; gz++)
            {
                var key = (gx, gy, gz);
                grid[key].Add(sleeve);
            }
}
```

**Why Spatial Grid?**
- ✅ Without grid: O(n²) neighbor search (check every pair)
- ✅ With grid: O(n × k) where k = sleeves per grid cell (typically << n)
- ✅ **Massive performance boost** for 100+ sleeves

#### **Step 3.3: Graph-Based BFS Clustering**
```csharp
var unprocessed = new HashSet<FamilyInstance>(groupSleeves);
var clusters = new List<List<FamilyInstance>>();

while (unprocessed.Count > 0)
{
    // Start new cluster with first unprocessed sleeve
    var start = unprocessed.First();
    var queue = new Queue<FamilyInstance>();
    var cluster = new List<FamilyInstance>();
    
    queue.Enqueue(start);
    unprocessed.Remove(start);
    
    // BFS to find all connected sleeves
    while (queue.Count > 0)
    {
        var current = queue.Dequeue();
        cluster.Add(current);
        
        // Find neighbors within JoinOpeningsDistance
        var neighbors = FindNeighborsUsingGrid(current, grid, bboxes, JoinOpeningsDistance);
        
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

**BFS Guarantees:**
- ✅ All transitively connected sleeves are in same cluster
- ✅ Example: A↔B, B↔C ⇒ A, B, C in same cluster (even if A ↔ C distance > 200mm)
- ✅ No duplicates (each sleeve visited once)

#### **Step 3.4: Neighbor Detection (Bounding Box Overlap)**
```csharp
bool AreNeighbors(FamilyInstance sleeve1, FamilyInstance sleeve2, double tolerance)
{
    var bbox1 = bboxes[sleeve1];
    var bbox2 = bboxes[sleeve2];
    
    // Check 3D bounding box overlap with tolerance
    bool xOverlap = bbox1.Max.X >= bbox2.Min.X - tolerance && 
                    bbox1.Min.X <= bbox2.Max.X + tolerance;
    bool yOverlap = bbox1.Max.Y >= bbox2.Min.Y - tolerance && 
                    bbox1.Min.Y <= bbox2.Max.Y + tolerance;
    bool zOverlap = bbox1.Max.Z >= bbox2.Min.Z - tolerance && 
                    bbox1.Min.Z <= bbox2.Max.Z + tolerance;
    
    return xOverlap && yOverlap && zOverlap;
}
```

**Why Bounding Box Instead of Center-to-Center?**
```
Example: Two rectangular sleeves

Sleeve A: 400×200mm at (0, 0, 0)
Sleeve B: 400×200mm at (250mm, 0, 0)

Center-to-Center Distance: 250mm (> 200mm threshold) ❌ Would NOT cluster

Bounding Box Overlap:
  A.BBox.Max.X = 200mm
  B.BBox.Min.X = 50mm
  Edge-to-Edge Distance = 50mm - 200mm = -150mm (overlapping!) ✅ WILL cluster

Bounding box is MORE ACCURATE for rectangular sleeves!
```

---

### **Phase 4: Cluster Sleeve Placement**

#### **Step 4.1: Calculate Combined Bounding Box**
```csharp
// ClusterBoundingBoxServices.GetClusterBoundingBox(cluster)
BoundingBoxXYZ combinedBbox = new BoundingBoxXYZ();

foreach (var sleeve in cluster)
{
    var bbox = sleeve.get_BoundingBox(null);
    
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

double width = combinedBbox.Max.X - combinedBbox.Min.X;
double height = combinedBbox.Max.Y - combinedBbox.Min.Y;
double depth = combinedBbox.Max.Z - combinedBbox.Min.Z;
XYZ centroid = (combinedBbox.Min + combinedBbox.Max) / 2.0;
```

**Result:** The cluster sleeve will be sized to EXACTLY cover all individual sleeves

#### **Step 4.2: Select Universal Family**
```csharp
// CURRENT: Uses special cluster families
familyName = (hostType == "Wall") ? "ClusterOpeningOnWallX" : "ClusterOpeningOnSlab";

// NEEDED: Use SAME universal families as individual sleeves
bool isCircular = cluster.All(s => s.Symbol.Family.Name.Contains("Circular"));

if (hostType == "Wall" || hostType == "Structural Framing")
    familyName = isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall";
else if (hostType == "Floor")
    familyName = isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab";
```

**Why Same Families?**
- ✅ Simplifies family management (only 4 families total)
- ✅ Cluster sleeve looks identical to individual sleeve (just larger)
- ✅ Parameters are consistent (Width, Height, Depth, MEP_Category)

#### **Step 4.3: Place Cluster Sleeve**
```csharp
var familySymbol = FindFamilySymbol(doc, familyName);
if (!familySymbol.IsActive) familySymbol.Activate();

var level = doc.GetElement(cluster.First().LevelId) as Level;

FamilyInstance clusterSleeve = doc.Create.NewFamilyInstance(
    centroid,           // At center of combined bounding box
    familySymbol,       // Universal family (Rectangular/Circular On Wall/Slab)
    level,
    StructuralType.NonStructural
);
```

#### **Step 4.4: Set Rotation (for Walls only)**
```csharp
if (hostType == "Wall")
{
    double rotationAngle = 0.0;
    if (orientation == "Y")
    {
        // Y-oriented walls need 90° rotation
        rotationAngle = Math.PI / 2;
    }
    
    if (rotationAngle != 0.0)
    {
        Line rotationAxis = Line.CreateBound(centroid, centroid + XYZ.BasisZ);
        ElementTransformUtils.RotateElement(doc, clusterSleeve.Id, rotationAxis, rotationAngle);
    }
}
```

#### **Step 4.5: Set Size Parameters**
```csharp
if (hostType == "Wall" || hostType == "Structural Framing")
{
    // Wall/Framing families are created in "Left" view
    // Parameter mapping: Width=Y, Height=Z, Depth=X
    clusterSleeve.LookupParameter("Width")?.Set(height);   // Y dimension
    clusterSleeve.LookupParameter("Height")?.Set(depth);   // Z dimension
    
    // Depth = host thickness (wall width or framing 'b' parameter)
    double hostThickness = GetHostThickness(cluster.First().Host);
    clusterSleeve.LookupParameter("Depth")?.Set(hostThickness); // X dimension
}
else // Floor
{
    // Floor families use standard mapping
    clusterSleeve.LookupParameter("Width")?.Set(width);
    clusterSleeve.LookupParameter("Height")?.Set(height);
    clusterSleeve.LookupParameter("Depth")?.Set(depth);
}
```

#### **Step 4.6: Set MEP_Category Parameter (CRITICAL NEW STEP)**
```csharp
// Mark cluster with category to prevent future cross-category clustering
var categoryParam = clusterSleeve.LookupParameter("MEP_Category");
if (categoryParam != null)
{
    categoryParam.Set(systemType); // "Ducts", "Pipes", "Duct Accessories", "Cable Trays"
}
```

**Why Critical?**
- ✅ Next time cluster command runs, it filters by MEP_Category
- ✅ Prevents re-clustering cluster sleeves with wrong category
- ✅ Maintains category separation across multiple cluster runs

#### **Step 4.7: Delete Individual Sleeves**
```csharp
foreach (var sleeve in cluster)
{
    doc.Delete(sleeve.Id);
}
```

---

## 🔧 Key Adaptations for Universal Families

### **1. Category Filtering (NEW)**
- **Old:** Family name implied category (DuctOpeningOnWall = Ducts)
- **New:** MUST filter by `MEP_Category` parameter before clustering
- **When:** After collecting all sleeves, before section box filtering

### **2. SystemType Detection (CHANGED)**
```csharp
// OLD:
string systemType = famName.Contains("duct") ? "Duct" : ...;

// NEW:
var categoryParam = sleeve.LookupParameter("MEP_Category");
string systemType = categoryParam?.AsString() ?? "Unknown";
```

### **3. Family Selection (CHANGED)**
```csharp
// OLD:
familyName = "ClusterOpeningOnWallX"; // Special cluster family

// NEW:
bool isCircular = cluster.All(s => s.Symbol.Family.Name.Contains("Circular"));
familyName = (hostType == "Wall") 
    ? (isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall")
    : (isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab");
```

### **4. MEP_Category Marking (NEW)**
- **What:** Set `MEP_Category` parameter on placed cluster sleeve
- **Why:** Ensures future cluster runs respect category boundaries
- **When:** After placing cluster sleeve, before deleting individual sleeves

### **5. Constructor Parameter (NEW)**
```csharp
public RectangularSleeveClusterCommandV2(string targetCategory)
{
    _targetCategory = targetCategory; // "Ducts", "Pipes", "Duct Accessories", "Cable Trays"
}
```

---

## 📊 Execution Flow Comparison

### **OLD (Category-Specific Families)**
```
1. Collect all *OpeningOnWall and *OpeningOnSlab families
2. Section box filter
3. Group by (HostType, FamilyName, Orientation)
   └─ FamilyName implies category
4. Form clusters using BFS + bounding box overlap
5. Place ClusterOpeningOnWallX or ClusterOpeningOnSlab
6. Delete individual sleeves
```

### **NEW (Universal Families)**
```
1. Collect 4 universal families (Rectangular/Circular On Wall/Slab)
2. Filter by MEP_Category parameter = targetCategory ← NEW
3. Section box filter
4. Group by (HostType, MEP_Category, Orientation) ← CHANGED
   └─ MEP_Category from parameter, not family name
5. Form clusters using BFS + bounding box overlap (UNCHANGED)
6. Determine if circular or rectangular from family names
7. Place SAME universal family (just larger size) ← CHANGED
8. Set MEP_Category parameter on cluster sleeve ← NEW
9. Delete individual sleeves
```

---

## ✅ What Stays the Same (No Changes Needed)

1. ✅ **Bounding box precomputation** - Still needed for performance
2. ✅ **Spatial hash grid** - Still optimal for neighbor search
3. ✅ **BFS clustering algorithm** - Graph structure unchanged
4. ✅ **Bounding box overlap detection** - Works for any shape
5. ✅ **Combined bounding box calculation** - `ClusterBoundingBoxServices.GetClusterBoundingBox()`
6. ✅ **Centroid placement** - Cluster placed at center of combined bbox
7. ✅ **Rotation logic** - Y-oriented walls still need 90° rotation
8. ✅ **Parameter mapping** - Wall/Floor dimension mapping unchanged
9. ✅ **Host thickness** - Still read from wall/framing for Depth parameter
10. ✅ **Section box filtering** - Still use `SectionBoxHelper`

---

## 🚨 Critical Success Factors

### **1. MUST Filter by MEP_Category**
```csharp
// ❌ WRONG (will cluster ducts with pipes!)
var sleeves = allSleeves.ToList();

// ✅ CORRECT
var sleeves = allSleeves.Where(s => 
    s.LookupParameter("MEP_Category")?.AsString() == targetCategory
).ToList();
```

### **2. MUST Set MEP_Category on Cluster Sleeve**
```csharp
// ❌ WRONG (future cluster runs will fail to identify category)
// (no parameter set)

// ✅ CORRECT
clusterSleeve.LookupParameter("MEP_Category")?.Set(systemType);
```

### **3. MUST Trigger Separately per Category**
```csharp
// ❌ WRONG (all categories cluster together)
var clusterCmd = new RectangularSleeveClusterCommandV2();

// ✅ CORRECT (each category clusters independently)
var clusterCmd = new RectangularSleeveClusterCommandV2("Ducts");
clusterCmd.Execute(...);

var clusterCmd2 = new RectangularSleeveClusterCommandV2("Pipes");
clusterCmd2.Execute(...);
```

### **4. MUST Use Same Universal Families**
```csharp
// ❌ WRONG (looking for non-existent cluster family)
familyName = "ClusterOpeningOnWallX";

// ✅ CORRECT (use same 4 universal families)
familyName = isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall";
```

---

## 📝 Implementation Checklist

- [x] Add `targetCategory` constructor parameter
- [x] Update sleeve collection to filter by 4 universal family names
- [x] Add MEP_Category filtering after collection
- [x] Update grouping to read systemType from MEP_Category parameter
- [x] Update family selection to use universal families (not cluster-specific)
- [x] Add MEP_Category parameter setting on placed cluster sleeve
- [ ] Test with single category (Ducts only)
- [ ] Test with multiple categories (Ducts, then Pipes)
- [ ] Verify no cross-category clustering occurs
- [ ] Verify cluster sleeves have correct MEP_Category parameter

---

## 🎯 Expected Results

### **Before Clustering (24 individual sleeves)**
```
22 Ducts:
  - 20× RectangularOpeningOnWall (MEP_Category="Ducts")
  - 2× CircularOpeningOnWall (MEP_Category="Ducts")

2 Dampers:
  - 2× RectangularOpeningOnWall (MEP_Category="Duct Accessories")
```

### **After Clustering (assuming 3 clusters form for ducts)**
```
3 Duct Clusters:
  - Cluster A: 8× ducts → 1× RectangularOpeningOnWall (MEP_Category="Ducts", larger size)
  - Cluster B: 10× ducts → 1× RectangularOpeningOnWall (MEP_Category="Ducts", larger size)
  - Cluster C: 4× ducts → 1× RectangularOpeningOnWall (MEP_Category="Ducts", larger size)
  
Individual Sleeves (not clustered):
  - None (all 22 ducts were within 200mm of neighbors)

2 Dampers:
  - 2× RectangularOpeningOnWall (MEP_Category="Duct Accessories") - UNCHANGED
  - Dampers not clustered because cluster command was called with targetCategory="Ducts"
```

---

## 🔍 Next Steps

1. ✅ **Planning Complete** - This document
2. ⏭️ **Code Review** - Verify current implementation matches plan
3. ⏭️ **Testing** - Run cluster command with `targetCategory="Ducts"`
4. ⏭️ **Integration** - Update `SleevePlacementExternalEvent` to trigger cluster after each category
5. ⏭️ **Validation** - Check ClashZone XML for cluster metadata

---

## 📚 References

- **Working Command:** `Commands/RectangularSleeveClusterCommandV2.cs`
- **Bounding Box Service:** `Services/ClusterBoundingBoxServices.cs`
- **Section Box Helper:** `Helpers/SectionBoxHelper.cs`
- **Cluster Configuration:** `Services/ClusterConfigurationManager.cs`

