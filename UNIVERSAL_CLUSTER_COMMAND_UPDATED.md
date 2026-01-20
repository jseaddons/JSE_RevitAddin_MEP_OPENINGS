# ✅ UNIVERSAL CLUSTER COMMAND - UPDATED FOR UNIVERSAL FAMILIES

## 🎯 Objective

Cluster sleeves for each MEP category **separately** using the 4 universal families, avoiding cross-category clustering.

---

## 🔧 CRITICAL CHANGES for Universal Families

### **Problem: Cross-Category Clustering Risk**

**Old System:**
- Different families per category (DuctOpeningOnWall, PipeOpeningOnWall, etc.)
- Cluster command could distinguish sleeves by family name
- Ducts wouldn't cluster with pipes automatically

**New System (Current):**
- **Same 4 universal families for ALL categories:**
  - `RectangularOpeningOnWall`
  - `CircularOpeningOnWall`
  - `RectangularOpeningOnSlab`
  - `CircularOpeningOnSlab`
- **Risk:** If we cluster all sleeves together, ducts WILL cluster with pipes/dampers!

### **Solution: Category-Filtered Clustering**

**Cluster command MUST:**
1. ✅ Run **separately for each category** (triggered after individual placement)
2. ✅ Filter sleeves by `MEP_Category` parameter before clustering
3. ✅ Use merge distance from `ClusterConfigurationManager.Instance.GetJoinOpeningsDistance()`

---

## 📋 Implementation Steps

### **Step 1: Collect Sleeves by Category**

```csharp
public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
{
    UIApplication uiapp = commandData.Application;
    UIDocument uidoc = uiapp.ActiveUIDocument;
    Document doc = uidoc.Document;
    
    // CRITICAL: Get target category (passed from SleevePlacementExternalEvent)
    string targetCategory = _targetCategory; // "Ducts", "Pipes", "Duct Accessories", or "Cable Trays"
    
    DebugLogger.Info($"[ClusterCommand] Starting cluster for category: {targetCategory}");
    
    // Step 1: Collect ALL sleeves using 4 universal families
    var allSleeves = new FilteredElementCollector(doc)
        .OfClass(typeof(FamilyInstance))
        .Cast<FamilyInstance>()
        .Where(fi => 
        {
            var famName = fi.Symbol?.Family?.Name ?? string.Empty;
            
            // Match ONLY the 4 universal family names
            return famName == "RectangularOpeningOnWall" ||
                   famName == "CircularOpeningOnWall" ||
                   famName == "RectangularOpeningOnSlab" ||
                   famName == "CircularOpeningOnSlab";
        })
        .ToList();
    
    DebugLogger.Info($"[ClusterCommand] Found {allSleeves.Count} total sleeves (all categories)");
    
    // Step 2: Filter by MEP_Category parameter to get ONLY this category's sleeves
    var categorySleeves = allSleeves
        .Where(sleeve => 
        {
            var categoryParam = sleeve.LookupParameter("MEP_Category");
            string sleeveCategory = categoryParam?.AsString() ?? "";
            return sleeveCategory == targetCategory;
        })
        .ToList();
    
    DebugLogger.Info($"[ClusterCommand] Filtered to {categorySleeves.Count} sleeves for category '{targetCategory}'");
    
    if (categorySleeves.Count == 0)
    {
        DebugLogger.Warning($"[ClusterCommand] No sleeves found for category '{targetCategory}' - nothing to cluster");
        return Result.Succeeded;
    }
    
    // Step 3: Get merge distance from configuration
    double mergeDistanceMm = ClusterConfigurationManager.Instance.GetJoinOpeningsDistance();
    double mergeDistanceFt = UnitUtils.ConvertToInternalUnits(mergeDistanceMm, UnitTypeId.Millimeters);
    
    DebugLogger.Info($"[ClusterCommand] Using merge distance: {mergeDistanceMm}mm ({mergeDistanceFt:F4}ft)");
    
    // Continue with existing clustering logic...
    return ClusterSleeves(doc, categorySleeves, mergeDistanceFt, targetCategory);
}
```

### **Step 2: Group Sleeves (Existing Algorithm - No Changes Needed)**

The existing bounding box overlap algorithm already works for any shape:

```csharp
// Group sleeves by (Host, Level, Normal/Insulated)
var grouped = categorySleeves
    .GroupBy(sleeve => new 
    {
        hostId = GetHostElementId(sleeve),
        levelId = GetLevelId(sleeve),
        isInsulated = IsInsulated(sleeve) // From "Insulated" parameter
    })
    .ToList();

foreach (var group in grouped)
{
    // Find clusters within this group using bounding box overlap
    var clusters = FindClusters(group.ToList(), mergeDistanceFt);
    
    // Create cluster sleeves
    foreach (var cluster in clusters)
    {
        if (cluster.Count > 1) // Only cluster if 2+ sleeves
        {
            CreateClusterSleeve(doc, cluster, targetCategory);
        }
    }
}
```

### **Step 3: Bounding Box Overlap (Existing - Works for All Shapes)**

```csharp
private List<List<FamilyInstance>> FindClusters(List<FamilyInstance> sleeves, double tolerance)
{
    var clusters = new List<List<FamilyInstance>>();
    var visited = new HashSet<ElementId>();
    
    foreach (var sleeve in sleeves)
    {
        if (visited.Contains(sleeve.Id)) continue;
        
        var cluster = new List<FamilyInstance> { sleeve };
        visited.Add(sleeve.Id);
        
        // Find all neighbors within tolerance (bounding box overlap)
        var neighbors = FindNeighbors(sleeve, sleeves, tolerance, visited);
        cluster.AddRange(neighbors);
        
        if (cluster.Count > 1)
        {
            clusters.Add(cluster);
        }
    }
    
    return clusters;
}

private List<FamilyInstance> FindNeighbors(FamilyInstance sleeve, List<FamilyInstance> allSleeves, 
                                           double tolerance, HashSet<ElementId> visited)
{
    var neighbors = new List<FamilyInstance>();
    var bbox1 = sleeve.get_BoundingBox(null);
    
    foreach (var other in allSleeves)
    {
        if (other.Id == sleeve.Id || visited.Contains(other.Id)) continue;
        
        var bbox2 = other.get_BoundingBox(null);
        
        // Check bounding box overlap with tolerance
        bool xOverlap = bbox1.Max.X >= bbox2.Min.X - tolerance && 
                        bbox1.Min.X <= bbox2.Max.X + tolerance;
        bool yOverlap = bbox1.Max.Y >= bbox2.Min.Y - tolerance && 
                        bbox1.Min.Y <= bbox2.Max.Y + tolerance;
        bool zOverlap = bbox1.Max.Z >= bbox2.Min.Z - tolerance && 
                        bbox1.Min.Z <= bbox2.Max.Z + tolerance;
        
        if (xOverlap && yOverlap && zOverlap)
        {
            neighbors.Add(other);
            visited.Add(other.Id);
        }
    }
    
    return neighbors;
}
```

### **Step 4: Create Cluster Sleeve (Updated for Universal Families)**

```csharp
private void CreateClusterSleeve(Document doc, List<FamilyInstance> cluster, string targetCategory)
{
    // Calculate bounding box of all sleeves in cluster
    BoundingBoxXYZ clusterBbox = CalculateClusterBoundingBox(cluster);
    
    // Determine if cluster is circular or rectangular
    bool isCircular = cluster.All(s => IsCircular(s));
    
    // Get host element from first sleeve
    var firstSleeve = cluster.First();
    var hostId = GetHostElementId(firstSleeve);
    var host = doc.GetElement(hostId);
    
    // Determine host type (Wall, Floor, or Structural Framing)
    bool isWallOrFraming = host.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Walls ||
                           host.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming;
    
    // Select correct universal family
    string targetFamilyName = isWallOrFraming 
        ? (isCircular ? "CircularOpeningOnWall" : "RectangularOpeningOnWall")
        : (isCircular ? "CircularOpeningOnSlab" : "RectangularOpeningOnSlab");
    
    // Find family symbol
    var familySymbol = new FilteredElementCollector(doc)
        .OfClass(typeof(FamilySymbol))
        .Cast<FamilySymbol>()
        .FirstOrDefault(fs => fs.Family.Name == targetFamilyName);
    
    if (familySymbol == null)
    {
        DebugLogger.Error($"[ClusterCommand] Family '{targetFamilyName}' not found");
        return;
    }
    
    if (!familySymbol.IsActive)
        familySymbol.Activate();
    
    // Create cluster sleeve at centroid
    XYZ centroid = CalculateCentroid(cluster);
    Level level = doc.GetElement(firstSleeve.LevelId) as Level;
    
    FamilyInstance clusterSleeve = doc.Create.NewFamilyInstance(
        centroid, 
        familySymbol, 
        level, 
        StructuralType.NonStructural);
    
    // Set parameters
    SetClusterSleeveParameters(clusterSleeve, cluster, clusterBbox, targetCategory);
    
    // Mark cluster in ClashZone XML
    MarkClusterInClashZones(cluster, clusterSleeve, targetCategory);
    
    // Delete individual sleeves
    foreach (var sleeve in cluster)
    {
        doc.Delete(sleeve.Id);
    }
    
    DebugLogger.Info($"[ClusterCommand] Created cluster sleeve for {cluster.Count} {targetCategory} sleeves");
}
```

### **Step 5: Set Cluster Sleeve Parameters**

```csharp
private void SetClusterSleeveParameters(FamilyInstance clusterSleeve, List<FamilyInstance> cluster, 
                                        BoundingBoxXYZ bbox, string targetCategory)
{
    // Set MEP_Category to identify which category this cluster belongs to
    var categoryParam = clusterSleeve.LookupParameter("MEP_Category");
    if (categoryParam != null && !categoryParam.IsReadOnly)
    {
        categoryParam.Set(targetCategory); // "Ducts", "Pipes", "Duct Accessories", or "Cable Trays"
    }
    
    // Set size parameters based on bounding box
    double width = bbox.Max.X - bbox.Min.X;
    double height = bbox.Max.Y - bbox.Min.Y;
    
    if (IsCircular(cluster.First()))
    {
        // Circular cluster
        double diameter = Math.Max(width, height);
        var diamParam = clusterSleeve.LookupParameter("Diameter") ?? clusterSleeve.LookupParameter("d");
        if (diamParam != null && !diamParam.IsReadOnly)
        {
            diamParam.Set(diameter);
        }
    }
    else
    {
        // Rectangular cluster
        var widthParam = clusterSleeve.LookupParameter("Width") ?? clusterSleeve.LookupParameter("a");
        if (widthParam != null && !widthParam.IsReadOnly)
        {
            widthParam.Set(width);
        }
        
        var heightParam = clusterSleeve.LookupParameter("Height") ?? clusterSleeve.LookupParameter("h");
        if (heightParam != null && !heightParam.IsReadOnly)
        {
            heightParam.Set(height);
        }
    }
    
    // Set host thickness (depth)
    var firstSleeve = cluster.First();
    var hostId = GetHostElementId(firstSleeve);
    var host = doc.GetElement(hostId);
    double thickness = GetHostThickness(host);
    
    var depthParam = clusterSleeve.LookupParameter("Wall Width") ?? 
                     clusterSleeve.LookupParameter("Depth") ?? 
                     clusterSleeve.LookupParameter("b");
    if (depthParam != null && !depthParam.IsReadOnly)
    {
        depthParam.Set(thickness);
    }
    
    // Set cluster metadata parameters
    var clusterCountParam = clusterSleeve.LookupParameter("Cluster_Count");
    if (clusterCountParam != null && !clusterCountParam.IsReadOnly)
    {
        clusterCountParam.Set(cluster.Count);
    }
    
    // Collect MEP element IDs
    var mepIds = cluster
        .Select(s => s.LookupParameter("MEP_Element_UniqueId")?.AsString())
        .Where(id => !string.IsNullOrEmpty(id))
        .Distinct()
        .ToList();
    
    var mepIdsParam = clusterSleeve.LookupParameter("Clustered_MEP_Elements");
    if (mepIdsParam != null && !mepIdsParam.IsReadOnly)
    {
        mepIdsParam.Set(string.Join(";", mepIds));
    }
}
```

---

## 📋 Execution Flow in SleevePlacementExternalEvent

```csharp
public void Execute(UIApplication uiApp)
{
    foreach (var category in selectedCategories) // "Ducts", "Pipes", "Duct Accessories", "Cable Trays"
    {
        // 1. Load clash zones for this category
        var clashZones = LoadClashZonesForCategory(category);
        
        // 2. Place individual sleeves
        var placementCommand = new UniversalSleevePlacementCommand(category, clashZones);
        placementCommand.Execute(commandData, ref message, elements);
        
        // 3. IMMEDIATELY cluster this category's sleeves
        var clusterCommand = new UniversalClusterCommand(category);
        clusterCommand.Execute(commandData, ref message, elements);
        
        DebugLogger.Info($"[SleevePlacement] Completed placement and clustering for {category}");
    }
}
```

---

## ✅ Key Benefits

1. ✅ **Prevents cross-category clustering** - Ducts won't cluster with pipes/dampers
2. ✅ **Same 4 universal families** - Consistent family usage across individual and cluster sleeves
3. ✅ **Category-specific clustering** - Each category clusters independently
4. ✅ **Configurable merge distance** - Reads from `ClusterConfigurationManager.Instance`
5. ✅ **Existing algorithm reused** - Bounding box overlap works for any shape (circular/rectangular)

---

## 🚨 CRITICAL Requirements

1. ✅ **MUST filter by `MEP_Category` parameter** before clustering
2. ✅ **MUST run cluster command separately for each category**
3. ✅ **MUST use same 4 universal family names** (no category-specific families)
4. ✅ **MUST read merge distance** from `ClusterConfigurationManager.Instance.GetJoinOpeningsDistance()`
5. ✅ **MUST mark cluster in ClashZone XML** to track which sleeves were clustered

---

## 📝 Next Steps

1. ✅ Update `RectangularSleeveClusterCommandV2` to accept `targetCategory` parameter
2. ✅ Add category filtering logic to sleeve collection
3. ✅ Update `SleevePlacementExternalEvent` to trigger cluster command after each category placement
4. ✅ Test with ducts and dampers to ensure no cross-category clustering
5. ✅ Validate cluster metadata is saved correctly to XML

