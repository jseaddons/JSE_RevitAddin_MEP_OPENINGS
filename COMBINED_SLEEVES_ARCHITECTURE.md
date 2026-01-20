# Combined Sleeves Architecture - Future Implementation Plan

## Overview

This document outlines the proposed architecture for **Combined Sleeves** - a feature that will cluster category-specific cluster sleeves together to create multi-category combined openings (e.g., Ducts + Pipes in a single larger sleeve).

**Status**: 📋 **Documentation Only** - Implementation deferred for future consideration

---

## Current Architecture (As-Is)

### Current Flow:
```
For each filter:
  ├─ Phase 1: Place individual sleeves per category
  │   └─ ExecuteUniversalSleevePlacement(filter) → Places individual sleeves
  │
  └─ Phase 2: Cluster per category (creates category-specific cluster sleeves)
      └─ ExecuteClusteringForCategory(filter) → Clusters individual sleeves within same category
```

### Current Limitations:
- Clustering only happens **within the same category** (e.g., Ducts cluster with Ducts only)
- **No cross-category clustering** (Ducts and Pipes cannot cluster together)
- Each category processes independently

---

## Proposed Architecture (Combined Sleeves)

### New Flow (3-Phase Approach):
```
For each filter:
  ├─ Phase 1: Place individual sleeves per category ✅ (EXISTING - Keep as-is)
  │   └─ ExecuteUniversalSleevePlacement(filter)
  │
  ├─ Phase 2: Cluster per category ✅ (EXISTING - Keep as-is)
  │   └─ ExecuteClusteringForCategory(filter)
  │       └─ Creates category-specific cluster sleeves (e.g., Ducts cluster, Pipes cluster)
  │
  └─ Phase 3: NEW - Cluster ALL cluster sleeves together (Multi-category combined clusters)
      └─ ExecuteCombinedClusteringForAllCategories(filters)
          └─ Combines category-specific clusters into multi-category combined sleeves
              └─ Example: Ducts cluster + Pipes cluster → Combined Ducts+Pipes sleeve
```

### Key Concept:
**Incremental Clustering Approach**:
1. **Individual sleeves** → Cluster within same category → **Category-specific cluster sleeves**
2. **Category-specific cluster sleeves** → Cluster across categories → **Combined multi-category sleeves**

---

## Technical Architecture

### Phase 3: Combined Clustering Process

#### Step 1: Load Cluster Sleeves from All Categories
```csharp
private List<ClusterSleeveInfo> LoadClusterSleevesFromAllCategories(List<OpeningFilter> filters)
{
    var allClusterSleeves = new List<ClusterSleeveInfo>();
    
    foreach (var filter in filters)
    {
        var xmlFilePath = GetXmlFilePathForFilter(filter);
        var clusterSleeves = LoadClusterSleevesFromXml(xmlFilePath, filter.Category);
        allClusterSleeves.AddRange(clusterSleeves);
    }
    
    return allClusterSleeves;
}
```

**Data Source**: Filter XML files (`{FilterName}_{category}.xml`)
- Extract clash zones with `ClusterSleeveInstanceId > 0` and `IsClusterResolved == true`
- Group by `ClusterSleeveInstanceId` to get unique cluster sleeves
- Extract bounding box from `ClusterSleeveBoundingBoxMinX/Y/Z` and `MaxX/Y/Z`

#### Step 2: Group by Host Type + Orientation
```csharp
var groups = allClusterSleeves.GroupBy(s => new { 
    HostType = s.HostType,      // Wall, Floor, Structural Framing
    Orientation = s.Orientation  // X-Wall, Y-Wall, Floor
});
```

**Grouping Logic**:
- **X-Wall/Framing**: All cluster sleeves on X-oriented walls/framing
- **Y-Wall/Framing**: All cluster sleeves on Y-oriented walls/framing
- **Floor**: All cluster sleeves on floors

#### Step 3: Cluster Within Each Group (Multi-Category)
```csharp
foreach (var group in groups)
{
    var clusters = FormClusters(group.ToList(), toleranceDist);
    
    foreach (var cluster in clusters)
    {
        var categoriesInCluster = cluster.Select(s => s.Category).Distinct().ToList();
        
        if (categoriesInCluster.Count > 1)
        {
            // Multi-category cluster → Create combined sleeve
            CreateCombinedClusterSleeve(cluster, categoriesInCluster);
        }
        // If single category → Keep as-is (no change needed)
    }
}
```

**Clustering Logic**:
- Reuse existing `FormClusters` method from `UniversalClusterService`
- Treat cluster sleeves like individual sleeves (same bounding box overlap algorithm)
- Check spatial proximity using `ClusterSleeveBoundingBox` coordinates

#### Step 4: Create Combined Cluster Sleeve
```csharp
private FamilyInstance CreateCombinedClusterSleeve(
    List<ClusterSleeveInfo> clusterSleeves, 
    List<string> categoriesInCluster)
{
    // 1. Calculate combined bounding box (union of all cluster sleeve bounding boxes)
    var combinedBbox = CalculateCombinedBoundingBox(clusterSleeves);
    
    // 2. Determine host type and orientation (from first cluster sleeve)
    var hostType = clusterSleeves[0].HostType;
    var orientation = clusterSleeves[0].Orientation;
    
    // 3. Calculate combined dimensions (max width/height across all categories)
    var combinedWidth = clusterSleeves.Max(s => s.Width);
    var combinedHeight = clusterSleeves.Max(s => s.Height);
    
    // 4. Select appropriate family symbol (based on host type)
    var familySymbol = GetFamilySymbolForHostType(hostType);
    
    // 5. Create combined cluster sleeve instance
    var combinedSleeve = _doc.Create.NewFamilyInstance(
        combinedBbox.Center,
        familySymbol,
        GetNearestLevel(combinedBbox.Center),
        StructuralType.NonStructural);
    
    // 6. Set combined dimensions
    SetCombinedDimensions(combinedSleeve, combinedWidth, combinedHeight, categoriesInCluster);
    
    // 7. Set metadata (categories involved)
    SetCombinedClusterMetadata(combinedSleeve, categoriesInCluster);
    
    return combinedSleeve;
}
```

#### Step 5: Update Flags and Cleanup
```csharp
// 1. Delete category-specific cluster sleeves
foreach (var clusterSleeve in clusterSleeves)
{
    var clusterSleeveElement = _doc.GetElement(new ElementId(clusterSleeve.ClusterSleeveInstanceId));
    _doc.Delete(clusterSleeveElement.Id);
}

// 2. Update Global XML flags for all categories involved
foreach (var category in categoriesInCluster)
{
    UpdateGlobalXmlFlagsForCombinedCluster(category, clusterSleeves, combinedSleeve);
}

// 3. Update Filter XML with combined cluster info
UpdateFilterXmlWithCombinedCluster(clusterSleeves, combinedSleeve, categoriesInCluster);
```

---

## Critical Considerations

### ⚠️ **Individual Sleeves May Also Be Affected**

**Important**: Combined clustering may affect **individual sleeves** (not just cluster sleeves):

#### Scenario 1: Individual Sleeve Near Combined Cluster
- **Issue**: Individual sleeve from Category A might be close to a combined cluster sleeve (Category B + Category C)
- **Action**: Individual sleeve should be **incorporated into combined cluster** if within tolerance
- **Impact**: Individual sleeve becomes part of multi-category combined cluster

#### Scenario 2: Individual Sleeve Between Two Cluster Sleeves
- **Issue**: Individual sleeve might be between two category-specific cluster sleeves that are being combined
- **Action**: Individual sleeve should be **absorbed into combined cluster** during combination
- **Impact**: Individual sleeve flags change: `IsResolved=false`, `IsClusterResolved=true`

#### Implementation Consideration:
```csharp
// During combined clustering, also check for nearby individual sleeves
var nearbyIndividualSleeves = FindIndividualSleevesNearClusterSleeves(clusterSleeves, toleranceDist);

if (nearbyIndividualSleeves.Count > 0)
{
    // Include individual sleeves in combined cluster
    allSleevesInCombinedCluster.AddRange(nearbyIndividualSleeves);
    
    // Update individual sleeve flags
    foreach (var individualSleeve in nearbyIndividualSleeves)
    {
        // Mark as cluster-resolved (individual sleeve becomes part of combined cluster)
        individualSleeve.IsResolved = false;
        individualSleeve.IsClusterResolved = true;
        individualSleeve.ClusterSleeveInstanceId = combinedSleeve.Id.IntegerValue;
        individualSleeve.CombinedClusterSleeveInstanceId = combinedSleeve.Id.IntegerValue; // NEW field
    }
}
```

### ⚠️ **Flag Management Complexity**

**Challenge**: Combined clusters span multiple categories, but flags are stored **per category** in Global XML.

**Solution**: Update flags in **all categories** involved:
```csharp
private void UpdateGlobalXmlFlagsForCombinedCluster(
    string category,
    List<ClusterSleeveInfo> clusterSleeves,
    FamilyInstance combinedSleeve)
{
    var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
    
    // Find all clash zones that were part of this category's cluster sleeves
    var categoryClashZones = GetClashZonesInClusterSleeves(clusterSleeves, category);
    
    foreach (var clashZone in categoryClashZones)
    {
        var entry = globalIndex.Entries.FirstOrDefault(e => e.Id == clashZone.Id.ToString());
        if (entry != null)
        {
            // Update flags for combined cluster
            entry.IsClusterResolved = true;
            entry.ClusterSleeveInstanceId = combinedSleeve.Id.IntegerValue;
            entry.CombinedClusterSleeveInstanceId = combinedSleeve.Id.IntegerValue; // NEW field
            entry.CategoriesInCombinedCluster = string.Join(",", categoriesInCluster); // NEW field
        }
    }
    
    GlobalIndexService.Save(_document, globalIndex);
}
```

### ⚠️ **New Data Fields Required**

**ClashZone Model**:
```csharp
// NEW: Combined cluster sleeve instance ID
public int CombinedClusterSleeveInstanceId { get; set; } = -1;

// NEW: Categories involved in combined cluster (e.g., "Ducts,Pipes")
public string CategoriesInCombinedCluster { get; set; } = string.Empty;

// NEW: Combined cluster bounding box (for cleanup detection)
public double CombinedClusterSleeveBoundingBoxMinX { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMinY { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMinZ { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMaxX { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMaxY { get; set; } = 0.0;
public double CombinedClusterSleeveBoundingBoxMaxZ { get; set; } = 0.0;
```

**GlobalIndexService.CategoryGlobalIndexEntry**:
```csharp
// NEW: Combined cluster sleeve instance ID
[XmlAttribute("CombinedClusterSleeveInstanceId")]
public int CombinedClusterSleeveInstanceId { get; set; } = -1;

// NEW: Categories involved in combined cluster
[XmlAttribute("CategoriesInCombinedCluster")]
public string CategoriesInCombinedCluster { get; set; } = string.Empty;
```

---

## Implementation Steps (When Ready)

### Step 1: Data Model Updates
- [ ] Add `CombinedClusterSleeveInstanceId` to `ClashZone` model
- [ ] Add `CategoriesInCombinedCluster` to `ClashZone` model
- [ ] Add combined cluster bounding box fields to `ClashZone` model
- [ ] Add combined cluster fields to `CategoryGlobalIndexEntry` model

### Step 2: Helper Methods
- [ ] Create `LoadClusterSleevesFromXml()` method
- [ ] Create `GetClusterBoundingBoxFromClashZone()` method
- [ ] Create `FindIndividualSleevesNearClusterSleeves()` method
- [ ] Create `CalculateCombinedBoundingBox()` method

### Step 3: Combined Clustering Logic
- [ ] Create `ExecuteCombinedClusteringForAllCategories()` method in `OpeningCommandOrchestrator`
- [ ] Create `ClusterCombinedSleeves()` method
- [ ] Create `CreateCombinedClusterSleeve()` method
- [ ] Reuse existing `FormClusters()` method for spatial clustering

### Step 4: Flag Management
- [ ] Create `UpdateGlobalXmlFlagsForCombinedCluster()` method
- [ ] Update `FlagManager` to handle combined cluster flags
- [ ] Update flag sync logic to include combined cluster flags

### Step 5: Cleanup Logic
- [ ] Delete category-specific cluster sleeves after creating combined cluster
- [ ] Update individual sleeves that become part of combined cluster
- [ ] Update Filter XML with combined cluster information

### Step 6: User Feedback
- [ ] Add progress messages for combined clustering phase
- [ ] Log combined cluster creation (categories involved)
- [ ] Show summary of combined clusters created

---

## Benefits

1. **Better Optimization**: Cross-category clustering (Ducts + Pipes = single larger opening)
2. **Handles New Linked Files**: All sleeves placed before combined clustering
3. **Incremental Approach**: Doesn't break existing architecture
4. **Reuses Existing Logic**: Same clustering algorithm for cluster sleeves
5. **Clear Separation**: Per-category clustering vs combined clustering

---

## Risks and Challenges

1. **Complexity**: Flag management across multiple categories
2. **Individual Sleeves**: Must handle individual sleeves being absorbed into combined clusters
3. **Backward Compatibility**: Existing single-category clusters must still work
4. **Performance**: Processing all cluster sleeves together may be slower
5. **Testing**: Must test with multiple categories and linked files

---

## Future Considerations

### Option 1: Skip Individual Sleeves
- **Approach**: Only combine cluster sleeves, ignore individual sleeves
- **Pros**: Simpler implementation
- **Cons**: May miss optimization opportunities

### Option 2: Include Individual Sleeves (Recommended)
- **Approach**: Check for nearby individual sleeves during combined clustering
- **Pros**: Maximum optimization
- **Cons**: More complex flag management

### Option 3: Two-Pass Combined Clustering
- **Approach**: 
  1. First pass: Combine cluster sleeves only
  2. Second pass: Include individual sleeves near combined clusters
- **Pros**: Balanced approach
- **Cons**: More processing time

---

## Example Scenario

### Setup:
- **Filter**: "Main Building"
- **Categories**: Ducts, Pipes, Cable Trays
- **Host**: X-Wall

### Phase 1: Individual Sleeves
- 10 Ducts sleeves placed
- 8 Pipes sleeves placed
- 5 Cable Trays sleeves placed

### Phase 2: Category-Specific Clusters
- Ducts: 3 individual sleeves → 1 Ducts cluster sleeve
- Pipes: 2 individual sleeves → 1 Pipes cluster sleeve
- Cable Trays: All individual sleeves remain (no clustering)

### Phase 3: Combined Clustering (NEW)
- **Input**: 1 Ducts cluster sleeve + 1 Pipes cluster sleeve + 5 individual Cable Trays sleeves
- **Check**: Ducts cluster and Pipes cluster are within tolerance
- **Action**: Create combined Ducts+Pipes cluster sleeve
- **Result**: 
  - Delete: Ducts cluster sleeve, Pipes cluster sleeve
  - Create: 1 combined Ducts+Pipes cluster sleeve
  - Individual Cable Trays sleeves: Check if near combined cluster → If yes, incorporate

### Final Result:
- 1 combined Ducts+Pipes cluster sleeve
- 5 individual Cable Trays sleeves (or incorporated into combined cluster if nearby)

---

## Conclusion

The combined sleeves architecture provides a **powerful optimization** by allowing cross-category clustering. However, it requires careful consideration of:

1. **Individual sleeves** that may be affected
2. **Flag management** across multiple categories
3. **Data model updates** to track combined clusters
4. **Backward compatibility** with existing single-category clusters

**Recommendation**: Implement in phases:
1. **Phase 1**: Combine cluster sleeves only (simpler)
2. **Phase 2**: Add individual sleeve incorporation (more complex)

This incremental approach allows testing and validation at each phase.

---

**Document Created**: 2024
**Status**: 📋 Planning Phase - Implementation Deferred
**Author**: AI Assistant
**Review Status**: Pending User Approval

