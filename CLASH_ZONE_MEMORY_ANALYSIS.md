# Clash Zone Memory Analysis

## 📊 **Memory Per Clash Zone - Detailed Breakdown**

### **Base ClashZone Object Size**

#### **1. Core Properties (Fixed Size)**
```
Guid Id                    = 16 bytes
ElementId MepElementId     = 8 bytes (reference)
ElementId StructuralElementId = 8 bytes (reference)
int MepElementIdValue      = 4 bytes
int StructuralElementIdValue = 4 bytes
bool IsResolved            = 1 byte (+ 7 bytes padding = 8 bytes)
bool IsClusterResolved     = 1 byte
bool IsCurrentClash        = 1 byte
bool? MarkedForClustering  = 2 bytes (nullable bool = 1 byte value + 1 byte null flag)
int ClusterSleeveInstanceId = 4 bytes
int AfterClusterSleevePlacedSleeveInstanceId = 4 bytes
int SleeveInstanceId       = 4 bytes
```
**Subtotal: ~60 bytes (excluding padding)**

#### **2. String Properties (Variable Size)**
```
MepElementUniqueId         = ~50-100 bytes (average 75 bytes)
SleeveFamilyName           = ~30-60 bytes (average 45 bytes)
MepElementCategory         = ~10-30 bytes (average 20 bytes)
StructuralElementType      = ~10-30 bytes (average 20 bytes)
DuctShape                  = ~10-20 bytes (average 15 bytes)
SourceDocKey               = ~20-40 bytes (average 30 bytes)
HostDocKey                 = ~20-40 bytes (average 30 bytes)
```
**Subtotal: ~265 bytes (average strings)**

#### **3. Double Properties (All 8 bytes each)**
```
MepElementSize             = 8 bytes
RequiredClearance          = 8 bytes
MepElementWidth            = 8 bytes
MepElementHeight           = 8 bytes
MepElementDiameter         = 8 bytes
SleeveWidth                = 8 bytes
SleeveHeight               = 8 bytes
SleeveDiameter             = 8 bytes
IntersectionPointX         = 8 bytes
IntersectionPointY         = 8 bytes
IntersectionPointZ         = 8 bytes
SleevePlacementPointX      = 8 bytes
SleevePlacementPointY      = 8 bytes
SleevePlacementPointZ      = 8 bytes
SleevePlacementPointActiveDocumentX = 8 bytes
SleevePlacementPointActiveDocumentY = 8 bytes
SleevePlacementPointActiveDocumentZ = 8 bytes
SleeveBoundingBoxMinX     = 8 bytes
SleeveBoundingBoxMinY      = 8 bytes
SleeveBoundingBoxMinZ      = 8 bytes
SleeveBoundingBoxMaxX      = 8 bytes
SleeveBoundingBoxMaxY      = 8 bytes
SleeveBoundingBoxMaxZ      = 8 bytes
ClusterSleeveBoundingBoxMinX = 8 bytes
ClusterSleeveBoundingBoxMinY = 8 bytes
ClusterSleeveBoundingBoxMinZ = 8 bytes
ClusterSleeveBoundingBoxMaxX = 8 bytes
ClusterSleeveBoundingBoxMaxY = 8 bytes
ClusterSleeveBoundingBoxMaxZ = 8 bytes
```
**Subtotal: ~240 bytes (30 doubles × 8 bytes)**

#### **4. Reference Objects (Pointers/References)**
```
XYZ IntersectionPoint      = 24 bytes (3 doubles) + 8 bytes reference = 32 bytes
XYZ SleevePlacementPoint   = 24 bytes (3 doubles) + 8 bytes reference = 32 bytes
XYZ SleevePlacementPointActiveDocument = 24 bytes + 8 bytes reference = 32 bytes
BoundingBoxXYZ ClashBoundingBox = 96 bytes (2 XYZ = 48 bytes each) + 8 bytes reference = 104 bytes
ElementId? ResolvedSleeveId    = 8 bytes (nullable reference)
ElementId? ClusterSleeveId     = 8 bytes (nullable reference)
MepElementSize MepElementSizeData = ~40 bytes (object overhead + properties) + 8 bytes reference = 48 bytes
```
**Subtotal: ~304 bytes**

#### **5. Object Overhead (CLR)**
```
Object header              = 8 bytes (object reference + sync block)
Type pointer               = 8 bytes
Method table pointer       = 8 bytes
Property metadata          = ~16 bytes
```
**Subtotal: ~40 bytes**

**Base ClashZone Object Total: ~909 bytes ≈ ~0.9 KB per object**

---

### **Processing Overhead (During Detection)**

#### **1. Temporary Objects Created Per Clash Zone**
```
BoundingBoxXYZ (for MEP)   = 96 bytes
BoundingBoxXYZ (for Structural) = 96 bytes
XYZ (intersection point)    = 24 bytes
MepElementSize calculation  = ~40 bytes
Geometry calculations       = ~100 bytes (temp variables)
```
**Processing Overhead: ~356 bytes ≈ 0.35 KB per clash zone during detection**

---

### **Collection Overhead**

#### **List<ClashZone> Memory**
```
List<> base object          = 24 bytes
Array capacity overhead    = ~25% capacity (to avoid frequent resizing)
Example: 1000 clash zones in List
- List object: 24 bytes
- Array capacity: 1250 × 8 bytes (references) = 10,000 bytes
- Actual objects: 1000 × 909 bytes = 909,000 bytes
```
**Collection Overhead: ~12.5% extra for List capacity**

#### **Dictionary/HashSet Overhead**
```
Dictionary<string, ClashZone> (when grouping/processing)
- Base: ~24 bytes
- Buckets: ~40 bytes per bucket (depends on load factor)
- Key strings: Additional memory per key
```
**Dictionary Overhead: ~200-400 bytes per entry depending on usage**

---

### **XML Serialization Overhead**

#### **During Save/Load**
```
XML element tags            = ~500 bytes per clash zone (XML markup)
String encoding overhead    = ~10% (UTF-8 encoding)
Serialization temp objects  = ~50 bytes per clash zone
```
**XML Overhead: ~550 bytes ≈ 0.55 KB per clash zone in XML**

---

## 📈 **Total Memory Per Clash Zone**

### **In Memory (Active Processing)**
```
Base ClashZone object           = 909 bytes (~0.9 KB)
Processing overhead (temp)      = 356 bytes (~0.35 KB)
Collection overhead (List)      = ~114 bytes (~12.5% of base)
```
**Total in Memory: ~1,379 bytes ≈ 1.4 KB per clash zone**

### **In XML File**
```
Base data                      = ~1,000 bytes (~1 KB - strings expanded)
XML markup                     = ~500 bytes (~0.5 KB)
```
**Total in XML: ~1,500 bytes ≈ 1.5 KB per clash zone**

---

## 🎯 **Real-World Estimates**

### **Scenario 1: Small Project (500 clash zones)**
```
In Memory:
- 500 × 1.4 KB = 700 KB ≈ 0.7 MB

In XML:
- 500 × 1.5 KB = 750 KB ≈ 0.75 MB
```

### **Scenario 2: Medium Project (5,000 clash zones)**
```
In Memory:
- 5,000 × 1.4 KB = 7,000 KB ≈ 7 MB

In XML:
- 5,000 × 1.5 KB = 7,500 KB ≈ 7.5 MB
```

### **Scenario 3: Large Project (50,000 clash zones)**
```
In Memory:
- 50,000 × 1.4 KB = 70,000 KB ≈ 70 MB

In XML:
- 50,000 × 1.5 KB = 75,000 KB ≈ 75 MB
```

### **Scenario 4: Very Large Project (500,000 clash zones)**
```
In Memory:
- 500,000 × 1.4 KB = 700,000 KB ≈ 700 MB ≈ 0.7 GB

In XML:
- 500,000 × 1.5 KB = 750,000 KB ≈ 750 MB ≈ 0.75 GB
```

---

## 🔍 **Memory During Processing Phases**

### **Phase 1: Intersection Detection**
```
Per Intersection:
- Element references: 16 bytes
- BoundingBoxXYZ (2 objects): 192 bytes
- Geometry calculations: ~200 bytes
- Total: ~408 bytes ≈ 0.4 KB per intersection

Processing 100,000 intersections:
- 100,000 × 0.4 KB = 40 MB (temporary, released after detection)
```

### **Phase 2: Clash Zone Creation**
```
Per Clash Zone Created:
- New ClashZone object: 909 bytes
- Additional temp objects: ~200 bytes
- Total: ~1.1 KB per created clash zone
```

### **Phase 3: Filtering & Optimization**
```
Additional Memory:
- HashSet for filtering: ~100 bytes per clash zone
- Dictionary lookups: ~50 bytes per clash zone
- Total overhead: ~150 bytes per clash zone during filtering
```

### **Phase 4: XML Serialization**
```
During Save:
- Serialization buffer: ~2 KB per clash zone (during serialization only)
- Released after save completes
```

---

## 💾 **Memory Calculation Formula**

### **Per Clash Zone:**
```
Base Memory = 1.4 KB (in memory)
XML Size = 1.5 KB (on disk)

With Processing Overhead:
Detection Phase = 1.4 KB + 0.35 KB = 1.75 KB per clash zone
Filtering Phase = 1.4 KB + 0.15 KB = 1.55 KB per clash zone
Save Phase = 1.4 KB + 2.0 KB (temp) = 3.4 KB peak per clash zone
```

### **For N Clash Zones:**
```
In-Memory Collection: N × 1.4 KB
Peak During Processing: N × 3.4 KB (worst case - during XML save)
XML File Size: N × 1.5 KB
```

---

## 🚀 **Memory Optimization Recommendations**

### **1. String Interning (Can Save ~20%)**
- Cache common strings (category names, file paths)
- **Savings**: ~0.3 KB per clash zone → **0.28 KB per clash zone**

### **2. Lazy Loading (Can Save ~30% during initial load)**
- Don't load all properties until needed
- **Savings**: ~0.4 KB per clash zone initially → **1.0 KB per clash zone**

### **3. XML Compression (Can Save ~60% on disk)**
- Compress XML files
- **Savings**: 1.5 KB → **0.6 KB per clash zone in XML**

### **4. Collection Optimization**
- Pre-size collections to avoid capacity growth
- Use object pools for temporary objects
- **Savings**: ~10-15% collection overhead

---

## 📋 **Summary Table**

| Item | Size | Notes |
|------|------|-------|
| **Base ClashZone Object** | 0.9 KB | Core properties |
| **In Memory (with overhead)** | **1.4 KB** | Including collection overhead |
| **In XML File** | **1.5 KB** | Serialized format |
| **Peak During Processing** | **3.4 KB** | Includes temp buffers |
| **Per 1,000 clash zones** | **1.4 MB** | In memory |
| **Per 10,000 clash zones** | **14 MB** | In memory |
| **Per 100,000 clash zones** | **140 MB** | In memory |
| **Per 1,000,000 clash zones** | **1.4 GB** | In memory |

---

## 🎯 **Practical Limits**

### **64GB System RAM:**
```
Available RAM (70%): ~45 GB
Maximum clash zones: ~32 million (theoretical)
Practical limit: ~1-2 million (accounting for Revit overhead)
```

### **32GB System RAM:**
```
Available RAM (70%): ~22 GB  
Maximum clash zones: ~16 million (theoretical)
Practical limit: ~500,000 - 1 million
```

### **16GB System RAM:**
```
Available RAM (70%): ~11 GB
Maximum clash zones: ~8 million (theoretical)
Practical limit: ~200,000 - 500,000
```

---

**Bottom Line**: Each clash zone uses approximately **~1.4 KB in memory** and **~1.5 KB in XML files**. For 100,000 clash zones, expect **~140 MB in memory** and **~150 MB in XML files**.

