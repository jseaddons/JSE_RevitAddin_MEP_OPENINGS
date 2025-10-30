# Memory Analysis: 72.4 KB per Clash Zone

## Current Situation
- **Theoretical**: 3 KB per clash zone
- **Realistic Estimate**: 15-25 KB per clash zone (with 25-47 KB max expected)
- **Actual Memory**: 72.4 KB per clash zone
- **Multiplier**: 4.8x higher than realistic estimate

## Identified Memory Components per ClashZone

### 1. Base ClashZone Object (~3 KB)
- Strings: MepElementUniqueId, SourceDocKey, HostDocKey, MepElementLevelName, PipeOpeningType, StructuralElementType, etc.
- Doubles: IntersectionPointX/Y/Z, SleeveBoundingBox coordinates, MepElementSize, RequiredClearance, etc.
- Guids: Id
- Booleans: IsResolved, IsClusterResolved, etc.

### 2. Parameter Snapshots
**Whitelist Size:**
- MEP common keys: ~14 parameters
- Host common keys: ~3 parameters
- Learned keys: Variable (could grow unbounded if not limited)

**Memory per SerializableKeyValue:**
- Key: ~25-30 bytes (average parameter name)
- Value: ~20-150 bytes (parameter values can be long)
- Object overhead: ~24 bytes (heap object overhead)
- **Total per parameter: ~70-200 bytes**

**If 17 parameters per element × 2 elements (MEP + Host):**
- ~17 × 150 bytes = 2,550 bytes = ~2.5 KB (worst case)
- But many parameters won't be found, so actual might be ~10-12 found × 100 bytes = ~1-1.2 KB

### 3. XYZ Objects (NOT Serialized - [XmlIgnore])
- IntersectionPoint (XYZ)
- SleevePlacementPoint (XYZ)
- SleevePlacementPointActiveDocument (XYZ)
- WallDirection (XYZ)
- StructuralElementNormal (XYZ)
- MepElementOrientation (XYZ)
- **Note**: These are marked [XmlIgnore] but still exist in memory
- Each XYZ: ~24 bytes (object overhead) + 24 bytes (3 doubles) = ~48 bytes
- 6 XYZ objects = ~288 bytes = ~0.3 KB

### 4. BoundingBoxXYZ Object (NOT Serialized - [XmlIgnore])
- ClashBoundingBox: ~24 bytes (object overhead) + 48 bytes (2 XYZ) = ~72 bytes = 0.07 KB

### 5. ElementId Objects
- MepElementId (ElementId) - struct, ~4 bytes
- StructuralElementId (ElementId) - struct, ~4 bytes
- ResolvedSleeveId (ElementId?) - nullable struct, ~8 bytes
- ClusterSleeveId (ElementId?) - nullable struct, ~8 bytes
- Total: ~24 bytes

### 6. Metadata Dictionary
- Dictionary<string, string> - could grow unbounded if code adds to it
- Overhead: ~100 bytes (empty dictionary)
- If populated: 24 bytes per entry + key/value strings

### 7. List Overhead
- MepParameterValues List: ~24 bytes (empty) + per-item overhead
- HostParameterValues List: ~24 bytes (empty) + per-item overhead
- Total: ~48 bytes

### 8. String Interning & Pooling
- .NET may not intern all strings, causing duplicate string allocations
- Long file paths in SourceDocKey/HostDocKey (could be 200+ bytes each)

## Expected Total
- Base: 3 KB
- Parameters: 1-2.5 KB
- XYZ objects: 0.3 KB
- BoundingBoxXYZ: 0.07 KB
- Other: 0.2 KB
- **Total Expected: ~5-6 KB per clash zone**

## Why We're Seeing 72.4 KB

### Likely Causes:

1. **Learned Parameter Keys Growing Unbounded**
   - If `LearnedParameterKeys` grows to 50+ parameters, and each is captured, we get ~50 × 150 bytes = 7.5 KB per element
   - For 2 elements = 15 KB per clash zone
   - Still doesn't explain 72 KB...

2. **XML Serialization Overhead**
   - If ClashZone objects are being serialized/deserialized repeatedly, XML serialization creates temporary objects

3. **Revit API Object References**
   - If any Revit API objects (Element, Document, Parameter objects) are accidentally stored or referenced, they hold large caches

4. **String Duplication**
   - If strings aren't being interned properly, duplicate strings for parameter names/values waste memory

5. **.NET GC Overhead**
   - Large object heap fragmentation
   - Many small allocations causing GC pressure

6. **Debug Logging Still Active**
   - Even though we disabled file writes, if string formatting/concatenation is still happening, those strings are allocated in memory

## Recommendations

### Immediate Fixes:

1. **Limit Learned Parameter Keys**
   ```csharp
   // Cap at 20 learned parameters max
   if (storage?.LearnedParameterKeys?.Count > 20)
   {
       storage.LearnedParameterKeys = storage.LearnedParameterKeys.Take(20).ToList();
   }
   ```

2. **Clear XYZ Objects After Capture**
   ```csharp
   // After serialization, clear non-essential XYZ objects
   clashZone.IntersectionPoint = null;
   clashZone.WallDirection = null;
   // etc.
   ```

3. **Clear Metadata Dictionary if Not Used**
   ```csharp
   if (clashZone.Metadata != null && clashZone.Metadata.Count == 0)
   {
       clashZone.Metadata = null;
   }
   ```

4. **Check for Accidental Revit API References**
   - Ensure no Element, Document, or Parameter objects are stored in ClashZone

5. **Profile Actual Memory Allocation**
   - Use CLR Profiler or PerfView to see actual allocations per ClashZone
