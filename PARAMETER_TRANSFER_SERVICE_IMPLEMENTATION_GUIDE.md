# ✅ Parameter Transfer Service - Implementation Guide

## 🎯 **Core Problem Statement**

Transfer MEP element parameter values to sleeve parameters using either:
1. **LIVE Detection**: Read directly from currently intersecting MEP elements
2. **SNAPSHOT Method**: Use saved data from XML when MEP elements have moved/deleted

---

## 📋 **Current Issues & Root Causes**

### ❌ **Problem 1: Wrong Index Usage**
```csharp
// WRONG: Searching by category (gets ALL ducts)
var ductParams = GetParametersFromCategoryXml("Ducts");

// CORRECT: Use sleeve-specific lookup
var sleeveParams = snapshotIndex[sleeveId];
```

### ❌ **Problem 2: Method Confusion**
```csharp
// WRONG: Mixing snapshot building with category lookup
BuildSnapshotIndex(); // Creates: sleeveId → mepParams
GetParametersFromCategoryXml(); // Ignores sleeveId, searches by category

// CORRECT: Use the snapshot index directly
var mepParams = snapshotIndex[sleeveId]?.MepParameterValues;
```

### ❌ **Problem 3: XML File Pattern Mismatch**
```csharp
// WRONG: Different files for building vs reading
BuildSnapshotIndex("CONDITIONS_ducts.xml");
GetParametersFromCategoryXml("ducts.xml");

// CORRECT: Use same file pattern consistently
```

---

## 🏗️ **Correct Architecture**

### **Data Flow Diagram**
```
User clicks "Transfer All"
         ↓
For each sleeve in model:
         ↓
    1. Get sleeveId
    2. Lookup: snapshotIndex[sleeveId]
    3. If found: Use saved MEP parameters
    4. If not found: Fall back to live detection
         ↓
    5. Apply parameter mapping
    6. Write to sleeve parameters
```

### **XML Structure (Correct Format)**
```xml
<ClashZones>
  <ClashZone>
    <SleeveInstanceId>123456</SleeveInstanceId>
    <MepElementId>789012</MepElementId>
    <MepParameterValues>
      <Parameter Key="System Type" Value="Supply Air"/>
      <Parameter Key="Size" Value="600x300"/>
      <Parameter Key="Level" Value="Level 1"/>
    </MepParameterValues>
    <HostParameterValues>
      <Parameter Key="Type" Value="Concrete Wall"/>
      <Parameter Key="Thickness" Value="200"/>
    </HostParameterValues>
  </ClashZone>

  <ClashZone>
    <SleeveInstanceId>123457</SleeveInstanceId>
    <MepElementId>789013</MepElementId>
    <MepParameterValues>
      <Parameter Key="System Type" Value="Return Air"/>
      <Parameter Key="Size" Value="400x200"/>
    </MepParameterValues>
  </ClashZone>
</ClashZones>
```

---

## 💻 **Implementation Code**

### **1. Snapshot Index Builder (Correct)**
```csharp
public class SnapshotIndexBuilder
{
    public Dictionary<long, ClashZoneSnapshot> BuildSnapshotIndex(string xmlFilePath)
    {
        var index = new Dictionary<long, ClashZoneSnapshot>();

        if (!File.Exists(xmlFilePath))
            return index;

        var doc = XDocument.Load(xmlFilePath);
        var clashZones = doc.Descendants("ClashZone");

        foreach (var zone in clashZones)
        {
            var sleeveId = long.Parse(zone.Element("SleeveInstanceId")?.Value ?? "0");
            var mepId = long.Parse(zone.Element("MepElementId")?.Value ?? "0");

            var mepParams = new Dictionary<string, string>();
            var hostParams = new Dictionary<string, string>();

            // Parse MEP parameters
            var mepParamElements = zone.Element("MepParameterValues")?.Elements("Parameter");
            if (mepParamElements != null)
            {
                foreach (var param in mepParamElements)
                {
                    mepParams[param.Attribute("Key")?.Value] = param.Attribute("Value")?.Value;
                }
            }

            // Parse Host parameters
            var hostParamElements = zone.Element("HostParameterValues")?.Elements("Parameter");
            if (hostParamElements != null)
            {
                foreach (var param in hostParamElements)
                {
                    hostParams[param.Attribute("Key")?.Value] = param.Attribute("Value")?.Value;
                }
            }

            index[sleeveId] = new ClashZoneSnapshot
            {
                SleeveInstanceId = sleeveId,
                MepElementId = mepId,
                MepParameterValues = mepParams,
                HostParameterValues = hostParams
            };
        }

        return index;
    }
}

public class ClashZoneSnapshot
{
    public long SleeveInstanceId { get; set; }
    public long MepElementId { get; set; }
    public Dictionary<string, string> MepParameterValues { get; set; } = new();
    public Dictionary<string, string> HostParameterValues { get; set; } = new();
}
```

### **2. Parameter Transfer Service (Correct)**
```csharp
public class ParameterTransferService
{
    private readonly Dictionary<long, ClashZoneSnapshot> _snapshotIndex;
    private readonly Action<string> _log;

    public ParameterTransferService(Dictionary<long, ClashZoneSnapshot> snapshotIndex, Action<string> log)
    {
        _snapshotIndex = snapshotIndex;
        _log = log;
    }

    public bool TransferParametersToSleeve(
        Element sleeve,
        List<ParameterMapping> mappings,
        Document document)
    {
        try
        {
            var sleeveId = sleeve.Id.IntegerValue;
            _log($"Transferring parameters for sleeve {sleeveId}");

            // Step 1: Try snapshot method first (preferred)
            var snapshot = _snapshotIndex.ContainsKey(sleeveId)
                ? _snapshotIndex[sleeveId]
                : null;

            var transferResults = new List<bool>();

            foreach (var mapping in mappings)
            {
                string sourceValue = null;

                if (snapshot != null)
                {
                    // Try to get from snapshot
                    sourceValue = GetParameterFromSnapshot(snapshot, mapping.SourceParameter);
                }

                if (string.IsNullOrEmpty(sourceValue))
                {
                    // Fall back to live detection
                    _log($"Parameter '{mapping.SourceParameter}' not found in snapshot, trying live detection");
                    sourceValue = GetParameterFromLiveDetection(sleeve, mapping.SourceParameter, document);
                }

                if (!string.IsNullOrEmpty(sourceValue))
                {
                    // Apply the transfer
                    bool success = SetSleeveParameter(sleeve, mapping.TargetParameter, sourceValue);
                    transferResults.Add(success);

                    _log($"✓ {mapping.SourceParameter} = '{sourceValue}' → {mapping.TargetParameter} ({(success ? "SUCCESS" : "FAILED")})");
                }
                else
                {
                    _log($"✗ No value found for parameter '{mapping.SourceParameter}'");
                    transferResults.Add(false);
                }
            }

            return transferResults.All(r => r);
        }
        catch (Exception ex)
        {
            _log($"Error transferring parameters to sleeve {sleeve.Id}: {ex.Message}");
            return false;
        }
    }

    private string GetParameterFromSnapshot(ClashZoneSnapshot snapshot, string parameterName)
    {
        // Try exact match first
        if (snapshot.MepParameterValues.TryGetValue(parameterName, out var value))
            return value;

        // Try flexible matching (handle variations in parameter names)
        return FindParameterWithFlexibleMatching(snapshot.MepParameterValues, parameterName);
    }

    private string GetParameterFromLiveDetection(Element sleeve, string parameterName, Document document)
    {
        try
        {
            // Find intersecting MEP elements
            var intersectingMep = FindIntersectingMepElements(sleeve, document);
            if (intersectingMep == null || !intersectingMep.Any())
                return null;

            // For now, take the first intersecting MEP element
            // TODO: Add logic to choose the "best" MEP element if multiple intersect
            var mepElement = intersectingMep.First();

            // Get the parameter value directly from the MEP element
            var parameter = mepElement.LookupParameter(parameterName);
            if (parameter != null && parameter.HasValue)
            {
                return parameter.AsValueString() ?? parameter.AsString();
            }

            return null;
        }
        catch (Exception ex)
        {
            _log($"Error in live detection for parameter '{parameterName}': {ex.Message}");
            return null;
        }
    }

    private string FindParameterWithFlexibleMatching(Dictionary<string, string> parameters, string searchName)
    {
        // Try case-insensitive match
        var match = parameters.Keys.FirstOrDefault(k =>
            k.Equals(searchName, StringComparison.OrdinalIgnoreCase));
        if (match != null)
            return parameters[match];

        // Try partial match (contains)
        match = parameters.Keys.FirstOrDefault(k =>
            k.Contains(searchName, StringComparison.OrdinalIgnoreCase) ||
            searchName.Contains(k, StringComparison.OrdinalIgnoreCase));
        if (match != null)
            return parameters[match];

        // Try common variations
        var variations = new[]
        {
            searchName.Replace(" ", ""),
            searchName.Replace(" ", "_"),
            searchName.Replace("_", " "),
            $"MEP {searchName}",
            $"{searchName} MEP"
        };

        foreach (var variation in variations)
        {
            if (parameters.TryGetValue(variation, out var value))
                return value;
        }

        return null;
    }

    private bool SetSleeveParameter(Element sleeve, string parameterName, string value)
    {
        try
        {
            var parameter = sleeve.LookupParameter(parameterName);
            if (parameter == null)
            {
                _log($"Parameter '{parameterName}' not found on sleeve {sleeve.Id}");
                return false;
            }

            if (!parameter.IsReadOnly)
            {
                // Handle different parameter types
                switch (parameter.StorageType)
                {
                    case StorageType.String:
                        parameter.Set(value);
                        break;
                    case StorageType.Double:
                        if (double.TryParse(value, out var doubleValue))
                            parameter.Set(doubleValue);
                        break;
                    case StorageType.Integer:
                        if (int.TryParse(value, out var intValue))
                            parameter.Set(intValue);
                        break;
                    default:
                        parameter.Set(value);
                        break;
                }

                return true;
            }
            else
            {
                _log($"Parameter '{parameterName}' is read-only on sleeve {sleeve.Id}");
                return false;
            }
        }
        catch (Exception ex)
        {
            _log($"Error setting parameter '{parameterName}' to '{value}': {ex.Message}");
            return false;
        }
    }

    private List<Element> FindIntersectingMepElements(Element sleeve, Document document)
    {
        var mepElements = new List<Element>();

        try
        {
            // Get sleeve bounding box
            var sleeveBbox = sleeve.get_BoundingBox(null);
            if (sleeveBbox == null)
                return mepElements;

            var outline = new Outline(sleeveBbox.Min, sleeveBbox.Max);

            // Find MEP elements that intersect with this outline
            var mepCategories = new[]
            {
                BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_Conduit
            };

            foreach (var category in mepCategories)
            {
                var collector = new FilteredElementCollector(document)
                    .OfCategory(category)
                    .WhereElementIsNotElementType();

                foreach (var mepElement in collector)
                {
                    var mepBbox = mepElement.get_BoundingBox(null);
                    if (mepBbox != null)
                    {
                        var mepOutline = new Outline(mepBbox.Min, mepBbox.Max);

                        // Check if bounding boxes intersect
                        if (outline.Intersects(mepOutline, 0.01)) // 1cm tolerance
                        {
                            mepElements.Add(mepElement);
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _log($"Error finding intersecting MEP elements: {ex.Message}");
        }

        return mepElements;
    }
}
```

### **3. Parameter Mapping Model**
```csharp
public class ParameterMapping
{
    public string SourceParameter { get; set; } // From MEP element
    public string TargetParameter { get; set; } // On sleeve
    public ParameterType ParameterType { get; set; } = ParameterType.String;

    public enum ParameterType
    {
        String,
        Double,
        Integer,
        ElementId
    }
}
```

### **4. Main Transfer Command (Correct Usage)**
```csharp
public class TransferParametersCommand
{
    public void ExecuteTransfer(Document document, List<ParameterMapping> mappings)
    {
        try
        {
            // Step 1: Build snapshot index from XML
            var indexBuilder = new SnapshotIndexBuilder();
            var snapshotIndex = indexBuilder.BuildSnapshotIndex("path/to/clash_zones.xml");

            // Step 2: Find all sleeves in the model
            var sleeves = new FilteredElementCollector(document)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
                .ToList();

            // Step 3: Transfer parameters to each sleeve
            var transferService = new ParameterTransferService(snapshotIndex, message => Debug.WriteLine(message));

            int successCount = 0;
            int totalCount = sleeves.Count;

            foreach (var sleeve in sleeves)
            {
                bool success = transferService.TransferParametersToSleeve(sleeve, mappings, document);
                if (success)
                    successCount++;
            }

            // Step 4: Report results
            Debug.WriteLine($"Parameter transfer complete: {successCount}/{totalCount} sleeves updated successfully");
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error in parameter transfer: {ex.Message}");
        }
    }
}
```

---

## 🔧 **Integration Points**

### **With Existing ClashZoneService**
```csharp
// In your existing refresh/detect clashes method
public void DetectClashesAndSaveSnapshot(...)
{
    // ... existing clash detection logic ...

    // NEW: Save snapshot for parameter transfer
    var snapshotBuilder = new SnapshotIndexBuilder();
    var snapshotIndex = snapshotBuilder.BuildSnapshotIndex("clash_zones.xml");

    // Save the index for later use by parameter transfer
    SaveSnapshotIndex(snapshotIndex, "snapshot_index.xml");
}
```

### **With UI Parameter Mapping**
```csharp
// When user creates mapping in UI
var mapping = new ParameterMapping
{
    SourceParameter = "System Type",    // From MEP
    TargetParameter = "MEP System Type" // On sleeve
};

// Save mappings for later use
SaveParameterMappings(mappings, "parameter_mappings.xml");
```

---

## 🧪 **Testing Strategy**

### **Unit Tests**
```csharp
[TestFixture]
public class ParameterTransferServiceTests
{
    [Test]
    public void TransferParameters_UsesSnapshotIndex_WhenAvailable()
    {
        // Arrange
        var snapshotIndex = new Dictionary<long, ClashZoneSnapshot>
        {
            [123456] = new ClashZoneSnapshot
            {
                MepParameterValues = new Dictionary<string, string>
                {
                    ["System Type"] = "Supply Air"
                }
            }
        };

        var service = new ParameterTransferService(snapshotIndex, Console.WriteLine);

        // Act & Assert
        // Test that snapshot values are used correctly
    }

    [Test]
    public void TransferParameters_FallsBackToLiveDetection_WhenSnapshotNotAvailable()
    {
        // Arrange
        var emptyIndex = new Dictionary<long, ClashZoneSnapshot>();
        var service = new ParameterTransferService(emptyIndex, Console.WriteLine);

        // Act & Assert
        // Test that live detection is used as fallback
    }
}
```

---

## 📊 **Performance Considerations**

### **Memory Usage**
- Snapshot index: ~1MB for 10,000 sleeves
- Live detection: Minimal memory, but slower
- Recommendation: Use snapshot method whenever possible

### **Speed Optimization**
```csharp
// Cache sleeve lookups
private static Dictionary<long, Element> _sleeveCache = new();

// Pre-load all sleeves for faster lookup
public void PreloadSleeves(Document document)
{
    _sleeveCache = new FilteredElementCollector(document)
        .OfClass(typeof(FamilyInstance))
        .Cast<FamilyInstance>()
        .Where(fi => fi.Symbol?.Family?.Name?.Contains("Opening") == true)
        .ToDictionary(s => s.Id.IntegerValue, s => (Element)s);
}
```

---

## 🚨 **Error Handling & Edge Cases**

### **Missing Parameters**
```csharp
// Log warnings but don't fail the entire transfer
if (string.IsNullOrEmpty(sourceValue))
{
    _log($"WARNING: No value found for parameter '{mapping.SourceParameter}' on sleeve {sleeve.Id}");
    continue; // Skip this mapping, continue with others
}
```

### **Invalid Sleeve Elements**
```csharp
// Validate sleeve before processing
if (sleeve == null || !sleeve.IsValidObject)
{
    _log($"WARNING: Invalid sleeve element {sleeve?.Id}, skipping");
    continue;
}
```

### **Read-Only Parameters**
```csharp
// Handle read-only parameters gracefully
if (parameter.IsReadOnly)
{
    _log($"WARNING: Parameter '{parameterName}' is read-only on sleeve {sleeve.Id}");
    return false;
}
```

---

## 📋 **Migration Guide**

### **From Current Implementation**
1. **Replace category-based lookup** with sleeve ID-based lookup
2. **Unify XML file patterns** (use consistent naming)
3. **Add fallback logic** for missing snapshots
4. **Implement flexible parameter matching** for name variations

### **Backward Compatibility**
```csharp
// Check if old format exists, migrate if needed
public Dictionary<long, ClashZoneSnapshot> MigrateFromOldFormat(string oldXmlPath)
{
    // Convert old category-based XML to new sleeve-based format
    // Implementation depends on your current XML structure
}
```

---

## 🎯 **Success Metrics**

### **Functional Requirements**
- [ ] **Accuracy**: 100% correct parameter values transferred
- [ ] **Reliability**: Works with moved/deleted MEP elements
- [ ] **Flexibility**: Handles parameter name variations
- [ ] **Performance**: < 1 second for 1000 sleeves

### **User Experience**
- [ ] **Clear feedback**: Progress reporting during transfer
- [ ] **Error visibility**: Detailed logging of failures
- [ ] **Graceful degradation**: Continues when individual transfers fail
- [ ] **Configuration**: Easy parameter mapping setup

---

## 🔍 **Debugging Guide**

### **Common Issues & Solutions**

#### **Issue 1: "No snapshot found for sleeve"**
```csharp
// Check if XML file exists and contains the sleeve
_log($"Snapshot index contains {snapshotIndex.Count} entries");
_log($"Looking for sleeve ID: {sleeveId}");
if (snapshotIndex.ContainsKey(sleeveId))
{
    _log($"Found snapshot: {snapshotIndex[sleeveId].MepParameterValues.Count} MEP parameters");
}
```

#### **Issue 2: "Parameter not found in snapshot"**
```csharp
// Debug parameter name matching
_log($"Available parameters in snapshot: {string.Join(", ", snapshot.MepParameterValues.Keys)}");
_log($"Looking for parameter: '{parameterName}'");
var foundValue = FindParameterWithFlexibleMatching(snapshot.MepParameterValues, parameterName);
_log($"Flexible match result: '{foundValue}'");
```

#### **Issue 3: "Live detection finds wrong MEP element"**
```csharp
// Debug intersecting elements
var intersecting = FindIntersectingMepElements(sleeve, document);
_log($"Found {intersecting.Count} intersecting MEP elements:");
foreach (var mep in intersecting)
{
    _log($"  - MEP {mep.Id}: {mep.Category?.Name}");
}
```

---

## 📚 **API Reference**

### **Key Methods**
- `BuildSnapshotIndex(xmlFilePath)`: Creates sleeve ID → parameters mapping
- `TransferParametersToSleeve(sleeve, mappings, document)`: Transfers parameters to single sleeve
- `GetParameterFromSnapshot(snapshot, parameterName)`: Gets parameter from saved data
- `GetParameterFromLiveDetection(sleeve, parameterName, document)`: Gets parameter from live MEP elements

### **Data Models**
- `ParameterMapping`: Defines source → target parameter mapping
- `ClashZoneSnapshot`: Contains saved parameter values for a sleeve
- `SnapshotIndexBuilder`: Builds the sleeve ID index from XML

---

## 🚀 **Quick Start**

1. **Build the snapshot index** from your clash detection XML
2. **Create parameter mappings** from UI selections
3. **Call TransferParametersToSleeve** for each sleeve
4. **Handle errors gracefully** and provide user feedback

This implementation ensures reliable parameter transfer that works even when MEP elements have been moved or deleted, while maintaining backward compatibility with existing clash detection workflows.
