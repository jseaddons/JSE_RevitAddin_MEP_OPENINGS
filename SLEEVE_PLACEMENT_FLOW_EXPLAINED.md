# 🔧 How Sleeve Placement Works - Plain English

## 📋 **Simple Overview**

**It places sleeves ONE BY ONE, but all within a SINGLE transaction (batch commit).**

---

## 🔄 **The Flow (Step by Step)**

### **Step 1: Read XML File (BEFORE Transaction Starts)**
- ✅ Reads clash zone data from XML file (e.g., `ventilation_ducts.xml`)
- ✅ Loads all clash zones into memory as `ClashZone` objects
- ✅ Each `ClashZone` contains:
  - Placement point (`SleevePlacementPoint` - X, Y, Z coordinates)
  - MEP element size (Width, Height, Diameter)
  - MEP element category (Ducts, Pipes, Dampers, etc.)
  - Structural element info (Wall, Floor, Framing)
  - **NO transaction yet - just reading data**

```
Example ClashZone:
- Id: 123
- Placement Point: (100ft, 200ft, 15ft)
- MEP Width: 300mm
- MEP Height: 150mm
- Category: "Ducts"
```

---

### **Step 2: Start ONE Transaction (Batch Mode)**
- ✅ Opens **ONE** Revit transaction
- ✅ This transaction will wrap **ALL** sleeve placements
- ✅ Transaction name: "Place [Category] Sleeves" (e.g., "Place Ducts Sleeves")

```csharp
using (var t = new Transaction(_doc, $"Place {_category} Sleeves"))
{
    t.Start();  // ONE transaction for ALL sleeves
    // ... place all sleeves here ...
    t.Commit(); // Save all at once
}
```

---

### **Step 3: Loop Through Clash Zones (ONE BY ONE)**
- ✅ Loops through each `ClashZone` object in the list
- ✅ Processes them **sequentially** (one after another)

```csharp
foreach (var clashZone in sortedClashZones)  // ONE BY ONE loop
{
    // Place this one sleeve...
}
```

**For EACH clash zone, the code:**

#### **3a. Gets Placement Point**
- ✅ Reads from `clashZone.SleevePlacementPoint` (already loaded from XML)
- ✅ No XML reading here - just using data from memory
- ✅ Example: `XYZ(100ft, 200ft, 15ft)`

#### **3b. Gets MEP Size**
- ✅ Reads from `clashZone.MepElementWidth`, `MepElementHeight`, `MepElementDiameter`
- ✅ Already loaded from XML - no file access needed
- ✅ Example: Width=300mm, Height=150mm

#### **3c. Calculates Clearance**
- ✅ Calls `GetClearanceFromConditions(category, mepSize)`
- ✅ Looks up clearance value from CONDITIONS XML file (in memory)
- ✅ **Calculates final sleeve size**: `Raw Size + (2 × Clearance)`
- ✅ Example:
  - Raw duct: 300mm × 150mm
  - Clearance: 50mm
  - Final sleeve: 400mm × 250mm

#### **3d. Finds Nearest Level**
- ✅ Uses pre-cached list of levels (no repeated searching)
- ✅ For walls/framing: finds nearest **bottom** level (where element starts)
- ✅ For floors: finds nearest level overall

#### **3e. Places ONE Sleeve**
- ✅ Calls `_doc.Create.NewFamilyInstance(...)`
- ✅ Creates **ONE** sleeve in Revit
- ✅ Sets all parameters (Width, Height, Diameter, etc.)
- ✅ **Still inside the same transaction** - not saved to Revit yet

```csharp
// Place ONE sleeve
var sleeveInstance = _doc.Create.NewFamilyInstance(
    placementPoint,    // From XML (clashZone.SleevePlacementPoint)
    familySymbol,      // Pre-loaded family
    nearestLevel,      // Found above
    StructuralType.NonStructural
);

// Set size (with clearance applied)
SetSleeveParameters(sleeveInstance, finalWidth, finalHeight, finalDiameter);
```

#### **3f. Updates ClashZone in Memory**
- ✅ Updates `clashZone.SleeveInstanceId` with new sleeve's ID
- ✅ Updates `clashZone.SleeveWidth`, `SleeveHeight`, etc.
- ✅ Sets `clashZone.IsResolved = true`
- ✅ **NO XML file write yet** - just updating in-memory objects

---

### **Step 4: Loop Continues**
- ✅ Goes to next clash zone
- ✅ Repeats Steps 3a-3f for next sleeve
- ✅ Places next sleeve (still in same transaction)
- ✅ Continues until all clash zones processed

**Example:**
```
Transaction Started
├── ClashZone 1 → Place Sleeve 1
├── ClashZone 2 → Place Sleeve 2
├── ClashZone 3 → Place Sleeve 3
├── ...
└── ClashZone 100 → Place Sleeve 100
Transaction Committed (ALL 100 sleeves saved at once)
```

---

### **Step 5: Commit Transaction (Save ALL)**
- ✅ After **ALL** sleeves placed, transaction commits
- ✅ **All sleeves appear in Revit at once** (batch commit)
- ✅ No partial saves - either all succeed or all rollback

---

### **Step 6: Update XML Files (Batched)**
- ✅ After transaction commits, updates XML files
- ✅ Groups updates by category/file (not per sleeve)
- ✅ Example:
  - All "Ducts" updates → `ventilation_ducts.xml`
  - All "Pipes" updates → `ventilation_pipes.xml`
- ✅ **Batched write** - loads XML once, updates all clash zones, saves once
- ✅ **NOT** one XML write per sleeve

---

## 📊 **Summary Table**

| Aspect | How It Works |
|--------|-------------|
| **XML Reading** | ✅ **BEFORE** transaction - all clash zones loaded into memory |
| **Transaction** | ✅ **ONE** transaction wraps all placements |
| **Placement** | ✅ **ONE BY ONE** - loops through each clash zone sequentially |
| **Placement Point** | ✅ From XML (already in memory - `clashZone.SleevePlacementPoint`) |
| **Clearance** | ✅ Calculated per sleeve from CONDITIONS XML (in memory) |
| **Final Size** | ✅ Raw Size + (2 × Clearance) - calculated per sleeve |
| **Revit Save** | ✅ **BATCH** - all sleeves saved at once when transaction commits |
| **XML Update** | ✅ **BATCH** - grouped by category, written after all placements |

---

## 🎯 **Key Points**

1. **Sequential Processing**: Sleeves are placed one after another (not in parallel)
2. **Batch Transaction**: All sleeves are in ONE transaction - all saved together
3. **No Per-Sleeve XML Writes**: XML is updated in batches, not after each sleeve
4. **Data Pre-Loaded**: All XML data is in memory before placement starts
5. **Clearance Applied**: Each sleeve gets clearance calculation before placement

---

## ⚡ **Performance Optimizations**

- ✅ **Pre-cache levels** - no repeated searching
- ✅ **Pre-cache family symbols** - no repeated loading
- ✅ **Batch XML updates** - not per-sleeve writes
- ✅ **Single transaction** - faster than many small transactions
- ✅ **In-memory data** - no file I/O during placement loop

---

## 📝 **Example Timeline**

```
0:00 - Read XML file (ventilation_ducts.xml)
       → Load 500 clash zones into memory

0:01 - Start Transaction "Place Ducts Sleeves"

0:01 - Loop Start:
       ├── ClashZone 1 → Place Sleeve 1
       ├── ClashZone 2 → Place Sleeve 2
       ├── ClashZone 3 → Place Sleeve 3
       ├── ...
       └── ClashZone 500 → Place Sleeve 500

0:45 - Loop End (all 500 sleeves placed)

0:45 - Commit Transaction
       → All 500 sleeves appear in Revit at once

0:46 - Update XML file (ventilation_ducts.xml)
       → Write all 500 updated clash zones at once

0:47 - Done!
```

---

## 🔍 **Code Reference**

**Placement Loop:**
```292:708:Services/UniversalSleevePlacerService.cs
                foreach (var clashZone in sortedClashZones)
                {
                    try
                    {
                        // Get placement point (from XML, already in memory)
                        var placementPointChosen = clashZone.SleevePlacementPoint;
                        
                        // Calculate clearance and final size
                        var clearance = GetClearanceFromConditions(...);
                        finalWidth = rawWidth + (2 * clearance);
                        
                        // Place ONE sleeve
                        var sleeveInstance = _doc.Create.NewFamilyInstance(
                            adjustedPlacementPoint,
                            familySymbol,
                            nearestLevel,
                            StructuralType.NonStructural
                        );
                        
                        // Update clash zone in memory (not XML yet)
                        clashZone.SleeveInstanceId = sleeveInstance.Id.IntegerValue;
                        clashZone.IsResolved = true;
                    }
                }
```

**Transaction Wrapper:**
```81:109:Commands/UniversalSleevePlacementCommand.cs
                using (var t = new Transaction(_doc, $"Place {_category} Sleeves"))
                {
                    if (t.Start() == TransactionStatus.Started)
                    {
                        // Place all sleeves inside ONE transaction
                        var result = placerService.PlaceAllSleevesInTransaction(filteredClashZones);
                        
                        // Commit ALL sleeves at once
                        var status = t.Commit();
                    }
                }
```

---

## ✅ **Answer to Your Questions**

**Q: Does it place sleeves one by one or in batches?**
- **A:** One by one (sequential loop), but all in ONE transaction (batch commit)

**Q: Does it read XML during placement?**
- **A:** No - XML is read BEFORE placement starts. All data is in memory.

**Q: How does it get placement points?**
- **A:** From `clashZone.SleevePlacementPoint` - already loaded from XML into memory.

**Q: How does it set clearance values?**
- **A:** Calculates per sleeve: `GetClearanceFromConditions()` → `Raw Size + (2 × Clearance)`

**Q: When does it place - batches or one by one?**
- **A:** Places one by one (loop), but saves in batch (single transaction commit)

