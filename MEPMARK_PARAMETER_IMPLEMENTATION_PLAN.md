# 🏷️ MEPMARK PARAMETER IMPLEMENTATION PLAN

## 🎯 Objective

Implement custom `MEPMARK` shared parameter with discipline-specific prefixes for MEP openings, replacing the generic Mark parameter system.

---

## 📋 Requirements Summary

### **1. Parameter Name**
- ✅ Use `MEPMARK` (shared parameter, NOT the standard Revit "Mark" parameter)
- ✅ Must be added to all 4 universal family files

### **2. Discipline-Specific Prefixes**

| Discipline | Category | Prefix | Example (Individual) | Example (Cluster) |
|-----------|----------|--------|---------------------|-------------------|
| **HVAC** | Ducts | `DCT` | `DCT-001` | `DCT_C_-001` |
| **HVAC** | Duct Accessories (Dampers) | `DMP` | `DMP-001` | `DMP_C_-001` |
| **Plumbing** | Pipes | `PLU` | `PLU-001` | `PLU_C_-001` |
| **Electrical** | Cable Trays | `ELE` | `ELE-001` | `ELE_C_-001` |

### **3. User Prefix (from UI)**
- ✅ User can input a custom prefix in UI (e.g., "JSE", "PROJ", "A")
- ✅ User prefix is PREPENDED to discipline prefix
- ✅ Example: User prefix "JSE" + Duct = `JSE-DCT-001`

### **4. Cluster Suffix**
- ✅ Cluster sleeves get `_C_` inserted after discipline prefix
- ✅ Format: `{UserPrefix}-{DisciplinePrefix}_C_-{Number}`
- ✅ Example: `JSE-DCT_C_-001` (JSE project, Duct cluster #1)

---

## 🔧 Complete Format Specification

### **Individual Sleeve Mark Format:**
```
{UserPrefix}-{DisciplinePrefix}-{Number:000}

Components:
- UserPrefix: From UI input (e.g., "JSE", "PROJ", "A") - OPTIONAL
- DisciplinePrefix: DCT/DMP/PLU/ELE (based on MEP category)
- Number: Sequential 001, 002, 003...

Examples:
- JSE-DCT-001  (User prefix + Duct individual #1)
- JSE-DMP-001  (User prefix + Damper individual #1)
- JSE-PLU-001  (User prefix + Pipe individual #1)
- JSE-ELE-001  (User prefix + Cable Tray individual #1)
- DCT-001      (No user prefix, Duct individual #1)
```

### **Cluster Sleeve Mark Format:**
```
{UserPrefix}-{DisciplinePrefix}_C_-{Number:000}

Components:
- UserPrefix: From UI input - OPTIONAL
- DisciplinePrefix: DCT/DMP/PLU/ELE
- _C_: Cluster indicator (literal text)
- Number: Sequential 001, 002, 003...

Examples:
- JSE-DCT_C_-001  (User prefix + Duct cluster #1)
- JSE-DMP_C_-001  (User prefix + Damper cluster #1)
- JSE-PLU_C_-001  (User prefix + Pipe cluster #1)
- JSE-ELE_C_-001  (User prefix + Cable Tray cluster #1)
- DCT_C_-001      (No user prefix, Duct cluster #1)
```

---

## 📊 Mapping Table

### **Old System (Mark parameter) → New System (MEPMARK parameter)**

| Old Prefix | Old Format | New Discipline Prefix | New Format (Individual) | New Format (Cluster) |
|-----------|-----------|---------------------|----------------------|---------------------|
| `P` | `P001` | `PLU` | `PLU-001` | `PLU_C_-001` |
| `D` | `D001` | `DCT` | `DCT-001` | `DCT_C_-001` |
| `D` | `D001` | `DMP` | `DMP-001` | `DMP_C_-001` |
| `C` | `C001` | `ELE` | `ELE-001` | `ELE_C_-001` |
| `CP` | `CP001` | `PLU` | `PLU_C_-001` | N/A |
| `CD` | `CD001` | `DCT` | `DCT_C_-001` | N/A |
| `CC` | `CC001` | `ELE` | `ELE_C_-001` | N/A |

**Key Changes:**
- ✅ More descriptive prefixes (DCT vs D, PLU vs P, ELE vs C)
- ✅ Consistent cluster notation (`_C_`)
- ✅ User prefix support
- ✅ Uses MEPMARK parameter instead of Mark

---

## 🏗️ Implementation Architecture

### **Phase 1: Shared Parameter Setup**

#### **Step 1.1: Create Shared Parameter File**
**Location:** `Resources/Opening family shared parameter.txt` (already exists)

**Required Parameter:**
```
*META	VERSION	MINVERSION
META	2	1
*GROUP	ID	NAME
GROUP	1	MEP Opening Parameters
*PARAM	GUID	NAME	DATATYPE	DATACATEGORY	GROUP	VISIBLE	DESCRIPTION	USERMODIFIABLE	HIDEWHENNOVALUE
PARAM	{GUID}	MEPMARK	TEXT	-	1	1	MEP Opening Mark (Discipline-Specific)	1	0
```

#### **Step 1.2: Add MEPMARK to Universal Families**
**Files to update:**
1. `Resources/RectangularOpeningOnWall.rfa`
2. `Resources/CircularOpeningOnWall.rfa`
3. `Resources/RectangularOpeningOnSlab.rfa`
4. `Resources/CircularOpeningOnSlab.rfa`

**Steps for each family:**
1. Open family in Revit Family Editor
2. Manage → Project Parameters → Add
3. Select "Shared parameter"
4. Browse to `Opening family shared parameter.txt`
5. Select `MEPMARK` parameter
6. Add to "MEP Opening Parameters" group
7. Make it Instance parameter
8. Save family

---

### **Phase 2: Mark Generation Service**

#### **Create: `Services/MepMarkService.cs`**

```csharp
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    public class MepMarkService
    {
        /// <summary>
        /// Generate MEPMARK value for individual sleeve
        /// </summary>
        public static string GenerateIndividualMark(
            string category,        // "Ducts", "Pipes", etc. (from MepCategoryConstants)
            string userPrefix,      // From UI input (e.g., "JSE", "PROJ")
            int sequenceNumber)     // 1, 2, 3...
        {
            string disciplinePrefix = GetDisciplinePrefix(category);
            
            if (string.IsNullOrWhiteSpace(userPrefix))
            {
                // No user prefix
                return $"{disciplinePrefix}-{sequenceNumber:000}";
            }
            else
            {
                // With user prefix
                return $"{userPrefix}-{disciplinePrefix}-{sequenceNumber:000}";
            }
        }
        
        /// <summary>
        /// Generate MEPMARK value for cluster sleeve
        /// </summary>
        public static string GenerateClusterMark(
            string category,        // "Ducts", "Pipes", etc.
            string userPrefix,      // From UI input
            int sequenceNumber)     // 1, 2, 3...
        {
            string disciplinePrefix = GetDisciplinePrefix(category);
            
            if (string.IsNullOrWhiteSpace(userPrefix))
            {
                // No user prefix
                return $"{disciplinePrefix}_C_-{sequenceNumber:000}";
            }
            else
            {
                // With user prefix
                return $"{userPrefix}-{disciplinePrefix}_C_-{sequenceNumber:000}";
            }
        }
        
        /// <summary>
        /// Get discipline-specific prefix
        /// </summary>
        private static string GetDisciplinePrefix(string category)
        {
            return category switch
            {
                MepCategoryConstants.DUCTS => "DCT",
                MepCategoryConstants.DUCT_ACCESSORIES => "DMP",
                MepCategoryConstants.PIPES => "PLU",
                MepCategoryConstants.CABLE_TRAYS => "ELE",
                _ => "OPN" // Generic opening
            };
        }
        
        /// <summary>
        /// Set MEPMARK parameter on sleeve
        /// </summary>
        public static bool SetMepMark(FamilyInstance sleeve, string markValue)
        {
            try
            {
                var mepMarkParam = sleeve.LookupParameter("MEPMARK");
                if (mepMarkParam != null && !mepMarkParam.IsReadOnly)
                {
                    mepMarkParam.Set(markValue);
                    DebugLogger.Info($"[MepMarkService] Set MEPMARK = '{markValue}' on sleeve {sleeve.Id}");
                    return true;
                }
                else
                {
                    DebugLogger.Warning($"[MepMarkService] MEPMARK parameter not found or read-only on sleeve {sleeve.Id}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MepMarkService] Error setting MEPMARK: {ex.Message}");
                return false;
            }
        }
    }
}
```

---

### **Phase 3: UI Integration**

#### **Update: `Views/EmergencyMainDialog.cs`**

**Add User Prefix Input:**
```csharp
// In form initialization:
private TextBox _userPrefixTextBox;

private void InitializeMarkPrefixControls()
{
    var prefixLabel = new Label
    {
        Text = "Project Prefix (optional):",
        Location = new Point(10, 10),
        AutoSize = true
    };
    
    _userPrefixTextBox = new TextBox
    {
        Name = "userPrefixTextBox",
        Location = new Point(150, 8),
        Size = new Size(100, 20),
        PlaceholderText = "e.g., JSE, PROJ"
    };
    
    // Add to panel
    _markParametersPanel.Controls.Add(prefixLabel);
    _markParametersPanel.Controls.Add(_userPrefixTextBox);
}

// Save prefix to OpeningConditions
private void SaveMarkPrefixToConditions()
{
    string userPrefix = _userPrefixTextBox.Text.Trim().ToUpper();
    // Store in OpeningConditions.UserMarkPrefix
}
```

---

### **Phase 4: Integration with Placement**

#### **Update: `Services/UniversalSleevePlacerService.cs`**

**After placing each sleeve:**
```csharp
// In SetSleeveParameters method:

// Generate and set MEPMARK
string userPrefix = _conditions.UserMarkPrefix ?? "";
int sleeveNumber = GetNextSleeveNumber(clashZone.MepElementCategory);
string mepMark = MepMarkService.GenerateIndividualMark(
    clashZone.MepElementCategory,
    userPrefix,
    sleeveNumber);

MepMarkService.SetMepMark(sleeveInstance, mepMark);
```

---

### **Phase 5: Integration with Clustering**

#### **Update: `Services/UniversalClusterService.cs`**

**After placing cluster sleeve:**
```csharp
// In PlaceClusterSleeve method:

// Generate and set MEPMARK for cluster
string userPrefix = GetUserPrefixFromConditions();
int clusterNumber = GetNextClusterNumber(groupKey.systemType);
string clusterMark = MepMarkService.GenerateClusterMark(
    groupKey.systemType,
    userPrefix,
    clusterNumber);

MepMarkService.SetMepMark(inst, clusterMark);
```

---

## 🔍 **Numbering Strategy**

### **Separate Counters per Category:**

```csharp
// Track counters separately
private int _ductIndividualCounter = 1;
private int _ductClusterCounter = 1;
private int _pipeIndividualCounter = 1;
private int _pipeClusterCounter = 1;
private int _damperIndividualCounter = 1;
private int _damperClusterCounter = 1;
private int _cableTrayIndividualCounter = 1;
private int _cableTrayClusterCounter = 1;
```

### **Example Sequence:**
```
Ducts (22 individual, 3 clusters):
  Individual: DCT-001, DCT-002, ..., DCT-022
  Clusters: DCT_C_-001, DCT_C_-002, DCT_C_-003

Pipes (15 individual, 2 clusters):
  Individual: PLU-001, PLU-002, ..., PLU-015
  Clusters: PLU_C_-001, PLU_C_-002

Dampers (2 individual, 0 clusters):
  Individual: DMP-001, DMP-002
```

---

## 📝 **Format Confirmation Questions**

### **Question 1: Cluster Mark Format**

**Which format do you prefer?**

**Option A:**
```
DCT_C_-001  (Discipline_C_-Number)
JSE-DCT_C_-001  (UserPrefix-Discipline_C_-Number)
```

**Option B:**
```
DCT-C-001  (Discipline-C-Number)
JSE-DCT-C-001  (UserPrefix-Discipline-C-Number)
```

**Option C:**
```
DCT_C001  (Discipline_CNumber)
JSE-DCT_C001  (UserPrefix-Discipline_CNumber)
```

**Please confirm:** Based on your message, I'm implementing **Option A** - is this correct?

---

### **Question 2: User Prefix Format**

**When user inputs prefix "JSE", should marks be:**

**Option A (with dash):**
```
JSE-DCT-001
JSE-DCT_C_-001
```

**Option B (no dash):**
```
JSEDCT-001
JSEDCT_C_-001
```

**Please confirm:** I'm assuming **Option A (with dash)** - correct?

---

### **Question 3: Numbering Scope**

**Should numbering be:**

**Option A: Per Category**
- Ducts: DCT-001, DCT-002, DCT-003...
- Pipes: PLU-001, PLU-002, PLU-003...
- (Each category starts at 001)

**Option B: Global**
- All openings: 001, 002, 003...
- Across all categories

**Please confirm:** I'm assuming **Option A (per category)** - correct?

---

### **Question 4: Counter Reset**

**When should counters reset?**

**Option A: Never**
- First run: DCT-001, DCT-002
- Second run (new sleeves): DCT-003, DCT-004
- (Remembers previous max number)

**Option B: Each Run**
- First run: DCT-001, DCT-002
- Second run: DCT-001, DCT-002 again
- (Starts from 001 each time)

**Please confirm:** I'm assuming **Option B (reset each run)** - correct?

---

## 🏗️ Implementation Checklist

### **Phase 1: Shared Parameter (Family Files)**
- [ ] Open `RectangularOpeningOnWall.rfa`
- [ ] Add `MEPMARK` shared parameter (Text, Instance)
- [ ] Repeat for `CircularOpeningOnWall.rfa`
- [ ] Repeat for `RectangularOpeningOnSlab.rfa`
- [ ] Repeat for `CircularOpeningOnSlab.rfa`
- [ ] Load updated families into project
- [ ] Test parameter exists on placed sleeves

### **Phase 2: Service Class**
- [ ] Create `Services/MepMarkService.cs`
- [ ] Implement `GenerateIndividualMark(category, userPrefix, number)`
- [ ] Implement `GenerateClusterMark(category, userPrefix, number)`
- [ ] Implement `GetDisciplinePrefix(category)` helper
- [ ] Implement `SetMepMark(sleeve, markValue)` helper

### **Phase 3: UI Controls**
- [ ] Add "Project Prefix" TextBox to EmergencyMainDialog
- [ ] Save user prefix to OpeningConditions
- [ ] Load user prefix from saved conditions

### **Phase 4: Placement Integration**
- [ ] Update `UniversalSleevePlacerService.SetSleeveParameters()`
- [ ] Generate MEPMARK after placing each sleeve
- [ ] Track counter per category

### **Phase 5: Clustering Integration**
- [ ] Update `UniversalClusterService.PlaceClusterSleeve()`
- [ ] Generate cluster MEPMARK after placing cluster
- [ ] Track cluster counter per category

### **Phase 6: Testing**
- [ ] Test with user prefix "JSE"
- [ ] Test with no user prefix
- [ ] Test individual sleeves get correct marks
- [ ] Test cluster sleeves get `_C_` suffix
- [ ] Test numbering is sequential per category
- [ ] Test marks persist after Revit restart

---

## 📚 Code Examples

### **Usage in Placement:**
```csharp
// After placing individual duct sleeve
string mark = MepMarkService.GenerateIndividualMark(
    category: "Ducts",
    userPrefix: "JSE",
    sequenceNumber: 1);
// Result: "JSE-DCT-001"

MepMarkService.SetMepMark(sleeveInstance, mark);
```

### **Usage in Clustering:**
```csharp
// After placing duct cluster sleeve
string clusterMark = MepMarkService.GenerateClusterMark(
    category: "Ducts",
    userPrefix: "JSE",
    sequenceNumber: 1);
// Result: "JSE-DCT_C_-001"

MepMarkService.SetMepMark(clusterInstance, clusterMark);
```

---

## ✅ Success Criteria

- ✅ MEPMARK parameter exists in all 4 universal families
- ✅ User can input project prefix in UI
- ✅ Individual sleeves get marks like `JSE-DCT-001`
- ✅ Cluster sleeves get marks like `JSE-DCT_C_-001`
- ✅ Numbering is sequential per category
- ✅ Marks are visible in Revit schedules
- ✅ Marks persist after save/reload

---

## 🎯 Next Steps

1. **CONFIRM FORMAT** - Answer the 4 questions above
2. **UPDATE FAMILIES** - Add MEPMARK shared parameter
3. **IMPLEMENT SERVICE** - Create MepMarkService.cs
4. **UPDATE UI** - Add user prefix input
5. **INTEGRATE PLACEMENT** - Generate marks during placement
6. **INTEGRATE CLUSTERING** - Generate cluster marks
7. **TEST** - Verify all scenarios work

---

## 📋 Open Questions

Please confirm:
1. ✅ Cluster format: `DCT_C_-001` (with underscores and dash)?
2. ✅ User prefix separator: `JSE-DCT-001` (with dash)?
3. ✅ Numbering scope: Per category (DCT-001, PLU-001 both start at 001)?
4. ✅ Counter reset: Each run starts at 001?

**Once confirmed, I'll proceed with implementation!**

