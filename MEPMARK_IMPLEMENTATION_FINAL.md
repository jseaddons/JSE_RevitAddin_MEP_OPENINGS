# 🏷️ MEPMARK PARAMETER - FINAL IMPLEMENTATION PLAN

## 🎯 Design Philosophy

**Just like clearances:**
- ✅ Provide sensible defaults (DCT, PLU, ELE, DMP)
- ✅ User can keep defaults or customize
- ✅ No fallbacks, no auto-detection
- ✅ User's choice is final

---

## 🎨 **Ultra-Compact UI (Final Design)**

```
┌──────────────────────────────────────────────────────┐
│ Clearance Settings (Existing Panel)                  │
├──────────────────────────────────────────────────────┤
│ MEP Type:  [Ducts                    ▼]  ← EXISTING  │
│                                                      │
│ Rectangular Duct:                                    │
│   Normal: [50] mm   Insulated: [100] mm             │
│ Round Duct:                                          │
│   Normal: [50] mm   Insulated: [100] mm             │
└──────────────────────────────────────────────────────┘

┌──────────────────────────────────────────────────────┐
│ Mark Prefix                                          │
├──────────────────────────────────────────────────────┤
│ Project:     [JSE    ]                               │
│ Discipline:  [DCT    ]  ← Changes with MEP Type ↑    │
└──────────────────────────────────────────────────────┘
```

**That's it!** Just 2 tiny textboxes. Minimalist and clean.

---

## 💾 **Data Model**

### **Add to `Models/OpeningConditions.cs`:**

```csharp
public class MarkPrefixSettings
{
    // Project-level prefix (applied to all categories)
    public string ProjectPrefix { get; set; } = "";
    
    // Category-specific discipline prefixes (defaults provided)
    public string DuctPrefix { get; set; } = "DCT";
    public string PipePrefix { get; set; } = "PLU";
    public string CableTrayPrefix { get; set; } = "ELE";
    public string DamperPrefix { get; set; } = "DMP";
    
    /// <summary>
    /// Get discipline prefix for a category
    /// </summary>
    public string GetDisciplinePrefix(string category)
    {
        return category switch
        {
            MepCategoryConstants.DUCTS => DuctPrefix,
            MepCategoryConstants.PIPES => PipePrefix,
            MepCategoryConstants.CABLE_TRAYS => CableTrayPrefix,
            MepCategoryConstants.DUCT_ACCESSORIES => DamperPrefix,
            _ => "OPN" // Generic fallback
        };
    }
}

// Add property to OpeningConditions class
public MarkPrefixSettings MarkPrefixes { get; set; } = new MarkPrefixSettings();
```

---

## 🔧 **UI Implementation**

### **Step 1: Create Controls (EmergencyMainDialog.cs)**

```csharp
private TextBox _projectPrefixTextBox;
private TextBox _disciplinePrefixTextBox;

// In-memory storage (synced with UI)
private Dictionary<string, string> _categoryPrefixes = new Dictionary<string, string>
{
    { MepCategoryConstants.DUCTS, "DCT" },
    { MepCategoryConstants.PIPES, "PLU" },
    { MepCategoryConstants.CABLE_TRAYS, "ELE" },
    { MepCategoryConstants.DUCT_ACCESSORIES, "DMP" }
};

private void CreateMarkPrefixPanel()
{
    var markPanel = new Panel 
    { 
        BorderStyle = BorderStyle.FixedSingle,
        Size = new Size(300, 70),
        Location = new Point(10, Y_POSITION) // Below clearance panel
    };
    
    var titleLabel = new Label 
    { 
        Text = "Mark Prefix", 
        Font = new Font(Font, FontStyle.Bold),
        Location = new Point(5, 5),
        AutoSize = true
    };
    
    // Row 1: Project Prefix
    var projectLabel = new Label { Text = "Project:", Location = new Point(10, 25), Width = 70 };
    _projectPrefixTextBox = new TextBox 
    { 
        Location = new Point(85, 23), 
        Size = new Size(80, 20),
        CharacterCasing = CharacterCasing.Upper // Auto-uppercase
    };
    
    // Row 2: Discipline Prefix
    var disciplineLabel = new Label { Text = "Discipline:", Location = new Point(10, 48), Width = 70 };
    _disciplinePrefixTextBox = new TextBox 
    { 
        Location = new Point(85, 46), 
        Size = new Size(80, 20),
        Text = "DCT", // Default for Ducts
        CharacterCasing = CharacterCasing.Upper // Auto-uppercase
    };
    
    markPanel.Controls.AddRange(new Control[] 
    { 
        titleLabel, 
        projectLabel, _projectPrefixTextBox,
        disciplineLabel, _disciplinePrefixTextBox 
    });
    
    this.Controls.Add(markPanel);
}
```

### **Step 2: Wire to Clearance Dropdown**

```csharp
private void WireMarkPrefixToMepTypeDropdown()
{
    // When MEP type changes in clearance dropdown
    _clearanceMepTypeDropdown.SelectedIndexChanged += (s, e) =>
    {
        string selectedType = _clearanceMepTypeDropdown.SelectedItem?.ToString() ?? "Ducts";
        string category = MepCategoryConstants.Normalize(selectedType);
        
        // Update discipline prefix textbox to show saved value for this category
        _disciplinePrefixTextBox.Text = _categoryPrefixes[category];
    };
    
    // When user types in discipline prefix textbox
    _disciplinePrefixTextBox.TextChanged += (s, e) =>
    {
        string selectedType = _clearanceMepTypeDropdown.SelectedItem?.ToString() ?? "Ducts";
        string category = MepCategoryConstants.Normalize(selectedType);
        
        // Save what user typed for this category
        _categoryPrefixes[category] = _disciplinePrefixTextBox.Text.Trim().ToUpper();
    };
}
```

### **Step 3: Save on OK Click**

```csharp
private void OkButton_Click(object sender, EventArgs e)
{
    // Read all settings from UI
    var conditions = new OpeningConditions
    {
        ClearanceSettings = ReadClearanceSettingsFromUI(),
        OpeningTypePreferences = ReadOpeningTypePreferencesFromUI(),
        MarkPrefixes = ReadMarkPrefixesFromUI() // NEW
    };
    
    // Save to XML
    SaveConditionsToXml(conditions);
}

private MarkPrefixSettings ReadMarkPrefixesFromUI()
{
    return new MarkPrefixSettings
    {
        ProjectPrefix = _projectPrefixTextBox.Text.Trim().ToUpper(),
        DuctPrefix = _categoryPrefixes[MepCategoryConstants.DUCTS],
        PipePrefix = _categoryPrefixes[MepCategoryConstants.PIPES],
        CableTrayPrefix = _categoryPrefixes[MepCategoryConstants.CABLE_TRAYS],
        DamperPrefix = _categoryPrefixes[MepCategoryConstants.DUCT_ACCESSORIES]
    };
}
```

### **Step 4: Load on Dialog Open**

```csharp
private void LoadMarkPrefixesFromConditions(OpeningConditions conditions)
{
    if (conditions?.MarkPrefixes != null)
    {
        _projectPrefixTextBox.Text = conditions.MarkPrefixes.ProjectPrefix;
        
        // Update in-memory dictionary
        _categoryPrefixes[MepCategoryConstants.DUCTS] = conditions.MarkPrefixes.DuctPrefix;
        _categoryPrefixes[MepCategoryConstants.PIPES] = conditions.MarkPrefixes.PipePrefix;
        _categoryPrefixes[MepCategoryConstants.CABLE_TRAYS] = conditions.MarkPrefixes.CableTrayPrefix;
        _categoryPrefixes[MepCategoryConstants.DUCT_ACCESSORIES] = conditions.MarkPrefixes.DamperPrefix;
        
        // Show prefix for currently selected MEP type
        string selectedType = _clearanceMepTypeDropdown.SelectedItem?.ToString() ?? "Ducts";
        string category = MepCategoryConstants.Normalize(selectedType);
        _disciplinePrefixTextBox.Text = _categoryPrefixes[category];
    }
}
```

---

## 📋 **Complete Workflow:**

### **User Configures Prefixes:**
```
1. User opens dialog
2. Enters "JSE" in Project Prefix textbox
3. Clearance dropdown shows "Ducts" (default)
4. Discipline Prefix shows "DCT" (default)
5. User keeps it or changes to "HVAC"
6. User selects "Pipes" in clearance dropdown
7. Discipline Prefix AUTO-CHANGES to "PLU"
8. User keeps it or changes to "PLB"
9. User selects "Cable Trays"
10. Discipline Prefix AUTO-CHANGES to "ELE"
11. User clicks OK
12. ALL prefixes saved to CONDITIONS XML
```

### **No Buttons, No Saves, No Preview - Just Works!**

---

## 🎯 **Mark Generation (During Placement/Clustering)**

```csharp
// Read from saved conditions
var markPrefixes = conditions.MarkPrefixes;

// Generate individual mark
string mark = string.IsNullOrEmpty(markPrefixes.ProjectPrefix)
    ? $"{markPrefixes.DuctPrefix}-001"
    : $"{markPrefixes.ProjectPrefix}-{markPrefixes.DuctPrefix}-001";
// Result: "JSE-DCT-001" or "DCT-001"

// Generate cluster mark
string clusterMark = string.IsNullOrEmpty(markPrefixes.ProjectPrefix)
    ? $"{markPrefixes.DuctPrefix}_C_-001"
    : $"{markPrefixes.ProjectPrefix}-{markPrefixes.DuctPrefix}_C_-001";
// Result: "JSE-DCT_C_-001" or "DCT_C_-001"

// Set parameter
sleeveInstance.LookupParameter("MEPMARK")?.Set(mark);
```

---

## ✅ **This Design is:**

- ✅ **Minimal** - Just 2 textboxes
- ✅ **Intuitive** - Follows clearance panel pattern
- ✅ **Automatic** - Syncs with MEP type selection
- ✅ **Flexible** - User can customize all 4 categories
- ✅ **Clean** - No clutter, no extra buttons

**Perfect! Shall I proceed with implementation?**
