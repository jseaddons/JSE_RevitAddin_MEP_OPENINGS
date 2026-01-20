# Global Configuration Integration Architecture

## 🎯 OBJECTIVE
Document how global configuration settings (from Configure Button) influence and override UI state settings (from Main UI) to create a unified, rule-based sleeve placement system.

## 🏗️ CONFIGURATION HIERARCHY OVERVIEW

### **Layer 1: Global Business Rules (Highest Priority)**
```
Source: Configure Button → SettingsDialog → GlobalApplicationSettings
Purpose: Business rules that enforce company standards and technical requirements
Override: Always takes precedence over UI preferences
Examples:
├── "Pipes >200mm diameter → Rectangular sleeves"
├── "Minimum clearance: 25mm"
├── "Maximum opening size: 2000mm"
├── "Cable trays always rectangular"
└── "Join openings within 100mm"
```

### **Layer 2: UI User Preferences (Lower Priority)**
```
Source: Main UI Controls (Real-time)
Purpose: User preferences for standard cases and default behavior
Override: Only applies when no global rule exists
Examples:
├── "Default opening type: Circular"
├── "Default clearance: 50mm"
├── "Preferred sleeve family"
└── "Parameter transfer preferences"
```

### **Layer 3: Element-Specific Analysis (Dynamic)**
```
Source: Element property analysis during placement
Purpose: Apply rules based on actual element characteristics
Override: Provides data for global rule evaluation
Examples:
├── Pipe diameter measurement
├── Duct size analysis
├── Element insulation status
└── Element location and angle
```

## 📊 CONFIGURATION INFLUENCE MATRIX

| Configuration Setting | Global Rule Source | UI Preference Source | Influence Level | Example |
|----------------------|-------------------|---------------------|-----------------|---------|
| **Opening Type** | `RoundOpeningsBecomeRectangularIfDiameterGreaterThan` | Radio buttons (Circular/Rectangular) | **HIGH** | Pipe 250mm → Rectangular (ignores UI "Circular") |
| **Minimum Clearance** | `MinOpeningWidth/Height` | Clearance textboxes | **HIGH** | UI sets 10mm, Global min=25mm → Uses 25mm |
| **Maximum Clearance** | `MaxOpeningWidth/Height` | Clearance textboxes | **HIGH** | UI sets 3000mm, Global max=2000mm → Uses 2000mm |
| **Processing Decision** | `ProcessDucts/ProcessPipes` | Category checkboxes | **CRITICAL** | Global: ProcessPipes=false → Skips all pipes |
| **Size Limits** | `MinOpeningWidth/Height` | N/A | **CRITICAL** | Calculated size < minimum → Uses minimum |
| **Rounding Rules** | `RoundOpeningSizesToNearest5mm` | N/A | **MEDIUM** | 347mm → 350mm (if rounding enabled) |
| **Join Rules** | `JoinOpeningsIfDistanceLessThan` | N/A | **MEDIUM** | Nearby openings → Merged into one |
| **Constraint Settings** | `CreateConstraintBetweenOpeningsAndHosts` | N/A | **LOW** | Global rule always applies |

## 🔧 ARCHITECTURE INTEGRATION PLAN

### **Phase 1: Configuration Resolution Service**
```csharp
// New service to handle all configuration conflicts
public class ConfigurationResolutionService
{
    // Resolves conflicts between global rules and UI preferences
    public ResolvedConfiguration ResolveConfiguration(
        string category, 
        ElementProperties elementProps, 
        UIUserPreferences uiPreferences);
}
```

### **Phase 2: Enhanced Placement Services**
```csharp
// Updated services to use resolved configuration
public class PlacementConfigurationService
{
    // Gets effective settings considering global rules
    public EffectiveSettings GetEffectiveSettings(string category, ElementProperties elementProps);
}

public class GlobalConfigurationService
{
    // Provides global rules and defaults
    public GlobalRules GetGlobalRules();
    public ConfigurationDefaults GetConfigurationDefaults();
}
```

### **Phase 3: Command Integration**
```csharp
// Updated commands to use resolved configuration
public class DuctSleeveCommand
{
    public Result ExecuteImpl()
    {
        foreach (var clashZone in clashZones)
        {
            // Analyze element properties
            var elementProps = AnalyzeElementProperties(clashZone.MepElementId);
            
            // Get resolved configuration (global rules + UI preferences)
            var resolvedConfig = ConfigurationResolutionService.Instance
                .ResolveConfiguration("Ducts", elementProps, GetUIUserPreferences());
            
            // Apply resolved configuration to sleeve placement
            PlaceSleeve(clashZone, resolvedConfig);
        }
    }
}
```

## 📋 DETAILED CONFIGURATION INFLUENCE ANALYSIS

### **1. Opening Type Resolution**
```csharp
public string ResolveOpeningType(string category, ElementProperties elementProps, string uiPreference)
{
    // Priority 1: Global size-based rules
    if (category.Equals("Pipes"))
    {
        var diameterThreshold = GlobalConfig.RoundOpeningsBecomeRectangularIfDiameterGreaterThan; // 200mm
        if (elementProps.Diameter > diameterThreshold)
        {
            return "Rectangular"; // Global rule wins
        }
    }
    
    // Priority 2: Global category rules
    if (category.Equals("Cable Trays"))
    {
        return "Rectangular"; // Always rectangular per global rule
    }
    
    // Priority 3: UI preference (fallback)
    return uiPreference;
}
```

### **2. Clearance Resolution**
```csharp
public double ResolveClearance(string category, ElementProperties elementProps, double uiClearance)
{
    // Priority 1: Global minimum clearance
    var globalMinClearance = GlobalConfig.MinOpeningWidth; // 25mm
    if (uiClearance < globalMinClearance)
    {
        return globalMinClearance; // Global rule wins
    }
    
    // Priority 2: Global maximum clearance
    var globalMaxClearance = GlobalConfig.MaxOpeningWidth; // 2000mm
    if (uiClearance > globalMaxClearance)
    {
        return globalMaxClearance; // Global rule wins
    }
    
    // Priority 3: UI preference (within limits)
    return uiClearance;
}
```

### **3. Processing Decision Resolution**
```csharp
public bool ShouldProcessElement(string category, ElementProperties elementProps)
{
    // Priority 1: Global processing flags
    if (!GlobalConfig.ShouldProcessElementType(category))
    {
        return false; // Global rule: Don't process this category
    }
    
    // Priority 2: Global size limits
    if (elementProps.Diameter < GlobalConfig.IgnoreOpeningsSmallerThan)
    {
        return false; // Global rule: Too small to process
    }
    
    // Priority 3: Global angle limits
    if (elementProps.Angle > GlobalConfig.IgnoreOpeningsWithAngleGreaterThan)
    {
        return false; // Global rule: Angle too steep
    }
    
    // Priority 4: UI preference (if element passes global rules)
    return GetUIProcessingPreference(category);
}
```

### **4. Size Calculation Resolution**
```csharp
public SleeveSize ResolveSleeveSize(ElementProperties elementProps, double resolvedClearance)
{
    // Calculate base size from element + clearance
    var baseSize = CalculateBaseSize(elementProps, resolvedClearance);
    
    // Apply global size limits
    return new SleeveSize
    {
        Width = ApplySizeLimits(baseSize.Width, 
            GlobalConfig.MinOpeningWidth, 
            GlobalConfig.MaxOpeningWidth),
        Height = ApplySizeLimits(baseSize.Height, 
            GlobalConfig.MinOpeningHeight, 
            GlobalConfig.MaxOpeningHeight),
        Depth = ApplySizeLimits(baseSize.Depth, 
            GlobalConfig.MinOpeningDepth, 
            GlobalConfig.MaxOpeningDepth)
    };
}
```

### **5. Rounding Rules Resolution**
```csharp
public SleeveSize ApplyRoundingRules(SleeveSize size)
{
    // Apply global rounding rules
    if (GlobalConfig.RoundOpeningSizesToNearest5mm)
    {
        size.Width = RoundToNearest5mm(size.Width);
        size.Height = RoundToNearest5mm(size.Height);
        size.Depth = RoundToNearest5mm(size.Depth);
    }
    
    return size;
}
```

## 🎯 CONFIGURATION RESOLUTION FLOW

### **Step-by-Step Resolution Process:**
```
1. ELEMENT ANALYSIS
   ├── Analyze element properties (size, type, insulation, etc.)
   ├── Determine element category
   └── Extract relevant measurements

2. GLOBAL RULES CHECK
   ├── Check processing flags (ProcessDucts, ProcessPipes, etc.)
   ├── Check size limits (MinOpeningWidth, MaxOpeningWidth, etc.)
   ├── Check type rules (RoundOpeningsBecomeRectangularIfDiameterGreaterThan)
   └── Check angle limits (IgnoreOpeningsWithAngleGreaterThan)

3. UI PREFERENCES RETRIEVAL
   ├── Get current clearance settings from UI
   ├── Get current opening type selection from UI
   ├── Get current parameter mappings from UI
   └── Get current filter conditions from UI

4. CONFLICT RESOLUTION
   ├── Apply global rules first (highest priority)
   ├── Use UI preferences as fallback (lower priority)
   ├── Log resolution decisions for transparency
   └── Create resolved configuration object

5. FINAL CONFIGURATION
   ├── Opening type (global rule or UI preference)
   ├── Clearance values (within global limits)
   ├── Size limits (enforced by global rules)
   ├── Rounding rules (applied from global config)
   └── Constraint settings (from global config)
```

## 📊 RESOLUTION EXAMPLES

### **Example 1: Pipe Opening Type Conflict**
```
Input:
├── Element: Pipe, diameter = 250mm
├── Global Rule: "Pipes >200mm → Rectangular"
└── UI Setting: "Opening Type: Circular"

Resolution Process:
1. Analyze: Pipe diameter = 250mm
2. Global Rule: 250mm > 200mm → Rectangular ✅
3. UI Setting: Circular ❌ (ignored due to global rule)
4. Result: Rectangular sleeve

Log Output:
[Config] Pipe 250mm: Global rule applied - Rectangular (overrides UI: Circular)
```

### **Example 2: Clearance Limits Conflict**
```
Input:
├── Element: Duct, size = 300mm
├── UI Clearance: 10mm
├── Global Min: 25mm
└── Global Max: 2000mm

Resolution Process:
1. Analyze: Duct size = 300mm
2. UI Clearance: 10mm
3. Global Min Check: 10mm < 25mm → Use 25mm ✅
4. Global Max Check: 25mm < 2000mm ✅
5. Result: Clearance = 25mm

Log Output:
[Config] Duct 300mm: Global minimum applied - 25mm (overrides UI: 10mm)
```

### **Example 3: Size Limits Conflict**
```
Input:
├── Element: Large duct, size = 1500mm
├── UI Clearance: 1000mm
├── Calculated Size: 3500mm (1500 + 2×1000)
├── Global Max: 2000mm

Resolution Process:
1. Analyze: Duct size = 1500mm
2. Calculate: 1500 + 2×1000 = 3500mm
3. Global Max Check: 3500mm > 2000mm → Use 2000mm ✅
4. Result: Opening size = 2000mm

Log Output:
[Config] Large duct: Global maximum applied - 2000mm (overrides calculated: 3500mm)
```

## 🔄 INTEGRATION WITH EXISTING ARCHITECTURE

### **Current Architecture (Before Integration):**
```
RefreshService → Saves UI state to XML
Commands → Read configuration from XML (stale data)
```

### **New Architecture (After Integration):**
```
RefreshService → Saves only intersection data to XML
ConfigurationResolutionService → Resolves global rules + UI preferences
Commands → Use resolved configuration (real-time + rule-based)
```

### **Integration Points:**
1. **RefreshService**: Remove UI state saving, keep intersection detection
2. **Commands**: Use ConfigurationResolutionService for all configuration
3. **UI**: Continue to provide real-time preferences
4. **GlobalConfig**: Provide business rules and limits
5. **Logging**: Track all resolution decisions for transparency

## 🧪 TESTING STRATEGY

### **Unit Tests:**
- [ ] ConfigurationResolutionService resolves conflicts correctly
- [ ] Global rules take precedence over UI preferences
- [ ] Size limits are enforced properly
- [ ] Rounding rules are applied correctly

### **Integration Tests:**
- [ ] Commands use resolved configuration
- [ ] UI preferences are respected when no global rule exists
- [ ] Global rules override UI preferences when applicable
- [ ] Logging shows resolution decisions

### **End-to-End Tests:**
- [ ] Complete flow with global rule conflicts
- [ ] Complete flow with UI preference conflicts
- [ ] Complete flow with mixed conflicts
- [ ] Performance impact of configuration resolution

## 📝 IMPLEMENTATION CHECKLIST

### **Phase 1: Core Resolution Service**
- [ ] Create ConfigurationResolutionService
- [ ] Implement opening type resolution
- [ ] Implement clearance resolution
- [ ] Implement processing decision resolution
- [ ] Add comprehensive logging

### **Phase 2: Command Integration**
- [ ] Update DuctSleeveCommand to use resolved configuration
- [ ] Update PipeSleeveCommand to use resolved configuration
- [ ] Update other commands to use resolved configuration
- [ ] Test command execution with resolved configuration

### **Phase 3: UI Integration**
- [ ] Update UI to show when global rules override preferences
- [ ] Add configuration source indicators
- [ ] Update logging to show resolution decisions
- [ ] Test UI behavior with global rule conflicts

### **Phase 4: Advanced Rules**
- [ ] Implement size-based type rules
- [ ] Implement rounding rules
- [ ] Implement join/merge rules
- [ ] Test complex configuration scenarios

## 🎯 SUCCESS CRITERIA

### **Functional Requirements:**
- [ ] Global rules always take precedence over UI preferences
- [ ] UI preferences are used when no global rule exists
- [ ] All configuration conflicts are resolved transparently
- [ ] Logging shows all resolution decisions

### **Non-Functional Requirements:**
- [ ] Configuration resolution adds <10ms per element
- [ ] No breaking changes to existing functionality
- [ ] Clear separation of concerns maintained
- [ ] Easy to add new global rules

### **User Experience:**
- [ ] Users understand when global rules override their preferences
- [ ] Logging provides clear feedback on resolution decisions
- [ ] UI shows current effective settings
- [ ] No unexpected behavior due to hidden global rules

## 📅 IMPLEMENTATION TIMELINE

- **Week 1**: ConfigurationResolutionService core implementation
- **Week 2**: Command integration and testing
- **Week 3**: UI integration and user feedback
- **Week 4**: Advanced rules and final testing

This architecture ensures that global business rules are always enforced while maintaining user flexibility for standard cases.
