# Architecture Separation Plan - Comprehensive Reference

## 🎯 OBJECTIVE
Separate concerns between data collection (Refresh) and placement configuration (Commands) to create a clean, maintainable architecture.

## 📚 RELATED DOCUMENTATION
- **GLOBAL_CONFIGURATION_INTEGRATION_ARCHITECTURE.md**: Detailed plan for how global settings influence UI state
- **Configuration Hierarchy**: Global Rules (highest) → UI Preferences (lower) → Element Analysis (dynamic)

## 📋 CONFIGURATION LAYERS

### LAYER 1: GLOBAL CONFIG (Application-wide)
**Purpose**: Application-level defaults and behavior
**Source**: SettingsDialog, ApplicationProfileService
**Examples**:
- Default clearance values
- Opening family paths  
- Parameter prefix settings
- Global behavior flags
- System-wide configurations

### LAYER 2: FILTER CONDITIONS (Filter-specific)
**Purpose**: Which elements to process
**Source**: UI Filter Selection
**Examples**:
- Selected reference files
- Selected host files
- Selected MEP categories
- Filter-specific settings
- Element filtering criteria

### LAYER 3: PARAMETER MAPPING (Category-specific)
**Purpose**: How to map parameters between elements
**Source**: UI Parameter Configuration
**Examples**:
- Parameter prefix mappings
- Category-specific parameter rules
- Transfer service configurations
- Parameter extraction settings

### LAYER 4: PLACEMENT SETTINGS (Real-time UI)
**Purpose**: How to place sleeves
**Source**: UI Controls (Real-time)
**Examples**:
- Clearance values
- Opening type selection
- Placement options
- Real-time user preferences

## 🏗️ PROPOSED ARCHITECTURE

### CURRENT (Combined - Problematic)
```
Refresh Button → RefreshService → Saves:
├── Clash Zones (intersections)
├── Clearance Settings (UI state) ❌
├── Filter Conditions ❌
└── Category-specific XML files

OK Button → Commands → Uses:
├── Clash Zones from XML
├── Clearance Settings from ClearanceManager (UI state)
└── Filter Conditions from XML ❌
```

### TARGET (Separated - Clean)
```
Refresh Button → RefreshService → Saves ONLY:
├── Clash Zones (intersections)
├── Element IDs and positions
└── Category-specific XML files

OK Button → Commands → Gets:
├── Clash Zones from XML (intersections)
├── Clearance Settings from UI directly (real-time)
├── Filter Conditions from UI directly (real-time)
├── Opening Type from UI directly (real-time)
├── Parameter Mapping from Services
└── Global Config from Services
```

## 🔧 IMPLEMENTATION PHASES

### PHASE 1: Create Configuration Services ✅
**Goal**: Extract configuration logic into dedicated services

**Services to Create**:
1. `GlobalConfigurationService` - Application-wide settings ✅ **COMPLETED**
2. `FilterConfigurationService` - Filter and element selection
3. `ParameterMappingService` - Parameter prefixes and mappings
4. `PlacementConfigurationService` - Real-time UI settings
5. `ConfigurationResolutionService` - **NEW**: Resolves conflicts between global rules and UI preferences

**Files to Create**:
- `Services/Configuration/GlobalConfigurationService.cs` ✅ **COMPLETED**
- `Services/Configuration/FilterConfigurationService.cs`
- `Services/Configuration/ParameterMappingService.cs`
- `Services/Configuration/PlacementConfigurationService.cs`
- `Services/Configuration/ConfigurationResolutionService.cs` ✅ **PLANNED**

**Files to Modify**:
- Extract logic from `ApplicationProfileService`
- Extract logic from `EmergencyMainDialog.cs`
- Extract logic from `ParameterExtractionService.cs`

**Testing**:
- [ ] Services can be instantiated
- [ ] Services return expected configuration data
- [ ] Services integrate with existing UI
- [ ] No breaking changes to current functionality

### PHASE 2: Update RefreshService ⏳
**Goal**: Remove UI state saving from RefreshService

**Changes**:
1. Remove clearance saving logic
2. Remove UI state persistence
3. Keep only intersection detection and clash zone storage
4. Update XML format to contain only intersection data

**Files to Modify**:
- `Services/RefreshService.cs`
- `Views/EmergencyMainDialog.cs` (OnRefreshClick method)

**Testing**:
- [ ] Refresh still detects intersections correctly
- [ ] XML files contain only intersection data
- [ ] No clearance/UI state saved in XML
- [ ] Performance improvement in refresh speed

### PHASE 3: Update Commands ⏳
**Goal**: Commands use new services for configuration

**Changes**:
1. Update `DuctSleeveCommand` to use new services
2. Update `PipeSleeveCommand` to use new services
3. Update other commands to use new services
4. Remove dependency on XML for configuration data

**Files to Modify**:
- `Commands/DuctSleeveCommand.cs`
- `Commands/PipeSleeveCommand.cs`
- `Commands/CableTraySleeveCommand.cs`
- `Commands/FireDamperPlaceCommand.cs`
- `Services/OpeningCommandOrchestrator.cs`

**Testing**:
- [ ] Commands get clearance from UI services
- [ ] Commands get opening type from UI services
- [ ] Commands get filter conditions from UI services
- [ ] Commands get parameter mapping from services
- [ ] Commands still read intersection data from XML
- [ ] Sleeve placement works correctly

### PHASE 4: Update UI Integration ⏳
**Goal**: UI uses new services for real-time data

**Changes**:
1. Update UI to use new services for real-time data
2. Remove old clearance saving logic from UI
3. Update parameter dropdown population
4. Ensure real-time configuration updates

**Files to Modify**:
- `Views/EmergencyMainDialog.cs`
- `Views/SettingsDialog.cs`
- UI event handlers

**Testing**:
- [ ] UI shows current configuration in real-time
- [ ] Parameter dropdowns work correctly
- [ ] Clearance changes are immediately available
- [ ] Opening type changes are immediately available
- [ ] Complete flow works end-to-end

## 🎯 BENEFITS OF SEPARATED ARCHITECTURE

### PERFORMANCE
- ✅ Faster refresh (no unnecessary UI state saving)
- ✅ Smaller XML files (only intersection data)
- ✅ Reduced I/O operations

### MAINTAINABILITY
- ✅ Single responsibility principle
- ✅ Clear separation of concerns
- ✅ Easier to debug and test
- ✅ Changes to one layer don't affect others

### USER EXPERIENCE
- ✅ Real-time configuration updates
- ✅ No stale data issues
- ✅ Immediate feedback on UI changes
- ✅ Predictable behavior

### CODE QUALITY
- ✅ Reduced complexity
- ✅ Better testability
- ✅ Cleaner interfaces
- ✅ Easier to extend

## 📊 DATA FLOW DIAGRAM

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   REFRESH       │    │   UI SERVICES    │    │   COMMANDS      │
│   (Data Only)   │    │  (Real-time)     │    │ (Placement)     │
├─────────────────┤    ├──────────────────┤    ├─────────────────┤
│ • Intersections │    │ • Clearances     │    │ • Read XML      │
│ • Clash Zones   │    │ • Opening Types  │    │ • Get UI State  │
│ • Element IDs   │    │ • Filter Conds   │    │ • Get Mappings  │
│ • Positions     │    │ • Param Prefix   │    │ • Place Sleeves │
└─────────────────┘    └──────────────────┘    └─────────────────┘
         │                       │                       │
         └───────────────────────┼───────────────────────┘
                                 │
                    ┌─────────────────────┐
                    │  GLOBAL CONFIG      │
                    │  (Application-wide) │
                    ├─────────────────────┤
                    │ • Default Values    │
                    │ • Family Paths      │
                    │ • System Settings   │
                    │ • Global Flags      │
                    └─────────────────────┘
```

## 🧪 TESTING STRATEGY

### Unit Tests
- [ ] Each service can be instantiated independently
- [ ] Services return expected configuration data
- [ ] Services handle edge cases gracefully

### Integration Tests
- [ ] Services work together correctly
- [ ] UI integration works seamlessly
- [ ] Command execution uses correct configuration

### End-to-End Tests
- [ ] Complete flow: Refresh → OK → Sleeve Placement
- [ ] Real-time configuration changes affect placement
- [ ] Performance improvements are measurable

## 📝 IMPLEMENTATION NOTES

### Current Dependencies
- `ApplicationProfileService` - Contains global config logic
- `EmergencyMainDialog` - Contains UI state logic
- `ParameterExtractionService` - Contains parameter mapping logic
- `RefreshService` - Contains mixed responsibilities

### Migration Strategy
1. **Backward Compatibility**: Ensure existing functionality works during migration
2. **Incremental Changes**: Test each phase before proceeding
3. **Rollback Plan**: Keep backup of working code
4. **Documentation**: Update all relevant documentation

### Risk Mitigation
- **Testing**: Comprehensive testing at each phase
- **Backup**: Keep working version as backup
- **Incremental**: Small, testable changes
- **Validation**: Verify each step works before proceeding

## 🚀 SUCCESS CRITERIA

### Phase 1 Success
- [ ] All configuration services created
- [ ] Services can be instantiated and used
- [ ] No breaking changes to existing functionality
- [ ] Clear separation of configuration concerns

### Phase 2 Success
- [ ] RefreshService only handles intersection detection
- [ ] XML files contain only intersection data
- [ ] Performance improvement measurable
- [ ] No UI state saved in XML

### Phase 3 Success
- [ ] Commands use new services for configuration
- [ ] Commands still read intersection data from XML
- [ ] Sleeve placement works correctly
- [ ] Real-time configuration updates work

### Phase 4 Success
- [ ] UI uses new services for real-time data
- [ ] Complete flow works end-to-end
- [ ] Performance improvements achieved
- [ ] Architecture is clean and maintainable

## 📅 TIMELINE

- **Phase 1**: Configuration Services (Current)
- **Phase 2**: RefreshService Cleanup (Next)
- **Phase 3**: Command Updates (After Phase 2)
- **Phase 4**: UI Integration (Final)

Each phase will be tested thoroughly before proceeding to the next.
