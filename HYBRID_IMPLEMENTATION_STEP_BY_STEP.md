# Hybrid Implementation - Step by Step Guide

## 🎯 **Overview**
This document provides a detailed, step-by-step implementation guide for the Hybrid Approach, where we combine UI State Manager for dynamic changes with Filter-based persistence. Each step is designed to be testable independently.

## 📚 **Related Documentation**
- **[Clash Zone & BCF Export Dependency Architecture](CLASH_ZONE_BCF_DEPENDENCY_ARCHITECTURE.md)** - Critical architectural understanding of BCF dependency and clash zone management
- **[BCF Openings Export Implementation Plan](BCF_OPENINGS_EXPORT_IMPLEMENTATION_PLAN.md)** - Detailed BCF implementation phases

## ✅ **Current Implementation Status (2025-09-24)**

### **Completed Features**
1. **✅ Duplicate Clash Detection** - Fixed issue where multiple refreshes created duplicate clash zones
2. **✅ DuctSleeveCommand Integration** - Fixed issue where DuctSleeveCommand ignored passed clash zones
3. **✅ 3D Section Box Filtering** - Added logic to only process clash zones visible in current 3D section box
4. **✅ Filter Reuse Logic** - Fixed issue where filters were recreated instead of reused

### **Key Architectural Understanding**
- **BCF Dependency**: Change detection requires BCF Opening Status for persistent storage
- **Clash Zone Purpose**: Clash zones serve as "working set" for preview and processing
- **3D Section Box Value**: Essential for performance optimization in large projects

### **Next Phase: BCF Implementation**
The system is now ready for **Phase 3: BCF Opening Status Implementation** as outlined in the BCF Export Implementation Plan.

## 🏗️ **Architecture Overview (Three-Tier System)**

```
1. GLOBAL APPLICATION CONFIGURATION
├── Location: C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\GlobalSettings.xml
├── Purpose: Filter-independent, user-wide settings
├── Examples: Default paths, UI preferences, default CONFIGURATION settings
├── Storage: XML file in user's AppData folder
└── Scope: Applies to ALL projects for this user

2. PROJECT FILTER CONFIGURATION  
├── Location: C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Projects\[ProjectName]\Filters\[FilterName].xml
├── Purpose: Project-specific, filter-dependent settings
├── FILTER: File selections, category selections, basic opening conditions
├── CONFIGURATION: Advanced processing rules (Limits, Elements, etc.)
├── Storage: XML file in user's AppData folder (project-specific subfolder)
└── Scope: Applies to THIS project/filter only

3. UI STATE MANAGER (Session-Based)
├── Location: Memory (temporary during session)
├── Purpose: Dynamic changes during current session
├── Examples: Live CONFIGURATION changes, real-time preview
├── Storage: Auto-save to Project Filter Configuration XML file
└── Scope: Current session only (temporary)

User Experience Flow:
├── App Startup → Load Global Settings → Set Defaults
├── Project Open → Load Project Filter → Override Defaults
├── User changes settings → UI State Manager → Immediate preview
├── Auto-save → UI State Manager → Project Filter → XML File
├── Manual Save → UI State Manager → Project Filter → XML File
└── Manual Load → XML File → Project Filter → UI State Manager

### **📋 File Structure:**
```
C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\
├── GlobalSettings.xml (Global defaults)
├── Families\ (Default families)
├── BCF\ (BCF files)
└── Projects\
    ├── OfficeBuilding_2024\
    │   └── Filters\
    │       ├── FireFighting.xml
    │       ├── Plumbing.xml
    │       └── Electrical.xml
    ├── Hospital_2024\
    │   └── Filters\
    │       ├── FireFighting.xml
    │       ├── Plumbing.xml
    │       └── Electrical.xml
    └── School_2024\
        └── Filters\
            ├── FireFighting.xml
            ├── Plumbing.xml
            └── Electrical.xml
```
```

## 📋 **Implementation Steps (Three-Tier Configuration System)**

### **STEP 1: Create Global Application Configuration**

**Objective**: Create global application settings that are filter-independent and user-wide

**Files to Create:**
- `Models/GlobalApplicationSettings.cs`
- `Services/GlobalConfigurationManager.cs`

**Implementation:**

```csharp
// Models/GlobalApplicationSettings.cs
using System;
using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Global application settings that are filter-independent and user-wide
    /// Based on CONVOID's ConVoidSettings.xml approach
    /// </summary>
    [XmlRoot("GlobalApplicationSettings")]
    public class GlobalApplicationSettings
    {
        // Installation and Paths (Filter-Independent)
        public string DefaultFamiliesPath { get; set; } = @"C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Families";
        public string DefaultBCFPath { get; set; } = @"C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\BCF";
        public string DefaultFiltersPath { get; set; } = @"C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Projects\[ProjectName]\Filters";
        public string DefaultProjectsPath { get; set; } = @"C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Projects";
        
        // Application Preferences (Filter-Independent)
        public string UILanguage { get; set; } = "English";
        public bool AutoSaveEnabled { get; set; } = true;
        public int AutoSaveIntervalSeconds { get; set; } = 30;
        public bool ShowNotifications { get; set; } = true;
        public bool EnableDebugLogging { get; set; } = false;
        
        // Default Opening Settings (Global Defaults - Filter-Independent)
        public double DefaultDuctClearance { get; set; } = 50.0;
        public double DefaultPipeClearance { get; set; } = 75.0;
        public double DefaultCableTrayClearance { get; set; } = 25.0;
        public double DefaultMinOpeningSize { get; set; } = 0.1;
        public double DefaultMaxOpeningSize { get; set; } = 1000.0;
        
        // CONFIGURATION Settings (Global Defaults - Filter-Independent)
        // LIMITS Section (REQUIRED - All Green Highlighted)
        public double MinOpeningSize { get; set; } = 80.0; // Ignore openings smaller than
        public double MaxCircularDiameter { get; set; } = 200.0; // Round openings become rectangular if diameter is greater than
        public double JoinOpeningsDistance { get; set; } = 100.0; // Join openings at a distance of
        public double MaxOpeningAngle { get; set; } = 45.0; // Ignore openings with an angle greater than
        public bool CreateSlopedOpenings { get; set; } = true; // Create openings with a slope
        public bool RoundUpDimensions { get; set; } = false; // Round up opening dimensions (Do not round up)
        public double RoundingIncrement { get; set; } = 5.0; // Rounding increment
        
        // ELEMENT FILTER Section (REQUIRED - Green Highlighted)
        public string ElementFilterView { get; set; } = "3D Coordination"; // Element Filter dropdown selection
        
        // MANAGE Section (REQUIRED - Important for workflow)
        public bool ResetApprovalStatusOnChanges { get; set; } = false; // Reset approval status of openings when changes occur
        public double DimensionChangeThreshold { get; set; } = 1.0; // Openings won't be marked as changed if change in dimensions is less than (mm)
        public double LocationChangeThreshold { get; set; } = 1.0; // Openings won't be marked as changed if change in location is less than (mm)
        
        // ELEMENTS Section (REMOVED - Red X Marks)
        // CreateVerticalParallelOpenings - REMOVED (Red X)
        // CreateHorizontalParallelOpenings - REMOVED (Red X)
        // CutOpeningWithHosts - REMOVED (Not highlighted)
        // CreateConstraintBetweenOpeningsAndHosts - REMOVED (Not highlighted)
        
        // BCF Settings (Global Defaults - Filter-Independent)
        public string DefaultBCFVersion { get; set; } = "2.1";
        public bool CompressBCFOutput { get; set; } = true;
        public bool GenerateViewpoints { get; set; } = true;
        public bool GenerateSnapshots { get; set; } = true;
        
        // User Information (Filter-Independent)
        public string DefaultUserName { get; set; } = "";
        public string DefaultUserDiscipline { get; set; } = "HVAC";
        public string DefaultUserInitials { get; set; } = "";
        
        public DateTime LastModified { get; set; } = DateTime.Now;
        public string Version { get; set; } = "1.0.0";
    }
}
```

```csharp
// Services/GlobalConfigurationManager.cs
using System;
using System.IO;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Singleton manager for global application settings
    /// Based on CONVOID's ConVoidSettings.xml approach
    /// </summary>
    public class GlobalConfigurationManager
    {
        private static GlobalConfigurationManager? _instance;
        private static readonly object _lock = new object();
        
        private GlobalApplicationSettings _settings;
        private readonly string _settingsPath;
        
        private GlobalConfigurationManager()
        {
            // Global Settings Location: C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\GlobalSettings.xml
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var configPath = Path.Combine(appDataPath, "JSE_MEP_Openings");
            _settingsPath = Path.Combine(configPath, "GlobalSettings.xml");
            
            LoadSettings();
        }
        
        public static GlobalConfigurationManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                            _instance = new GlobalConfigurationManager();
                    }
                }
                return _instance;
            }
        }
        
        public GlobalApplicationSettings Settings => _settings;
        
        public void SaveSettings()
        {
            try
            {
                var directory = Path.GetDirectoryName(_settingsPath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);
                    
                _settings.LastModified = DateTime.Now;
                
                var serializer = new XmlSerializer(typeof(GlobalApplicationSettings));
                using (var writer = new StreamWriter(_settingsPath))
                {
                    serializer.Serialize(writer, _settings);
                }
                
                DebugLogger.Log($"Global settings saved to: {_settingsPath}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Failed to save global settings: {ex.Message}");
                throw;
            }
        }
        
        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var serializer = new XmlSerializer(typeof(GlobalApplicationSettings));
                    using (var reader = new StreamReader(_settingsPath))
                    {
                        _settings = (GlobalApplicationSettings)serializer.Deserialize(reader)!;
                    }
                    DebugLogger.Log($"Global settings loaded from: {_settingsPath}");
                }
                else
                {
                    _settings = new GlobalApplicationSettings();
                    SaveSettings(); // Create default settings file
                    DebugLogger.Log("Default global settings created");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Failed to load global settings: {ex.Message}");
                _settings = new GlobalApplicationSettings();
            }
        }
        
        public void ResetToDefaults()
        {
            _settings = new GlobalApplicationSettings();
            SaveSettings();
            DebugLogger.Log("Global settings reset to defaults");
        }
        
        public void UpdateUserInfo(string userName, string discipline, string initials)
        {
            _settings.DefaultUserName = userName;
            _settings.DefaultUserDiscipline = discipline;
            _settings.DefaultUserInitials = initials;
            SaveSettings();
            DebugLogger.Log($"User info updated: {userName} ({discipline}) - {initials}");
        }
    }
}
```

**Testing Step 1:**
1. Create the files `Models/GlobalApplicationSettings.cs` and `Services/GlobalConfigurationManager.cs`
2. Build the project to ensure no compilation errors
3. Test the singleton pattern by accessing `GlobalConfigurationManager.Instance`
4. Test save/load operations
5. Verify that the XML file is created in the correct location
6. Test reset to defaults functionality

**Expected Result:** Global configuration system compiles successfully and can save/load settings to/from XML file.

---

### **STEP 2: Create UI State Manager Models**

**Objective**: Create the data models for dynamic UI state management that use global defaults

**Files to Create:**
- `Models/UIStateManagerModels.cs`
- `Models/OpeningConditionsState.cs`
- `Models/ConfigurationSettingsState.cs`

**Implementation:**

```csharp
// Models/UIStateManagerModels.cs
using System;
using System.ComponentModel;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Represents the current state of opening conditions in the UI
    /// </summary>
    public class OpeningConditionsState : INotifyPropertyChanged
    {
        private double _ductClearance = 50.0;
        private double _pipeClearance = 75.0;
        private double _cableTrayClearance = 25.0;
        private double _oversizeValue = 0.0;
        private double _levelAdjustment = 0.0;
        private bool _isModified = false;
        private DateTime _lastModified = DateTime.Now;

        public double DuctClearance
        {
            get => _ductClearance;
            set
            {
                if (_ductClearance != value)
                {
                    _ductClearance = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(DuctClearance));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double PipeClearance
        {
            get => _pipeClearance;
            set
            {
                if (_pipeClearance != value)
                {
                    _pipeClearance = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(PipeClearance));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double CableTrayClearance
        {
            get => _cableTrayClearance;
            set
            {
                if (_cableTrayClearance != value)
                {
                    _cableTrayClearance = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(CableTrayClearance));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double OversizeValue
        {
            get => _oversizeValue;
            set
            {
                if (_oversizeValue != value)
                {
                    _oversizeValue = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(OversizeValue));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double LevelAdjustment
        {
            get => _levelAdjustment;
            set
            {
                if (_levelAdjustment != value)
                {
                    _levelAdjustment = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(LevelAdjustment));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool IsModified
        {
            get => _isModified;
            private set
            {
                if (_isModified != value)
                {
                    _isModified = value;
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public DateTime LastModified
        {
            get => _lastModified;
            private set
            {
                if (_lastModified != value)
                {
                    _lastModified = value;
                    OnPropertyChanged(nameof(LastModified));
                }
            }
        }

        public void MarkAsSaved()
        {
            _isModified = false;
            OnPropertyChanged(nameof(IsModified));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// Represents the current state of configuration settings in the UI
    /// </summary>
    public class ConfigurationSettingsState : INotifyPropertyChanged
    {
        private double _minOpeningSize = 0.1;
        private double _maxOpeningSize = 1000.0;
        private bool _includeInsulation = true;
        private bool _autoUpdateOpenings = false;
        private bool _isModified = false;
        private DateTime _lastModified = DateTime.Now;

        public double MinOpeningSize
        {
            get => _minOpeningSize;
            set
            {
                if (_minOpeningSize != value)
                {
                    _minOpeningSize = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(MinOpeningSize));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public double MaxOpeningSize
        {
            get => _maxOpeningSize;
            set
            {
                if (_maxOpeningSize != value)
                {
                    _maxOpeningSize = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(MaxOpeningSize));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool IncludeInsulation
        {
            get => _includeInsulation;
            set
            {
                if (_includeInsulation != value)
                {
                    _includeInsulation = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(IncludeInsulation));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool AutoUpdateOpenings
        {
            get => _autoUpdateOpenings;
            set
            {
                if (_autoUpdateOpenings != value)
                {
                    _autoUpdateOpenings = value;
                    _isModified = true;
                    _lastModified = DateTime.Now;
                    OnPropertyChanged(nameof(AutoUpdateOpenings));
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public bool IsModified
        {
            get => _isModified;
            private set
            {
                if (_isModified != value)
                {
                    _isModified = value;
                    OnPropertyChanged(nameof(IsModified));
                }
            }
        }

        public DateTime LastModified
        {
            get => _lastModified;
            private set
            {
                if (_lastModified != value)
                {
                    _lastModified = value;
                    OnPropertyChanged(nameof(LastModified));
                }
            }
        }

        public void MarkAsSaved()
        {
            _isModified = false;
            OnPropertyChanged(nameof(IsModified));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
```

**Testing Step 1:**
1. Create the file `Models/UIStateManagerModels.cs`
2. Build the project to ensure no compilation errors
3. Test the models by creating instances and changing properties
4. Verify that `IsModified` flag works correctly
5. Verify that `PropertyChanged` events fire correctly

**Expected Result:** Models compile successfully and property change notifications work.

---

### **STEP 3: Create UI State Manager Service**

**Objective**: Create the service that manages dynamic UI state using global defaults

**Files to Create:**
- `Services/UIStateManagerService.cs`

**Implementation:**

```csharp
// Services/UIStateManagerService.cs
using System;
using System.Timers;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Manages the dynamic UI state for opening conditions and configuration settings
    /// Uses global defaults from GlobalConfigurationManager
    /// </summary>
    public class UIStateManagerService : IDisposable
    {
        private readonly OpeningConditionsState _openingConditions;
        private readonly ConfigurationSettingsState _configurationSettings;
        private readonly Timer _autoSaveTimer;
        private readonly FilterManagementService _filterManagementService;
        private readonly GlobalConfigurationManager _globalConfig;
        private bool _disposed = false;

        public event EventHandler<OpeningConditionsChangedEventArgs>? OpeningConditionsChanged;
        public event EventHandler<ConfigurationSettingsChangedEventArgs>? ConfigurationSettingsChanged;
        public event EventHandler<AutoSaveEventArgs>? AutoSaveTriggered;

        public UIStateManagerService(FilterManagementService filterManagementService)
        {
            _filterManagementService = filterManagementService ?? throw new ArgumentNullException(nameof(filterManagementService));
            _globalConfig = GlobalConfigurationManager.Instance;
            
            // Initialize with global defaults
            _openingConditions = new OpeningConditionsState
            {
                DuctClearance = _globalConfig.Settings.DefaultDuctClearance,
                PipeClearance = _globalConfig.Settings.DefaultPipeClearance,
                CableTrayClearance = _globalConfig.Settings.DefaultCableTrayClearance,
                OversizeValue = 0.0,
                LevelAdjustment = 0.0
            };
            
            _configurationSettings = new ConfigurationSettingsState
            {
                MinOpeningSize = _globalConfig.Settings.DefaultMinOpeningSize,
                MaxOpeningSize = _globalConfig.Settings.DefaultMaxOpeningSize,
                IncludeInsulation = true,
                AutoUpdateOpenings = false
            };
            
            // Auto-save timer (use global setting)
            var autoSaveInterval = _globalConfig.Settings.AutoSaveEnabled ? 
                _globalConfig.Settings.AutoSaveIntervalSeconds * 1000 : 0;
            _autoSaveTimer = new Timer(autoSaveInterval);
            _autoSaveTimer.Elapsed += OnAutoSaveTimerElapsed;
            _autoSaveTimer.AutoReset = false; // Only fire once per change
            
            // Subscribe to property changes
            _openingConditions.PropertyChanged += OnOpeningConditionsPropertyChanged;
            _configurationSettings.PropertyChanged += OnConfigurationSettingsPropertyChanged;
        }

        public OpeningConditionsState OpeningConditions => _openingConditions;
        public ConfigurationSettingsState ConfigurationSettings => _configurationSettings;

        public bool HasUnsavedChanges => _openingConditions.IsModified || _configurationSettings.IsModified;

        private void OnOpeningConditionsPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(OpeningConditionsState.IsModified) && _openingConditions.IsModified)
            {
                // Start auto-save timer
                _autoSaveTimer.Stop();
                _autoSaveTimer.Start();
                
                // Notify subscribers
                OpeningConditionsChanged?.Invoke(this, new OpeningConditionsChangedEventArgs(_openingConditions));
            }
        }

        private void OnConfigurationSettingsPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ConfigurationSettingsState.IsModified) && _configurationSettings.IsModified)
            {
                // Start auto-save timer
                _autoSaveTimer.Stop();
                _autoSaveTimer.Start();
                
                // Notify subscribers
                ConfigurationSettingsChanged?.Invoke(this, new ConfigurationSettingsChangedEventArgs(_configurationSettings));
            }
        }

        private void OnAutoSaveTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            try
            {
                AutoSaveTriggered?.Invoke(this, new AutoSaveEventArgs());
                SaveToFilter();
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Auto-save failed: {ex.Message}");
            }
        }

        public void SaveToFilter()
        {
            try
            {
                // Convert UI state to filter data
                var openingConditions = new OpeningConditions
                {
                    DuctClearance = _openingConditions.DuctClearance,
                    PipeClearance = _openingConditions.PipeClearance,
                    CableTrayClearance = _openingConditions.CableTrayClearance,
                    OversizeValue = _openingConditions.OversizeValue,
                    LevelAdjustment = _openingConditions.LevelAdjustment
                };

                var configurationSettings = new ConfigurationSettings
                {
                    MinOpeningSize = _configurationSettings.MinOpeningSize,
                    MaxOpeningSize = _configurationSettings.MaxOpeningSize,
                    IncludeInsulation = _configurationSettings.IncludeInsulation,
                    AutoUpdateOpenings = _configurationSettings.AutoUpdateOpenings
                };

                // Save to filter (this will be implemented in Step 3)
                _filterManagementService.SaveOpeningConditions(openingConditions);
                _filterManagementService.SaveConfigurationSettings(configurationSettings);

                // Mark as saved
                _openingConditions.MarkAsSaved();
                _configurationSettings.MarkAsSaved();

                DebugLogger.Log("UI State saved to filter successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Failed to save UI state to filter: {ex.Message}");
                throw;
            }
        }

        public void LoadFromFilter()
        {
            try
            {
                // Load from filter (this will be implemented in Step 3)
                var openingConditions = _filterManagementService.LoadOpeningConditions();
                var configurationSettings = _filterManagementService.LoadConfigurationSettings();

                if (openingConditions != null)
                {
                    _openingConditions.DuctClearance = openingConditions.DuctClearance;
                    _openingConditions.PipeClearance = openingConditions.PipeClearance;
                    _openingConditions.CableTrayClearance = openingConditions.CableTrayClearance;
                    _openingConditions.OversizeValue = openingConditions.OversizeValue;
                    _openingConditions.LevelAdjustment = openingConditions.LevelAdjustment;
                }

                if (configurationSettings != null)
                {
                    _configurationSettings.MinOpeningSize = configurationSettings.MinOpeningSize;
                    _configurationSettings.MaxOpeningSize = configurationSettings.MaxOpeningSize;
                    _configurationSettings.IncludeInsulation = configurationSettings.IncludeInsulation;
                    _configurationSettings.AutoUpdateOpenings = configurationSettings.AutoUpdateOpenings;
                }

                // Mark as saved (no modifications yet)
                _openingConditions.MarkAsSaved();
                _configurationSettings.MarkAsSaved();

                DebugLogger.Log("UI State loaded from filter successfully");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Failed to load UI state from filter: {ex.Message}");
                throw;
            }
        }

        public void ResetToDefaults()
        {
            _openingConditions.DuctClearance = 50.0;
            _openingConditions.PipeClearance = 75.0;
            _openingConditions.CableTrayClearance = 25.0;
            _openingConditions.OversizeValue = 0.0;
            _openingConditions.LevelAdjustment = 0.0;

            _configurationSettings.MinOpeningSize = 0.1;
            _configurationSettings.MaxOpeningSize = 1000.0;
            _configurationSettings.IncludeInsulation = true;
            _configurationSettings.AutoUpdateOpenings = false;

            DebugLogger.Log("UI State reset to defaults");
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _autoSaveTimer?.Dispose();
                _disposed = true;
            }
        }
    }

    // Event argument classes
    public class OpeningConditionsChangedEventArgs : EventArgs
    {
        public OpeningConditionsState OpeningConditions { get; }
        public DateTime Timestamp { get; }

        public OpeningConditionsChangedEventArgs(OpeningConditionsState openingConditions)
        {
            OpeningConditions = openingConditions;
            Timestamp = DateTime.Now;
        }
    }

    public class ConfigurationSettingsChangedEventArgs : EventArgs
    {
        public ConfigurationSettingsState ConfigurationSettings { get; }
        public DateTime Timestamp { get; }

        public ConfigurationSettingsChangedEventArgs(ConfigurationSettingsState configurationSettings)
        {
            ConfigurationSettings = configurationSettings;
            Timestamp = DateTime.Now;
        }
    }

    public class AutoSaveEventArgs : EventArgs
    {
        public DateTime Timestamp { get; }

        public AutoSaveEventArgs()
        {
            Timestamp = DateTime.Now;
        }
    }
}
```

**Testing Step 2:**
1. Create the file `Services/UIStateManagerService.cs`
2. Build the project to ensure no compilation errors
3. Create a test console application to test the service
4. Test property change notifications
5. Test auto-save timer functionality
6. Test save/load operations (will fail until Step 3)

**Expected Result:** Service compiles successfully and property change notifications work.

---

### **STEP 4: Create Project Filter Configuration**

**Objective**: Create project-specific filter configuration that can override global defaults

**Files to Create:**
- `Models/ProjectFilterConfiguration.cs`
- `Services/ProjectFilterConfigurationService.cs`

**Implementation:**

```csharp
// Models/ProjectFilterConfiguration.cs
using System;
using System.Xml.Serialization;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Project-specific filter configuration that can override global defaults
    /// Based on CONVOID's Extensible Storage approach
    /// </summary>
    [XmlRoot("ProjectFilterConfiguration")]
    public class ProjectFilterConfiguration
    {
        // Filter Identity
        public string FilterName { get; set; } = "";
        public string ProjectName { get; set; } = "";
        public string ProjectPath { get; set; } = "";
        
        // File Selections (Project-Specific)
        public List<string> ReferenceFiles { get; set; } = new List<string>();
        public List<string> HostFiles { get; set; } = new List<string>();
        public List<string> LinkedModels { get; set; } = new List<string>();
        
        // Opening Conditions (Project-Specific Overrides)
        public double? DuctClearance { get; set; } = null; // null = use global default
        public double? PipeClearance { get; set; } = null;
        public double? CableTrayClearance { get; set; } = null;
        public double? OversizeValue { get; set; } = null;
        public double? LevelAdjustment { get; set; } = null;
        
        // Configuration Settings (Project-Specific Overrides)
        public double? MinOpeningSize { get; set; } = null;
        public double? MaxOpeningSize { get; set; } = null;
        public bool? IncludeInsulation { get; set; } = null;
        public bool? AutoUpdateOpenings { get; set; } = null;
        
        // CONFIGURATION Settings (Project-Specific Overrides)
        // LIMITS Section Overrides (REQUIRED - All Green Highlighted)
        public double? MinOpeningSize { get; set; } = null; // Ignore openings smaller than
        public double? MaxCircularDiameter { get; set; } = null; // Round openings become rectangular if diameter is greater than
        public double? JoinOpeningsDistance { get; set; } = null; // Join openings at a distance of
        public double? MaxOpeningAngle { get; set; } = null; // Ignore openings with an angle greater than
        public bool? CreateSlopedOpenings { get; set; } = null; // Create openings with a slope
        public bool? RoundUpDimensions { get; set; } = null; // Round up opening dimensions (Do not round up)
        public double? RoundingIncrement { get; set; } = null; // Rounding increment
        
        // ELEMENT FILTER Section Overrides (REQUIRED - Green Highlighted)
        public string? ElementFilterView { get; set; } = null; // Element Filter dropdown selection
        
        // MANAGE Section Overrides (REQUIRED - Important for workflow)
        public bool? ResetApprovalStatusOnChanges { get; set; } = null; // Reset approval status of openings when changes occur
        public double? DimensionChangeThreshold { get; set; } = null; // Openings won't be marked as changed if change in dimensions is less than (mm)
        public double? LocationChangeThreshold { get; set; } = null; // Openings won't be marked as changed if change in location is less than (mm)
        
        // ELEMENTS Section (REMOVED - Red X Marks)
        // CreateVerticalParallelOpenings - REMOVED (Red X)
        // CreateHorizontalParallelOpenings - REMOVED (Red X)
        // CutOpeningWithHosts - REMOVED (Not highlighted)
        // CreateConstraintBetweenOpeningsAndHosts - REMOVED (Not highlighted)
        
        // BCF Data (Project-Specific)
        public BcfProjectData? BcfData { get; set; } = null;
        
        // Metadata
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public DateTime LastModified { get; set; } = DateTime.Now;
        public string CreatedBy { get; set; } = "";
        public string LastModifiedBy { get; set; } = "";
    }
}
```

```csharp
// Services/ProjectFilterConfigurationService.cs
using System;
using System.IO;
using System.Xml.Serialization;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for managing project-specific filter configurations
    /// Based on CONVOID's Extensible Storage approach
    /// </summary>
    public class ProjectFilterConfigurationService
    {
        private readonly GlobalConfigurationManager _globalConfig;
        private ProjectFilterConfiguration? _currentFilter;
        
        public ProjectFilterConfigurationService()
        {
            _globalConfig = GlobalConfigurationManager.Instance;
        }
        
        public ProjectFilterConfiguration? CurrentFilter => _currentFilter;
        
        public void CreateNewFilter(string filterName, string projectName, string projectPath)
        {
            _currentFilter = new ProjectFilterConfiguration
            {
                FilterName = filterName,
                ProjectName = projectName,
                ProjectPath = projectPath,
                CreatedBy = _globalConfig.Settings.DefaultUserName,
                LastModifiedBy = _globalConfig.Settings.DefaultUserName
            };
            
            DebugLogger.Log($"New project filter created: {filterName}");
        }
        
        public void LoadFilter(string filterName)
        {
            try
            {
                var filterPath = GetFilterFilePath(filterName);
                if (File.Exists(filterPath))
                {
                    var serializer = new XmlSerializer(typeof(ProjectFilterConfiguration));
                    using (var reader = new StreamReader(filterPath))
                    {
                        _currentFilter = (ProjectFilterConfiguration)serializer.Deserialize(reader)!;
                    }
                    DebugLogger.Log($"Project filter loaded: {filterName}");
                }
                else
                {
                    DebugLogger.Log($"Project filter not found: {filterName}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Failed to load project filter: {ex.Message}");
                throw;
            }
        }
        
        public void SaveFilter()
        {
            if (_currentFilter == null)
            {
                throw new InvalidOperationException("No current filter to save");
            }
            
            try
            {
                var filterPath = GetFilterFilePath(_currentFilter.FilterName);
                var directory = Path.GetDirectoryName(filterPath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);
                
                _currentFilter.LastModified = DateTime.Now;
                _currentFilter.LastModifiedBy = _globalConfig.Settings.DefaultUserName;
                
                var serializer = new XmlSerializer(typeof(ProjectFilterConfiguration));
                using (var writer = new StreamWriter(filterPath))
                {
                    serializer.Serialize(writer, _currentFilter);
                }
                
                DebugLogger.Log($"Project filter saved: {_currentFilter.FilterName}");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Failed to save project filter: {ex.Message}");
                throw;
            }
        }
        
        public OpeningConditions GetEffectiveOpeningConditions()
        {
            var global = _globalConfig.Settings;
            var filter = _currentFilter;
            
            return new OpeningConditions
            {
                DuctClearance = filter?.DuctClearance ?? global.DefaultDuctClearance,
                PipeClearance = filter?.PipeClearance ?? global.DefaultPipeClearance,
                CableTrayClearance = filter?.CableTrayClearance ?? global.DefaultCableTrayClearance,
                OversizeValue = filter?.OversizeValue ?? 0.0,
                LevelAdjustment = filter?.LevelAdjustment ?? 0.0
            };
        }
        
        public ConfigurationSettings GetEffectiveConfigurationSettings()
        {
            var global = _globalConfig.Settings;
            var filter = _currentFilter;
            
            return new ConfigurationSettings
            {
                // Basic Settings
                MinOpeningSize = filter?.MinOpeningSize ?? global.DefaultMinOpeningSize,
                MaxOpeningSize = filter?.MaxOpeningSize ?? global.DefaultMaxOpeningSize,
                IncludeInsulation = filter?.IncludeInsulation ?? true,
                AutoUpdateOpenings = filter?.AutoUpdateOpenings ?? false,
                
                // CONFIGURATION Settings (LIMITS Section - REQUIRED)
                MinOpeningSize = filter?.MinOpeningSize ?? global.MinOpeningSize,
                MaxCircularDiameter = filter?.MaxCircularDiameter ?? global.MaxCircularDiameter,
                JoinOpeningsDistance = filter?.JoinOpeningsDistance ?? global.JoinOpeningsDistance,
                MaxOpeningAngle = filter?.MaxOpeningAngle ?? global.MaxOpeningAngle,
                CreateSlopedOpenings = filter?.CreateSlopedOpenings ?? global.CreateSlopedOpenings,
                RoundUpDimensions = filter?.RoundUpDimensions ?? global.RoundUpDimensions,
                RoundingIncrement = filter?.RoundingIncrement ?? global.RoundingIncrement,
                
                // CONFIGURATION Settings (ELEMENT FILTER Section - REQUIRED)
                ElementFilterView = filter?.ElementFilterView ?? global.ElementFilterView,
                
                // CONFIGURATION Settings (MANAGE Section - REQUIRED)
                ResetApprovalStatusOnChanges = filter?.ResetApprovalStatusOnChanges ?? global.ResetApprovalStatusOnChanges,
                DimensionChangeThreshold = filter?.DimensionChangeThreshold ?? global.DimensionChangeThreshold,
                LocationChangeThreshold = filter?.LocationChangeThreshold ?? global.LocationChangeThreshold
                
                // CONFIGURATION Settings (ELEMENTS Section - REMOVED)
                // CreateVerticalParallelOpenings - REMOVED (Red X)
                // CreateHorizontalParallelOpenings - REMOVED (Red X)
                // CutOpeningWithHosts - REMOVED (Not highlighted)
                // CreateConstraintBetweenOpeningsAndHosts - REMOVED (Not highlighted)
            };
        }
        
        private string GetFilterFilePath(string filterName)
        {
            // Project Filter Location: C:\Users\[Username]\AppData\Roaming\JSE_MEP_Openings\Projects\[ProjectName]\Filters\[FilterName].xml
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var projectPath = Path.Combine(appDataPath, "JSE_MEP_Openings", "Projects", GetCurrentProjectName());
            var filtersPath = Path.Combine(projectPath, "Filters");
            Directory.CreateDirectory(filtersPath); // Ensure directory exists
            return Path.Combine(filtersPath, $"{filterName}.xml");
        }
        
        private string GetCurrentProjectName()
        {
            // Get current Revit project name or use default
            // This should be implemented to get the actual project name from Revit
            return "DefaultProject"; // Placeholder - implement based on Revit document name
        }
    }
}
```

**Testing Step 4:**
1. Create the files `Models/ProjectFilterConfiguration.cs` and `Services/ProjectFilterConfigurationService.cs`
2. Build the project to ensure no compilation errors
3. Test creating a new filter
4. Test loading an existing filter
5. Test saving a filter
6. Test effective settings calculation (global defaults + filter overrides)

**Expected Result:** Project filter configuration system compiles successfully and can manage project-specific settings.

---

### **STEP 5: Extend Filter Management Service**

**Objective**: Add methods to FilterManagementService to handle opening conditions and configuration settings

**Files to Modify:**
- `Services/FilterManagementService.cs`

**Implementation:**

```csharp
// Add these methods to existing FilterManagementService.cs

public class FilterManagementService
{
    // ... existing code ...

    /// <summary>
    /// Save opening conditions to the current filter
    /// </summary>
    public void SaveOpeningConditions(OpeningConditions openingConditions)
    {
        try
        {
            var currentFilter = GetCurrentFilter();
            if (currentFilter == null)
            {
                throw new InvalidOperationException("No current filter selected");
            }

            // Create or update opening conditions in filter
            currentFilter.OpeningConditions = openingConditions;
            currentFilter.LastModified = DateTime.Now;

            // Save to file
            SaveFilterToFile(currentFilter);

            DebugLogger.Log($"Opening conditions saved to filter: {currentFilter.Name}");
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"Failed to save opening conditions: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Load opening conditions from the current filter
    /// </summary>
    public OpeningConditions? LoadOpeningConditions()
    {
        try
        {
            var currentFilter = GetCurrentFilter();
            if (currentFilter?.OpeningConditions == null)
            {
                return null;
            }

            DebugLogger.Log($"Opening conditions loaded from filter: {currentFilter.Name}");
            return currentFilter.OpeningConditions;
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"Failed to load opening conditions: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Save configuration settings to the current filter
    /// </summary>
    public void SaveConfigurationSettings(ConfigurationSettings configurationSettings)
    {
        try
        {
            var currentFilter = GetCurrentFilter();
            if (currentFilter == null)
            {
                throw new InvalidOperationException("No current filter selected");
            }

            // Create or update configuration settings in filter
            currentFilter.ConfigurationSettings = configurationSettings;
            currentFilter.LastModified = DateTime.Now;

            // Save to file
            SaveFilterToFile(currentFilter);

            DebugLogger.Log($"Configuration settings saved to filter: {currentFilter.Name}");
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"Failed to save configuration settings: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Load configuration settings from the current filter
    /// </summary>
    public ConfigurationSettings? LoadConfigurationSettings()
    {
        try
        {
            var currentFilter = GetCurrentFilter();
            if (currentFilter?.ConfigurationSettings == null)
            {
                return null;
            }

            DebugLogger.Log($"Configuration settings loaded from filter: {currentFilter.Name}");
            return currentFilter.ConfigurationSettings;
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"Failed to load configuration settings: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Get the currently selected filter
    /// </summary>
    private OpeningFilter? GetCurrentFilter()
    {
        // This will be implemented based on your current filter selection logic
        // For now, return a default filter for testing
        return new OpeningFilter
        {
            Name = "Default Filter",
            Category = MepCategory.Ducts,
            OpeningType = OpeningType.RectangularSleeves,
            IsEnabled = true
        };
    }

    /// <summary>
    /// Save filter to file
    /// </summary>
    private void SaveFilterToFile(OpeningFilter filter)
    {
        try
        {
            var filterPath = GetFilterFilePath(filter.Name);
            var directory = Path.GetDirectoryName(filterPath);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Serialize filter to XML
            var serializer = new XmlSerializer(typeof(OpeningFilter));
            using (var writer = new StreamWriter(filterPath))
            {
                serializer.Serialize(writer, filter);
            }

            DebugLogger.Log($"Filter saved to file: {filterPath}");
        }
        catch (Exception ex)
        {
            DebugLogger.Log($"Failed to save filter to file: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Get the file path for a filter
    /// </summary>
    private string GetFilterFilePath(string filterName)
    {
        var documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var filtersPath = Path.Combine(documentsPath, "JSE_MEP_Openings", "Filters");
        return Path.Combine(filtersPath, $"{filterName}.xml");
    }
}
```

**Testing Step 3:**
1. Modify `Services/FilterManagementService.cs` to add the new methods
2. Build the project to ensure no compilation errors
3. Test the save/load operations from Step 2
4. Verify that files are created in the correct location
5. Verify that XML serialization works correctly

**Expected Result:** Filter management service can save and load opening conditions and configuration settings.

---

### **STEP 4: Create Test Console Application**

**Objective**: Create a test application to verify the hybrid approach works correctly

**Files to Create:**
- `TestHybridApproach.cs`

**Implementation:**

```csharp
// TestHybridApproach.cs
using System;
using System.Threading;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS
{
    /// <summary>
    /// Test console application for the hybrid approach
    /// </summary>
    public class TestHybridApproach
    {
        private static UIStateManagerService? _uiStateManager;
        private static FilterManagementService? _filterManagementService;

        public static void Main(string[] args)
        {
            Console.WriteLine("=== Hybrid Approach Test ===");
            Console.WriteLine();

            try
            {
                // Initialize services
                InitializeServices();

                // Test 1: Property Change Notifications
                TestPropertyChangeNotifications();

                // Test 2: Auto-save Functionality
                TestAutoSaveFunctionality();

                // Test 3: Manual Save/Load
                TestManualSaveLoad();

                // Test 4: Reset to Defaults
                TestResetToDefaults();

                Console.WriteLine("=== All Tests Passed! ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
            finally
            {
                _uiStateManager?.Dispose();
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static void InitializeServices()
        {
            Console.WriteLine("Initializing services...");
            
            _filterManagementService = new FilterManagementService();
            _uiStateManager = new UIStateManagerService(_filterManagementService);

            // Subscribe to events
            _uiStateManager.OpeningConditionsChanged += OnOpeningConditionsChanged;
            _uiStateManager.ConfigurationSettingsChanged += OnConfigurationSettingsChanged;
            _uiStateManager.AutoSaveTriggered += OnAutoSaveTriggered;

            Console.WriteLine("Services initialized successfully.");
            Console.WriteLine();
        }

        private static void TestPropertyChangeNotifications()
        {
            Console.WriteLine("=== Test 1: Property Change Notifications ===");
            
            // Test opening conditions
            Console.WriteLine("Testing opening conditions...");
            _uiStateManager!.OpeningConditions.DuctClearance = 100.0;
            Console.WriteLine($"Duct clearance set to: {_uiStateManager.OpeningConditions.DuctClearance}");
            Console.WriteLine($"Is modified: {_uiStateManager.OpeningConditions.IsModified}");
            
            // Test configuration settings
            Console.WriteLine("Testing configuration settings...");
            _uiStateManager.ConfigurationSettings.MinOpeningSize = 0.5;
            Console.WriteLine($"Min opening size set to: {_uiStateManager.ConfigurationSettings.MinOpeningSize}");
            Console.WriteLine($"Is modified: {_uiStateManager.ConfigurationSettings.IsModified}");
            
            Console.WriteLine("Property change notifications test completed.");
            Console.WriteLine();
        }

        private static void TestAutoSaveFunctionality()
        {
            Console.WriteLine("=== Test 2: Auto-save Functionality ===");
            
            // Make changes to trigger auto-save
            _uiStateManager!.OpeningConditions.PipeClearance = 125.0;
            _uiStateManager.ConfigurationSettings.MaxOpeningSize = 2000.0;
            
            Console.WriteLine("Changes made, waiting for auto-save...");
            Console.WriteLine("(Auto-save should trigger after 30 seconds)");
            
            // Wait for auto-save (in real test, you might want to reduce timer)
            Thread.Sleep(35000);
            
            Console.WriteLine("Auto-save test completed.");
            Console.WriteLine();
        }

        private static void TestManualSaveLoad()
        {
            Console.WriteLine("=== Test 3: Manual Save/Load ===");
            
            // Make changes
            _uiStateManager!.OpeningConditions.CableTrayClearance = 50.0;
            _uiStateManager.ConfigurationSettings.IncludeInsulation = false;
            
            Console.WriteLine("Changes made:");
            Console.WriteLine($"Cable tray clearance: {_uiStateManager.OpeningConditions.CableTrayClearance}");
            Console.WriteLine($"Include insulation: {_uiStateManager.ConfigurationSettings.IncludeInsulation}");
            
            // Save manually
            Console.WriteLine("Saving to filter...");
            _uiStateManager.SaveToFilter();
            Console.WriteLine("Save completed.");
            
            // Reset values
            _uiStateManager.OpeningConditions.CableTrayClearance = 25.0;
            _uiStateManager.ConfigurationSettings.IncludeInsulation = true;
            
            Console.WriteLine("Values reset:");
            Console.WriteLine($"Cable tray clearance: {_uiStateManager.OpeningConditions.CableTrayClearance}");
            Console.WriteLine($"Include insulation: {_uiStateManager.ConfigurationSettings.IncludeInsulation}");
            
            // Load from filter
            Console.WriteLine("Loading from filter...");
            _uiStateManager.LoadFromFilter();
            Console.WriteLine("Load completed.");
            
            Console.WriteLine("Values after load:");
            Console.WriteLine($"Cable tray clearance: {_uiStateManager.OpeningConditions.CableTrayClearance}");
            Console.WriteLine($"Include insulation: {_uiStateManager.ConfigurationSettings.IncludeInsulation}");
            
            Console.WriteLine("Manual save/load test completed.");
            Console.WriteLine();
        }

        private static void TestResetToDefaults()
        {
            Console.WriteLine("=== Test 4: Reset to Defaults ===");
            
            // Make changes
            _uiStateManager!.OpeningConditions.DuctClearance = 200.0;
            _uiStateManager.ConfigurationSettings.MinOpeningSize = 1.0;
            
            Console.WriteLine("Values before reset:");
            Console.WriteLine($"Duct clearance: {_uiStateManager.OpeningConditions.DuctClearance}");
            Console.WriteLine($"Min opening size: {_uiStateManager.ConfigurationSettings.MinOpeningSize}");
            
            // Reset to defaults
            _uiStateManager.ResetToDefaults();
            
            Console.WriteLine("Values after reset:");
            Console.WriteLine($"Duct clearance: {_uiStateManager.OpeningConditions.DuctClearance}");
            Console.WriteLine($"Min opening size: {_uiStateManager.ConfigurationSettings.MinOpeningSize}");
            
            Console.WriteLine("Reset to defaults test completed.");
            Console.WriteLine();
        }

        private static void OnOpeningConditionsChanged(object? sender, OpeningConditionsChangedEventArgs e)
        {
            Console.WriteLine($"[EVENT] Opening conditions changed at {e.Timestamp:HH:mm:ss}");
        }

        private static void OnConfigurationSettingsChanged(object? sender, ConfigurationSettingsChangedEventArgs e)
        {
            Console.WriteLine($"[EVENT] Configuration settings changed at {e.Timestamp:HH:mm:ss}");
        }

        private static void OnAutoSaveTriggered(object? sender, AutoSaveEventArgs e)
        {
            Console.WriteLine($"[EVENT] Auto-save triggered at {e.Timestamp:HH:mm:ss}");
        }
    }
}
```

**Testing Step 4:**
1. Create the file `TestHybridApproach.cs`
2. Build the project to ensure no compilation errors
3. Run the test application
4. Verify that all tests pass
5. Check that files are created in the correct location
6. Verify that auto-save works (wait 30 seconds)

**Expected Result:** Test application runs successfully and all functionality works as expected.

---

### **STEP 5: Integration with Existing UI**

**Objective**: Integrate the hybrid approach with the existing EmergencyMainDialog

**Files to Modify:**
- `Views/EmergencyMainDialog.cs`

**Implementation:**

```csharp
// Add these fields to EmergencyMainDialog class
private UIStateManagerService? _uiStateManager;
private bool _isInitialized = false;

// Add this method to initialize the UI State Manager
private void InitializeUIStateManager()
{
    try
    {
        if (_uiStateManager == null)
        {
            _uiStateManager = new UIStateManagerService(_filterManagementService);
            
            // Subscribe to events
            _uiStateManager.OpeningConditionsChanged += OnOpeningConditionsChanged;
            _uiStateManager.ConfigurationSettingsChanged += OnConfigurationSettingsChanged;
            _uiStateManager.AutoSaveTriggered += OnAutoSaveTriggered;
            
            // Load initial state from filter
            _uiStateManager.LoadFromFilter();
            
            // Bind UI controls to state
            BindUIToState();
            
            _isInitialized = true;
            
            DebugLogger.Log("UI State Manager initialized successfully");
        }
    }
    catch (Exception ex)
    {
        DebugLogger.Log($"Failed to initialize UI State Manager: {ex.Message}");
        throw;
    }
}

// Add this method to bind UI controls to state
private void BindUIToState()
{
    if (_uiStateManager == null) return;

    // Bind opening conditions
    _uiStateManager.OpeningConditions.PropertyChanged += (s, e) =>
    {
        if (e.PropertyName == nameof(OpeningConditionsState.DuctClearance))
        {
            // Update UI control (assuming you have a textbox for duct clearance)
            // ductClearanceTextBox.Text = _uiStateManager.OpeningConditions.DuctClearance.ToString();
        }
        // Add similar bindings for other properties
    };

    // Bind configuration settings
    _uiStateManager.ConfigurationSettings.PropertyChanged += (s, e) =>
    {
        if (e.PropertyName == nameof(ConfigurationSettingsState.MinOpeningSize))
        {
            // Update UI control (assuming you have a textbox for min opening size)
            // minOpeningSizeTextBox.Text = _uiStateManager.ConfigurationSettings.MinOpeningSize.ToString();
        }
        // Add similar bindings for other properties
    };
}

// Add these event handlers
private void OnOpeningConditionsChanged(object? sender, OpeningConditionsChangedEventArgs e)
{
    // Update UI to show that changes have been made
    UpdateUIForChanges();
}

private void OnConfigurationSettingsChanged(object? sender, ConfigurationSettingsChangedEventArgs e)
{
    // Update UI to show that changes have been made
    UpdateUIForChanges();
}

private void OnAutoSaveTriggered(object? sender, AutoSaveEventArgs e)
{
    // Show auto-save notification to user
    ShowAutoSaveNotification();
}

private void UpdateUIForChanges()
{
    if (_uiStateManager?.HasUnsavedChanges == true)
    {
        // Update UI to show unsaved changes (e.g., change button colors, show asterisk, etc.)
        // This is where you'd update the visual state of your UI
    }
}

private void ShowAutoSaveNotification()
{
    // Show a brief notification that auto-save occurred
    // This could be a tooltip, status bar message, or toast notification
}

// Modify the constructor to initialize the UI State Manager
public EmergencyMainDialog()
{
    InitializeComponent();
    
    // ... existing initialization code ...
    
    // Initialize UI State Manager
    InitializeUIStateManager();
}

// Add methods for manual save/load
private void OnSaveSettingsClick(object? sender, EventArgs e)
{
    try
    {
        _uiStateManager?.SaveToFilter();
        ShowNotification("Settings saved successfully");
    }
    catch (Exception ex)
    {
        ShowError($"Failed to save settings: {ex.Message}");
    }
}

private void OnLoadSettingsClick(object? sender, EventArgs e)
{
    try
    {
        _uiStateManager?.LoadFromFilter();
        ShowNotification("Settings loaded successfully");
    }
    catch (Exception ex)
    {
        ShowError($"Failed to load settings: {ex.Message}");
    }
}

private void OnResetSettingsClick(object? sender, EventArgs e)
{
    try
    {
        _uiStateManager?.ResetToDefaults();
        ShowNotification("Settings reset to defaults");
    }
    catch (Exception ex)
    {
        ShowError($"Failed to reset settings: {ex.Message}");
    }
}

// Add cleanup in Dispose method
protected override void Dispose(bool disposing)
{
    if (disposing)
    {
        _uiStateManager?.Dispose();
    }
    base.Dispose(disposing);
}
```

**Testing Step 5:**
1. Modify `Views/EmergencyMainDialog.cs` to add the UI State Manager integration
2. Build the project to ensure no compilation errors
3. Test the integration by running the application
4. Verify that property changes update the UI
5. Verify that auto-save works
6. Test manual save/load operations

**Expected Result:** UI integration works correctly and the hybrid approach is fully functional.

---

## 🎯 **Summary of Implementation Steps (Three-Tier Configuration System)**

1. **Step 1**: Create Global Application Configuration - Test global settings management
2. **Step 2**: Create UI State Manager Models - Test property change notifications
3. **Step 3**: Create UI State Manager Service - Test service functionality with global defaults
4. **Step 4**: Create Project Filter Configuration - Test project-specific settings management
5. **Step 5**: Extend Filter Management Service - Test save/load operations
6. **Step 6**: Create Test Console Application - Test complete three-tier functionality
7. **Step 7**: Integration with Existing UI - Test UI integration

## ✅ **Benefits of Three-Tier Configuration System**

### **🌍 Global Application Configuration:**
- **Filter-independent**: Applies to all projects
- **User-wide**: Same settings across all projects
- **Occasionally changed**: Set once, rarely modified
- **Application-level**: Paths, preferences, defaults

### **📁 Project Filter Configuration:**
- **Project-specific**: Different for each project
- **Filter-dependent**: Tied to specific filters
- **Frequently changed**: Modified per project needs
- **FILTER**: File selections, category selections, basic opening conditions
- **CONFIGURATION**: Advanced processing rules (Limits, Elements, etc.)

### **⚡ UI State Manager:**
- **Session-based**: Temporary changes
- **Dynamic**: Real-time updates
- **Auto-saved**: Automatically persisted
- **User-level**: Live experimentation

### **🎯 Overall Benefits:**
- **Dynamic Changes**: Users can experiment with settings
- **Real-time Preview**: Immediate feedback on changes
- **Auto-save**: Changes are automatically persisted
- **Manual Control**: Users can save/load specific configurations
- **Performance**: No constant file writes
- **Flexibility**: Best of both worlds
- **CONVOID Compliance**: Matches official architecture

## 🚀 **Next Steps After Implementation**

1. **Test each step independently** before proceeding to the next
2. **Verify functionality** at each step
3. **Fix any issues** before moving forward
4. **Document any changes** or modifications made
5. **Prepare for UI integration** in Step 5

This step-by-step approach ensures that each component works correctly before integration, making debugging much easier and ensuring a robust implementation.
