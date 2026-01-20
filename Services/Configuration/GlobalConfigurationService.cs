using System;
using System.Collections.Generic;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Configuration
{
    /// <summary>
    /// Service for accessing global application configuration
    /// Provides clean interface to GlobalConfigurationManager and handles application-wide settings
    /// </summary>
    public class GlobalConfigurationService
    {
        private static GlobalConfigurationService _instance;
        private static readonly object _lock = new object();
        
        private readonly GlobalConfigurationManager _configManager;
        
        private GlobalConfigurationService()
        {
            _configManager = GlobalConfigurationManager.Instance;
            DebugLogger.Info("[GlobalConfigService] Initialized");
        }
        
        public static GlobalConfigurationService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new GlobalConfigurationService();
                        }
                    }
                }
                return _instance;
            }
        }
        
        #region Path Configuration
        
        /// <summary>
        /// Get default families path for opening families
        /// </summary>
        public string GetDefaultFamiliesPath()
        {
            return _configManager.Settings.DefaultFamiliesPath;
        }
        
        /// <summary>
        /// Get default BCF path for clash detection
        /// </summary>
        public string GetDefaultBCFPath()
        {
            return _configManager.Settings.DefaultBCFPath;
        }
        
        /// <summary>
        /// Get default filters path for project filters
        /// </summary>
        public string GetDefaultFiltersPath()
        {
            return _configManager.Settings.DefaultFiltersPath;
        }
        
        /// <summary>
        /// Get default projects path
        /// </summary>
        public string GetDefaultProjectsPath()
        {
            return _configManager.Settings.DefaultProjectsPath;
        }
        
        #endregion
        
        #region Application Preferences
        
        /// <summary>
        /// Get UI language setting
        /// </summary>
        public string GetUILanguage()
        {
            return _configManager.Settings.UILanguage;
        }
        
        /// <summary>
        /// Check if auto-save is enabled
        /// </summary>
        public bool IsAutoSaveEnabled()
        {
            return _configManager.Settings.AutoSaveEnabled;
        }
        
        /// <summary>
        /// Get auto-save interval in seconds
        /// </summary>
        public int GetAutoSaveIntervalSeconds()
        {
            return _configManager.Settings.AutoSaveIntervalSeconds;
        }
        
        /// <summary>
        /// Check if debug logs should be shown
        /// </summary>
        public bool ShouldShowDebugLogs()
        {
            return _configManager.Settings.ShowDebugLogs;
        }
        
        #endregion
        
        #region Configuration Defaults
        
        /// <summary>
        /// Get opening size limits
        /// </summary>
        public OpeningSizeLimits GetOpeningSizeLimits()
        {
            var defaults = _configManager.Settings.ConfigurationDefaults;
            return new OpeningSizeLimits
            {
                MaxWidth = defaults.MaxOpeningWidth,
                MaxHeight = defaults.MaxOpeningHeight,
                MaxDepth = defaults.MaxOpeningDepth,
                MinWidth = defaults.MinOpeningWidth,
                MinHeight = defaults.MinOpeningHeight,
                MinDepth = defaults.MinOpeningDepth
            };
        }
        
        /// <summary>
        /// Get element processing flags
        /// </summary>
        public ElementProcessingFlags GetElementProcessingFlags()
        {
            var defaults = _configManager.Settings.ConfigurationDefaults;
            return new ElementProcessingFlags
            {
                ProcessDucts = defaults.ProcessDucts,
                ProcessPipes = defaults.ProcessPipes,
                ProcessCableTrays = defaults.ProcessCableTrays,
                ProcessConduits = defaults.ProcessConduits,
                ProcessDuctInsulation = defaults.ProcessDuctInsulation,
                ProcessPipeInsulation = defaults.ProcessPipeInsulation
            };
        }
        
        /// <summary>
        /// Get management settings
        /// </summary>
        public ManagementSettings GetManagementSettings()
        {
            var defaults = _configManager.Settings.ConfigurationDefaults;
            return new ManagementSettings
            {
                ResetApprovalStatusOnChanges = defaults.ResetApprovalStatusOnChanges,
                DimensionChangeThreshold = defaults.DimensionChangeThreshold,
                LocationChangeThreshold = defaults.LocationChangeThreshold,
                CutOpeningWithHosts = defaults.CutOpeningWithHosts,
                CreateConstraintBetweenOpeningsAndHosts = defaults.CreateConstraintBetweenOpeningsAndHosts,
                AdoptProvisionForVoidsFromLinkedModel = defaults.AdoptProvisionForVoidsFromLinkedModel,
                IncludeHostElementsNotVisibleIn3DView = defaults.IncludeHostElementsNotVisibleIn3DView,
                IncludeReferenceElementsNotVisibleIn3DView = defaults.IncludeReferenceElementsNotVisibleIn3DView,
                IncludeHostElementsInDemolishedPhase = defaults.IncludeHostElementsInDemolishedPhase
            };
        }
        
        /// <summary>
        /// Get processing limits
        /// </summary>
        public ProcessingLimits GetProcessingLimits()
        {
            var defaults = _configManager.Settings.ConfigurationDefaults;
            return new ProcessingLimits
            {
                IgnoreOpeningsSmallerThan = defaults.IgnoreOpeningsSmallerThan,
                RoundOpeningsBecomeRectangularIfDiameterGreaterThan = defaults.RoundOpeningsBecomeRectangularIfDiameterGreaterThan,
                JoinOpeningsIfDistanceLessThan = defaults.JoinOpeningsIfDistanceLessThan,
                IgnoreOpeningsWithAngleGreaterThan = defaults.IgnoreOpeningsWithAngleGreaterThan
            };
        }
        
        #endregion
        
        #region Convenience Methods
        
        /// <summary>
        /// Check if a specific element type should be processed
        /// </summary>
        public bool ShouldProcessElementType(string elementType)
        {
            var flags = GetElementProcessingFlags();
            
            return elementType.ToLower() switch
            {
                "duct" or "ducts" => flags.ProcessDucts,
                "pipe" or "pipes" => flags.ProcessPipes,
                "cable tray" or "cable trays" => flags.ProcessCableTrays,
                "conduit" or "conduits" => flags.ProcessConduits,
                _ => true // Default to process if not specified
            };
        }
        
        /// <summary>
        /// Get all configuration as a dictionary for easy access
        /// </summary>
        public Dictionary<string, object> GetAllConfiguration()
        {
            return new Dictionary<string, object>
            {
                ["OpeningSizeLimits"] = GetOpeningSizeLimits(),
                ["ElementProcessingFlags"] = GetElementProcessingFlags(),
                ["ManagementSettings"] = GetManagementSettings(),
                ["ProcessingLimits"] = GetProcessingLimits(),
                ["Paths"] = new
                {
                    FamiliesPath = GetDefaultFamiliesPath(),
                    BCFPath = GetDefaultBCFPath(),
                    FiltersPath = GetDefaultFiltersPath(),
                    ProjectsPath = GetDefaultProjectsPath()
                },
                ["Preferences"] = new
                {
                    UILanguage = GetUILanguage(),
                    AutoSaveEnabled = IsAutoSaveEnabled(),
                    AutoSaveIntervalSeconds = GetAutoSaveIntervalSeconds(),
                    ShowDebugLogs = ShouldShowDebugLogs()
                }
            };
        }
        
        #endregion
        
        #region Helper Classes
        
        /// <summary>
        /// Opening size limits configuration
        /// </summary>
        public class OpeningSizeLimits
        {
            public double MaxWidth { get; set; }
            public double MaxHeight { get; set; }
            public double MaxDepth { get; set; }
            public double MinWidth { get; set; }
            public double MinHeight { get; set; }
            public double MinDepth { get; set; }
        }
        
        /// <summary>
        /// Element processing flags configuration
        /// </summary>
        public class ElementProcessingFlags
        {
            public bool ProcessDucts { get; set; }
            public bool ProcessPipes { get; set; }
            public bool ProcessCableTrays { get; set; }
            public bool ProcessConduits { get; set; }
            public bool ProcessDuctInsulation { get; set; }
            public bool ProcessPipeInsulation { get; set; }
        }
        
        /// <summary>
        /// Management settings configuration
        /// </summary>
        public class ManagementSettings
        {
            public bool ResetApprovalStatusOnChanges { get; set; }
            public double DimensionChangeThreshold { get; set; }
            public double LocationChangeThreshold { get; set; }
            public bool CutOpeningWithHosts { get; set; }
            public bool CreateConstraintBetweenOpeningsAndHosts { get; set; }
            public bool AdoptProvisionForVoidsFromLinkedModel { get; set; }
            public bool IncludeHostElementsNotVisibleIn3DView { get; set; }
            public bool IncludeReferenceElementsNotVisibleIn3DView { get; set; }
            public bool IncludeHostElementsInDemolishedPhase { get; set; }
        }
        
        /// <summary>
        /// Processing limits configuration
        /// </summary>
        public class ProcessingLimits
        {
            public double IgnoreOpeningsSmallerThan { get; set; }
            public double RoundOpeningsBecomeRectangularIfDiameterGreaterThan { get; set; }
            public double JoinOpeningsIfDistanceLessThan { get; set; }
            public double IgnoreOpeningsWithAngleGreaterThan { get; set; }
        }
        
        #endregion
    }
}
