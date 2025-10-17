using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Configuration
{
    /// <summary>
    /// Service for resolving conflicts between global configuration rules and UI user preferences
    /// Implements the architecture defined in GLOBAL_CONFIGURATION_INTEGRATION_ARCHITECTURE.md
    /// </summary>
    public class ConfigurationResolutionService
    {
        private static ConfigurationResolutionService _instance;
        private static readonly object _lock = new object();
        
        private readonly GlobalConfigurationService _globalConfigService;
        
        private ConfigurationResolutionService()
        {
            _globalConfigService = GlobalConfigurationService.Instance;
            DebugLogger.Info("[ConfigurationResolutionService] Initialized");
        }
        
        public static ConfigurationResolutionService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ConfigurationResolutionService();
                        }
                    }
                }
                return _instance;
            }
        }
        
        /// <summary>
        /// Resolves opening type conflicts between global rules and UI preferences
        /// Priority: Global rules > UI preferences
        /// </summary>
        /// <param name="category">MEP category (e.g., "Pipes", "Ducts")</param>
        /// <param name="elementProps">Element properties including size and shape</param>
        /// <param name="uiPreference">UI preference ("Circular" or "Rectangular")</param>
        /// <returns>Resolved opening type</returns>
        public string ResolveOpeningType(string category, ElementProperties elementProps, string uiPreference)
        {
            try
            {
                // Priority 1: Global size-based rules
                if (category.Equals("Pipes", StringComparison.OrdinalIgnoreCase))
                {
                    var processingLimits = _globalConfigService.GetProcessingLimits();
                    var diameterThreshold = processingLimits.RoundOpeningsBecomeRectangularIfDiameterGreaterThan; // 200mm
                    
                    // Convert diameter to mm for comparison
                    var diameterMm = UnitUtils.ConvertFromInternalUnits(elementProps.Diameter, UnitTypeId.Millimeters);
                    
                    if (diameterMm > diameterThreshold)
                    {
                        DebugLogger.Info($"[ConfigResolution] Pipe {diameterMm:F1}mm > {diameterThreshold}mm threshold → Rectangular (global rule overrides UI: {uiPreference})");
                        return "Rectangular"; // Global rule wins
                    }
                }
                
                // Priority 2: Global category rules
                if (category.Equals("Cable Trays", StringComparison.OrdinalIgnoreCase))
                {
                    DebugLogger.Info($"[ConfigResolution] Cable Trays → Rectangular (global rule overrides UI: {uiPreference})");
                    return "Rectangular"; // Always rectangular per global rule
                }
                
                // Priority 3: UI preference (fallback)
                DebugLogger.Info($"[ConfigResolution] {category} → {uiPreference} (UI preference, no global rule applies)");
                return uiPreference;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ConfigResolution] Error resolving opening type for {category}: {ex.Message}");
                return uiPreference; // Fallback to UI preference
            }
        }
        
        /// <summary>
        /// Resolves clearance conflicts between global limits and UI preferences
        /// Priority: Global limits > UI preferences
        /// </summary>
        /// <param name="category">MEP category</param>
        /// <param name="uiClearance">UI clearance value in mm</param>
        /// <returns>Resolved clearance value in mm</returns>
        public double ResolveClearance(string category, double uiClearance)
        {
            try
            {
                var sizeLimits = _globalConfigService.GetOpeningSizeLimits();
                
                // Priority 1: Global minimum clearance
                var globalMinClearance = sizeLimits.MinWidth; // Use MinWidth as minimum clearance
                if (uiClearance < globalMinClearance)
                {
                    DebugLogger.Info($"[ConfigResolution] {category} clearance {uiClearance}mm < global minimum {globalMinClearance}mm → Using {globalMinClearance}mm");
                    return globalMinClearance; // Global rule wins
                }
                
                // Priority 2: Global maximum clearance
                var globalMaxClearance = sizeLimits.MaxWidth; // Use MaxWidth as maximum clearance
                if (uiClearance > globalMaxClearance)
                {
                    DebugLogger.Info($"[ConfigResolution] {category} clearance {uiClearance}mm > global maximum {globalMaxClearance}mm → Using {globalMaxClearance}mm");
                    return globalMaxClearance; // Global rule wins
                }
                
                // Priority 3: UI preference (within limits)
                DebugLogger.Info($"[ConfigResolution] {category} clearance {uiClearance}mm within limits → Using UI value");
                return uiClearance;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ConfigResolution] Error resolving clearance for {category}: {ex.Message}");
                return uiClearance; // Fallback to UI preference
            }
        }
        
        /// <summary>
        /// Determines if an element should be processed based on global rules
        /// </summary>
        /// <param name="category">MEP category</param>
        /// <param name="elementProps">Element properties</param>
        /// <returns>True if element should be processed</returns>
        public bool ShouldProcessElement(string category, ElementProperties elementProps)
        {
            try
            {
                // Priority 1: Global processing flags
                if (!_globalConfigService.ShouldProcessElementType(category))
                {
                    DebugLogger.Info($"[ConfigResolution] {category} processing disabled by global rule");
                    return false; // Global rule: Don't process this category
                }
                
                // Priority 2: Global size limits
                var processingLimits = _globalConfigService.GetProcessingLimits();
                var diameterMm = UnitUtils.ConvertFromInternalUnits(elementProps.Diameter, UnitTypeId.Millimeters);
                
                if (diameterMm < processingLimits.IgnoreOpeningsSmallerThan)
                {
                    DebugLogger.Info($"[ConfigResolution] {category} diameter {diameterMm:F1}mm < global minimum {processingLimits.IgnoreOpeningsSmallerThan}mm → Skip processing");
                    return false; // Global rule: Too small to process
                }
                
                // Priority 3: Global angle limits
                if (elementProps.Angle > processingLimits.IgnoreOpeningsWithAngleGreaterThan)
                {
                    DebugLogger.Info($"[ConfigResolution] {category} angle {elementProps.Angle:F1}° > global maximum {processingLimits.IgnoreOpeningsWithAngleGreaterThan}° → Skip processing");
                    return false; // Global rule: Angle too steep
                }
                
                DebugLogger.Info($"[ConfigResolution] {category} passes all global rules → Process element");
                return true;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ConfigResolution] Error checking if element should be processed: {ex.Message}");
                return true; // Fallback to process
            }
        }
        
        /// <summary>
        /// Creates a resolved configuration object with all resolved settings
        /// </summary>
        /// <param name="category">MEP category</param>
        /// <param name="elementProps">Element properties</param>
        /// <param name="uiPreferences">UI user preferences</param>
        /// <returns>Resolved configuration</returns>
        public ResolvedConfiguration ResolveConfiguration(string category, ElementProperties elementProps, UIUserPreferences uiPreferences)
        {
            return new ResolvedConfiguration
            {
                Category = category,
                OpeningType = ResolveOpeningType(category, elementProps, uiPreferences.OpeningType),
                Clearance = ResolveClearance(category, uiPreferences.Clearance),
                ShouldProcess = ShouldProcessElement(category, elementProps),
                ElementProperties = elementProps,
                UIPreferences = uiPreferences,
                ResolutionTimestamp = DateTime.Now
            };
        }
    }
    
    /// <summary>
    /// Element properties for configuration resolution
    /// </summary>
    public class ElementProperties
    {
        public double Diameter { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Angle { get; set; }
        public string Shape { get; set; }
        public bool IsInsulated { get; set; }
        public double InsulationThickness { get; set; }
    }
    
    /// <summary>
    /// UI user preferences for configuration resolution
    /// </summary>
    public class UIUserPreferences
    {
        public string OpeningType { get; set; }
        public double Clearance { get; set; }
        public Dictionary<string, double> CategorySpecificClearances { get; set; }
    }
    
    /// <summary>
    /// Resolved configuration after applying global rules and UI preferences
    /// </summary>
    public class ResolvedConfiguration
    {
        public string Category { get; set; }
        public string OpeningType { get; set; }
        public double Clearance { get; set; }
        public bool ShouldProcess { get; set; }
        public ElementProperties ElementProperties { get; set; }
        public UIUserPreferences UIPreferences { get; set; }
        public DateTime ResolutionTimestamp { get; set; }
        
        public override string ToString()
        {
            return $"[ResolvedConfig] {Category}: {OpeningType} opening, {Clearance}mm clearance, Process={ShouldProcess}";
        }
    }
}
