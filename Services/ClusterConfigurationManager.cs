using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ⚠️ CRITICAL: Singleton configuration manager for cluster sleeve placement ⚠️
    /// 
    /// Provides thread-safe access to cluster configuration settings that affect
    /// how closely-spaced sleeves are merged into cluster openings.
    /// 
    /// This manager bridges the gap between the orchestrator (which has access to
    /// user settings) and the cluster command (which is invoked via IExternalCommand
    /// interface and cannot receive custom parameters).
    /// </summary>
    public class ClusterConfigurationManager
    {
        private static ClusterConfigurationManager _instance;
        private static readonly object _lock = new object();
        
        /// <summary>
        /// Distance threshold (in millimeters) for merging sleeves into clusters.
        /// Sleeves within this edge-to-edge distance will be grouped together.
        /// 
        /// Default: 200mm (balanced clustering)
        /// Range: 50mm (tight) to 500mm (aggressive)
        /// </summary>
        public double JoinOpeningsDistance { get; private set; } = 200.0;
        
        /// <summary>
        /// When this configuration was last updated
        /// </summary>
        public DateTime LastUpdated { get; private set; } = DateTime.Now;
        
        /// <summary>
        /// Source of the current configuration (for debugging)
        /// </summary>
        public string ConfigurationSource { get; private set; } = "Default";
        
        /// <summary>
        /// Thread-safe singleton instance
        /// </summary>
        public static ClusterConfigurationManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ClusterConfigurationManager();
                            DebugLogger.Info("[ClusterConfig] Initialized with default settings");
                        }
                    }
                }
                return _instance;
            }
        }
        
        /// <summary>
        /// Private constructor to enforce singleton pattern
        /// </summary>
        private ClusterConfigurationManager()
        {
            // Default initialization
        }
        
        /// <summary>
        /// Set the join openings distance from user configuration
        /// </summary>
        /// <param name="distanceMm">Distance in millimeters (50-500mm recommended)</param>
        /// <param name="source">Source of the configuration (for logging)</param>
        public void SetJoinOpeningsDistance(double distanceMm, string source = "User Configuration")
        {
            // Validate input
            if (distanceMm < 10.0 || distanceMm > 1000.0)
            {
                DebugLogger.Warning($"[ClusterConfig] Invalid JoinOpeningsDistance: {distanceMm}mm (must be 10-1000mm). Using default 200mm.");
                distanceMm = 200.0;
                source = "Default (invalid input)";
            }
            
            lock (_lock)
            {
                JoinOpeningsDistance = distanceMm;
                LastUpdated = DateTime.Now;
                ConfigurationSource = source;
                
                DebugLogger.Info($"[ClusterConfig] ✓ JoinOpeningsDistance set to: {distanceMm}mm");
                DebugLogger.Info($"[ClusterConfig]   Source: {source}");
                DebugLogger.Info($"[ClusterConfig]   Updated: {LastUpdated:yyyy-MM-dd HH:mm:ss}");
                DebugLogger.Info($"[ClusterConfig]   Internal units: {GetJoinOpeningsDistanceInFeet():F6} feet");
            }
        }
        
        /// <summary>
        /// Get the join openings distance converted to Revit internal units (feet)
        /// </summary>
        public double GetJoinOpeningsDistanceInFeet()
        {
            return UnitUtils.ConvertToInternalUnits(JoinOpeningsDistance, UnitTypeId.Millimeters);
        }
        
        /// <summary>
        /// Reset to default configuration
        /// </summary>
        public void Reset()
        {
            lock (_lock)
            {
                JoinOpeningsDistance = 200.0;
                LastUpdated = DateTime.Now;
                ConfigurationSource = "Reset to Default";
                
                DebugLogger.Info("[ClusterConfig] Configuration reset to default (200mm)");
            }
        }
        
        /// <summary>
        /// Get current configuration summary for logging
        /// </summary>
        public string GetConfigurationSummary()
        {
            return $"JoinOpeningsDistance: {JoinOpeningsDistance}mm ({GetJoinOpeningsDistanceInFeet():F6} ft), " +
                   $"Source: {ConfigurationSource}, Updated: {LastUpdated:yyyy-MM-dd HH:mm:ss}";
        }
        
        /// <summary>
        /// Load configuration from SettingsModel
        /// </summary>
        public void LoadFromSettings(Models.SettingsModel settings, string source = "SettingsModel")
        {
            if (settings == null)
            {
                DebugLogger.Warning("[ClusterConfig] Null settings provided - using defaults");
                return;
            }
            
            SetJoinOpeningsDistance(settings.JoinOpeningsDistance, source);
        }
    }
}

