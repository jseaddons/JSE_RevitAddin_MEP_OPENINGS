using System;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Switch
{
    /// <summary>
    /// ✅ MASTER SWITCH: Central control for all switches
    /// 
    /// USAGE:
    ///   MasterSwitch.EnableDiagnostics();     // Enable diagnostic mode
    ///   MasterSwitch.DisableDiagnostics();     // Disable diagnostic mode
    ///   MasterSwitch.EnableAllOptimizations(); // Enable all optimizations
    ///   MasterSwitch.DisableAllOptimizations(); // Disable all optimizations
    ///   MasterSwitch.SetDeploymentMode();      // Set to deployment mode (minimal logging)
    ///   MasterSwitch.SetDevelopmentMode();     // Set to development mode (full logging)
    ///   
    ///   // ✅ SINGLE SETTING: One property to control both diagnostic logging and deployment mode
    ///   MasterSwitch.DiagnosticLogging = true;  // Enable diagnostic logging + Disable deployment mode
    ///   MasterSwitch.DiagnosticLogging = false; // Disable diagnostic logging + Enable deployment mode
    /// </summary>
    public static class MasterSwitch
    {
        /// <summary>
        /// ✅ SINGLE SETTING: One property to control both diagnostic logging and deployment mode
        /// 
        /// When true:  Enables diagnostic logging (UseDiagnosticMode = true) AND disables deployment mode (DeploymentMode = false)
        /// When false: Disables diagnostic logging (UseDiagnosticMode = false) AND enables deployment mode (DeploymentMode = true)
        /// 
        /// ⚠️ PERFORMANCE WARNING: Diagnostic mode ON can slow down refresh by 3x (30 zones/sec → 9 zones/sec)
        /// This is due to 273+ logging calls in refresh code. Use sparingly for debugging only.
        /// 
        /// USAGE:
        ///   MasterSwitch.DiagnosticLogging = true;  // Turn ON diagnostic logging, turn OFF deployment mode (SLOWER - 9 zones/sec)
        ///   MasterSwitch.DiagnosticLogging = false; // Turn OFF diagnostic logging, turn ON deployment mode (FASTER - 30 zones/sec)
        /// </summary>
        public static bool DiagnosticLogging
        {
            get
            {
                // Diagnostic logging is enabled when: UseDiagnosticMode is true AND DeploymentMode is false
                return OptimizationFlags.UseDiagnosticMode && !DeploymentConfiguration.DeploymentMode;
            }
            set
            {
                if (value)
                {
                    // Enable diagnostic logging: Set UseDiagnosticMode = true AND DeploymentMode = false
                    // ⚠️ WARNING: This will slow down refresh significantly (3x slower) due to excessive logging
                    OptimizationFlags.UseDiagnosticMode = true;
                    DeploymentConfiguration.DeploymentMode = false;
                    OptimizationFlags.LogPerformanceMetrics = true;
                    OptimizationFlags.DisableVerboseLogging = false; // ✅ ENABLE VERBOSE LOGGING
                    
                    // Use System.Diagnostics.Debug to avoid logging overhead when enabling logging
                    System.Diagnostics.Debug.WriteLine("[MasterSwitch] ✅ DiagnosticLogging = true: Diagnostic logging ENABLED (WARNING: 3x slower performance)");
                }
                else
                {
                    // Disable diagnostic logging: Set UseDiagnosticMode = false AND DeploymentMode = true
                    // ✅ This will restore normal performance (30 zones/sec)
                    OptimizationFlags.UseDiagnosticMode = false;
                    DeploymentConfiguration.DeploymentMode = true;
                    OptimizationFlags.LogPerformanceMetrics = false;
                    OptimizationFlags.DisableVerboseLogging = true;  // ✅ DISABLE VERBOSE LOGGING
                    
                    System.Diagnostics.Debug.WriteLine("[MasterSwitch] ⚠️ DiagnosticLogging = false: Diagnostic logging DISABLED, Deployment mode ENABLED (FASTER performance)");
                }
            }
        }
        /// <summary>
        /// Enable diagnostic mode (full logging)
        /// </summary>
        public static void EnableDiagnostics()
        {
            DiagnosticSwitch.Enable();
        }

        /// <summary>
        /// Disable diagnostic mode (deployment mode - minimal logging)
        /// </summary>
        public static void DisableDiagnostics()
        {
            DiagnosticSwitch.Disable();
        }

        /// <summary>
        /// Toggle diagnostic mode
        /// </summary>
        public static void ToggleDiagnostics()
        {
            DiagnosticSwitch.Toggle();
        }

        /// <summary>
        /// Enable all optimizations (high-impact, low-risk)
        /// </summary>
        public static void EnableAllOptimizations()
        {
            OptimizationSwitch.EnableAllSafeOptimizations();
        }

        /// <summary>
        /// Disable all optimizations (safe mode)
        /// </summary>
        public static void DisableAllOptimizations()
        {
            OptimizationSwitch.DisableAllOptimizations();
        }

        /// <summary>
        /// Set to deployment mode (minimal logging, all optimizations enabled)
        /// </summary>
        public static void SetDeploymentMode()
        {
            DiagnosticSwitch.Disable();
            OptimizationSwitch.EnableAllSafeOptimizations();
            
            System.Diagnostics.Debug.WriteLine("[MasterSwitch] 🚀 Deployment mode activated - Minimal logging, all optimizations enabled");
        }

        /// <summary>
        /// Set to development mode (full logging, all optimizations enabled)
        /// </summary>
        public static void SetDevelopmentMode()
        {
            DiagnosticSwitch.Enable();
            OptimizationSwitch.EnableAllSafeOptimizations();
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[MasterSwitch] 🔧 Development mode activated - Full logging, all optimizations enabled");
            }
        }

        /// <summary>
        /// Set to debug mode (full logging, optimizations disabled for easier debugging)
        /// </summary>
        public static void SetDebugMode()
        {
            DiagnosticSwitch.Enable();
            OptimizationSwitch.DisableAllOptimizations();
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[MasterSwitch] 🐛 Debug mode activated - Full logging, optimizations disabled");
            }
        }

        /// <summary>
        /// Get complete status of all switches
        /// </summary>
        public static string GetCompleteStatus()
        {
            var status = new System.Text.StringBuilder();
            status.AppendLine("════════════════════════════════════════");
            status.AppendLine("🔌 MASTER SWITCH STATUS");
            status.AppendLine("════════════════════════════════════════");
            status.AppendLine();
            status.AppendLine(DiagnosticSwitch.GetStatus());
            status.AppendLine();
            status.AppendLine(OptimizationSwitch.GetStatus());
            status.AppendLine("════════════════════════════════════════");
            return status.ToString();
        }
    }
}

