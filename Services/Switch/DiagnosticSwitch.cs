using System;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Switch
{
    /// <summary>
    /// ✅ DIAGNOSTIC SWITCH: Easy control for diagnostic mode and logging
    /// 
    /// USAGE:
    ///   DiagnosticSwitch.Enable();  // Enable full diagnostic logging
    ///   DiagnosticSwitch.Disable(); // Disable diagnostic logging (deployment mode)
    ///   DiagnosticSwitch.Toggle();  // Toggle current state
    ///   bool isEnabled = DiagnosticSwitch.IsEnabled; // Check current state
    /// 
    /// WHAT IT CONTROLS:
    ///   - DeploymentConfiguration.DeploymentMode (false = full logging, true = minimal logging)
    ///   - OptimizationFlags.UseDiagnosticMode (true = diagnostic mode ON)
    ///   - OptimizationFlags.LogPerformanceMetrics (true = performance logging ON)
    /// </summary>
    public static class DiagnosticSwitch
    {
        /// <summary>
        /// Check if diagnostic mode is currently enabled
        /// </summary>
        public static bool IsEnabled
        {
            get
            {
                return !DeploymentConfiguration.DeploymentMode && 
                       OptimizationFlags.UseDiagnosticMode;
            }
        }

        /// <summary>
        /// Enable full diagnostic mode with all logging
        /// </summary>
        public static void Enable()
        {
            DeploymentConfiguration.DeploymentMode = false;  // Enable full logging
            OptimizationFlags.UseDiagnosticMode = true;      // Enable diagnostic mode
            OptimizationFlags.LogPerformanceMetrics = true;  // Enable performance metrics
            OptimizationFlags.DisableVerboseLogging = false; // ✅ ENABLE VERBOSE LOGGING
            DebugLogger.IsEnabled = true;
            LoggingConfiguration.EnableRefreshButton = true;
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[DiagnosticSwitch] ✅ Diagnostic mode ENABLED - Full logging active");
            }
        }

        /// <summary>
        /// Disable diagnostic mode (deployment mode - minimal logging)
        /// </summary>
        public static void Disable()
        {
            DeploymentConfiguration.DeploymentMode = true;   // Disable full logging
            OptimizationFlags.UseDiagnosticMode = false;     // Disable diagnostic mode
            OptimizationFlags.LogPerformanceMetrics = false; // Disable performance metrics
            OptimizationFlags.DisableVerboseLogging = true;  // ✅ DISABLE VERBOSE LOGGING
            DebugLogger.IsEnabled = false;
            LoggingConfiguration.EnableRefreshButton = false;
            
            // Note: Can't use DebugLogger here since DeploymentMode is now true
            System.Diagnostics.Debug.WriteLine("[DiagnosticSwitch] ⚠️ Diagnostic mode DISABLED - Deployment mode active (minimal logging)");
        }

        /// <summary>
        /// Toggle diagnostic mode (enable if disabled, disable if enabled)
        /// </summary>
        public static void Toggle()
        {
            if (IsEnabled)
            {
                Disable();
            }
            else
            {
                Enable();
            }
        }

        /// <summary>
        /// Get current diagnostic mode status as a formatted string
        /// </summary>
        public static string GetStatus()
        {
            return IsEnabled
                ? "✅ Diagnostic Mode: ENABLED (Full logging active)"
                : "⚠️ Diagnostic Mode: DISABLED (Deployment mode - minimal logging)";
        }
    }
}

