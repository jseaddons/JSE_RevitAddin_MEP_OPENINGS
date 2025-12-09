using System;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Switch
{
    /// <summary>
    /// ✅ OPTIMIZATION SWITCH: Easy control for optimization flags
    /// 
    /// USAGE:
    ///   OptimizationSwitch.EnableBulkClusterSave();  // Enable bulk cluster save optimization
    ///   OptimizationSwitch.DisableBulkClusterSave(); // Disable bulk cluster save
    ///   OptimizationSwitch.EnableBatchParameterCapture(); // Enable batch parameter capture
    ///   OptimizationSwitch.DisableAllOptimizations(); // Disable all optimizations (safe mode)
    /// </summary>
    public static class OptimizationSwitch
    {
        /// <summary>
        /// Enable bulk cluster save optimization (90%+ faster database saves)
        /// </summary>
        public static void EnableBulkClusterSave()
        {
            OptimizationFlags.UseBulkClusterSave = true;
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[OptimizationSwitch] ✅ Bulk cluster save optimization ENABLED");
            }
        }

        /// <summary>
        /// Disable bulk cluster save optimization (use original method)
        /// </summary>
        public static void DisableBulkClusterSave()
        {
            OptimizationFlags.UseBulkClusterSave = false;
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[OptimizationSwitch] ⚠️ Bulk cluster save optimization DISABLED");
            }
        }

        /// <summary>
        /// Enable batch parameter capture optimization (30-50% faster)
        /// </summary>
        public static void EnableBatchParameterCapture()
        {
            OptimizationFlags.UseBatchParameterCapture = true;
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[OptimizationSwitch] ✅ Batch parameter capture optimization ENABLED");
            }
        }

        /// <summary>
        /// Disable batch parameter capture optimization
        /// </summary>
        public static void DisableBatchParameterCapture()
        {
            OptimizationFlags.UseBatchParameterCapture = false;
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[OptimizationSwitch] ⚠️ Batch parameter capture optimization DISABLED");
            }
        }

        /// <summary>
        /// Enable precompute host solids optimization (significant performance gain)
        /// </summary>
        public static void EnablePrecomputeHostSolids()
        {
            OptimizationFlags.PrecomputeHostSolids = true;
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[OptimizationSwitch] ✅ Precompute host solids optimization ENABLED");
            }
        }

        /// <summary>
        /// Disable precompute host solids optimization
        /// </summary>
        public static void DisablePrecomputeHostSolids()
        {
            OptimizationFlags.PrecomputeHostSolids = false;
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[OptimizationSwitch] ⚠️ Precompute host solids optimization DISABLED");
            }
        }

        /// <summary>
        /// Disable all optimizations (safe mode - use original methods)
        /// </summary>
        public static void DisableAllOptimizations()
        {
            OptimizationFlags.UseBulkClusterSave = false;
            OptimizationFlags.UseBatchParameterCapture = false;
            OptimizationFlags.PrecomputeHostSolids = false;
            OptimizationFlags.UseBatchParameterLookups = false;
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[OptimizationSwitch] ⚠️ All optimizations DISABLED - Safe mode active");
            }
        }

        /// <summary>
        /// Enable all safe optimizations (high-impact, low-risk)
        /// </summary>
        public static void EnableAllSafeOptimizations()
        {
            OptimizationFlags.UseBulkClusterSave = true;
            OptimizationFlags.UseBatchParameterCapture = true;
            OptimizationFlags.PrecomputeHostSolids = true;
            OptimizationFlags.UseBatchParameterLookups = true;
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info("[OptimizationSwitch] ✅ All safe optimizations ENABLED");
            }
        }

        /// <summary>
        /// Get current optimization status as a formatted string
        /// </summary>
        public static string GetStatus()
        {
            var status = new System.Text.StringBuilder();
            status.AppendLine("📊 Optimization Status:");
            status.AppendLine($"  Bulk Cluster Save: {(OptimizationFlags.UseBulkClusterSave ? "✅ ENABLED" : "❌ DISABLED")}");
            status.AppendLine($"  Batch Parameter Capture: {(OptimizationFlags.UseBatchParameterCapture ? "✅ ENABLED" : "❌ DISABLED")}");
            status.AppendLine($"  Precompute Host Solids: {(OptimizationFlags.PrecomputeHostSolids ? "✅ ENABLED" : "❌ DISABLED")}");
            status.AppendLine($"  Batch Parameter Lookups: {(OptimizationFlags.UseBatchParameterLookups ? "✅ ENABLED" : "❌ DISABLED")}");
            return status.ToString();
        }
    }
}

