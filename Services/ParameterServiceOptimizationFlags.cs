using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ SOLID REFACTORING: Feature flags for safe rollout
    /// Follows ENHANCEMENT 10 from comprehensive architecture document
    /// </summary>
    public static class ParameterServiceOptimizationFlags
    {
        /// <summary>
        /// Enable SOLID-refactored helper methods for parameter transfer
        /// If false, falls back to legacy inline code
        /// </summary>
        public static bool UseRefactoredParameterTransfer { get; set; } = true;
        
        /// <summary>
        /// Enable SOLID-refactored helper methods for marking operations
        /// If false, falls back to legacy inline code
        /// </summary>
        public static bool UseRefactoredMarkingOperations { get; set; } = true;
        
        /// <summary>
        /// Enable SOLID-refactored ParameterSnapshotService and ParameterCaptureService
        /// Uses composed services (IParameterCapture, IParameterPolicy, IParameterKeyStore)
        /// If false, uses legacy monolithic code
        /// </summary>
        public static bool UseSolidParameterServices { get; set; } = true; // ✅ ENABLED FOR TESTING
        
        /// <summary>
        /// Enable dependency injection for services
        /// If false, creates services directly (legacy behavior)
        /// </summary>
        public static bool UseDependencyInjection { get; set; } = false; // Start with false for safety
        
        /// <summary>
        /// Enable batch processing optimizations
        /// If false, uses original processing logic
        /// </summary>
        public static bool UseBatchProcessing { get; set; } = false; // ⚠️ DISABLED BY USER REQUEST (Safe Fallback)
        
        /// <summary>
        /// Enable detailed performance diagnostics
        /// </summary>
        public static bool EnablePerformanceDiagnostics { get; set; } = false;
        
        /// <summary>
        /// Load flags from user settings or config file
        /// </summary>
        public static void LoadFromSettings()
        {
            try
            {
                // TODO: Load from Settings.Default or config file
                // For now, use defaults above
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ParameterServiceOptimizationFlags] Failed to load settings: {ex.Message}");
                // Fall back to safe defaults
                UseRefactoredParameterTransfer = false;
                UseRefactoredMarkingOperations = false;
                UseDependencyInjection = false;
            }
        }
        
        /// <summary>
        /// Reset all flags to safe defaults (legacy behavior)
        /// </summary>
        public static void ResetToSafeDefaults()
        {
            UseRefactoredParameterTransfer = false;
            UseRefactoredMarkingOperations = false;
            UseSolidParameterServices = false;
            UseDependencyInjection = false;
            UseBatchProcessing = true; // This is already working
            EnablePerformanceDiagnostics = false;
        }
        
        /// <summary>
        /// Enable all optimizations (for testing)
        /// </summary>
        public static void EnableAllOptimizations()
        {
            UseRefactoredParameterTransfer = true;
            UseRefactoredMarkingOperations = true;
            UseSolidParameterServices = true;
            UseDependencyInjection = true;
            UseBatchProcessing = true;
            EnablePerformanceDiagnostics = true;
        }
    }
}
