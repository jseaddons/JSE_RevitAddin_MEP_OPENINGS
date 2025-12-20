using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.ParameterCapture;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// ✅ SOLID REFACTORING: Facade pattern for backward compatibility
    /// Composes IParameterCapture, IParameterPolicy, and IParameterKeyStore
    /// Preserves all 28 features from comprehensive architecture
    /// </summary>
    public partial class ParameterSnapshotService
    {
        // ✅ SOLID: Composed services (Dependency Inversion Principle)
        private static IParameterCapture _parameterCapture;
        private static IParameterPolicy _snapshotPolicy;
        private static IParameterKeyStore _keyStore;
        
        // Static constructor to initialize composed services
        static ParameterSnapshotService()
        {
            InitializeServices();
        }
        
        /// <summary>
        /// ✅ SOLID: Initialize composed services
        /// Can be swapped for testing or different implementations
        /// </summary>
        private static void InitializeServices()
        {
            _keyStore = new FileParameterKeyStore();
            _snapshotPolicy = new SnapshotParameterPolicy(_keyStore);
            _parameterCapture = new RevitParameterCapture();
        }
        
        /// <summary>
        /// ✅ SOLID: Allow dependency injection for testing
        /// </summary>
        public static void SetServices(
            IParameterCapture parameterCapture,
            IParameterPolicy parameterPolicy,
            IParameterKeyStore keyStore)
        {
            _parameterCapture = parameterCapture ?? throw new ArgumentNullException(nameof(parameterCapture));
            _snapshotPolicy = parameterPolicy ?? throw new ArgumentNullException(nameof(parameterPolicy));
            _keyStore = keyStore ?? throw new ArgumentNullException(nameof(keyStore));
        }
        
        /// <summary>
        /// ✅ SOLID: Reset to default services
        /// </summary>
        public static void ResetToDefaultServices()
        {
            InitializeServices();
        }
        
        // ========================================================================
        // ✅ BACKWARD-COMPATIBLE PUBLIC API
        // Delegates to composed services based on feature flag
        // ========================================================================
        
        /// <summary>
        /// Build a whitelist for the current run
        /// ✅ FEATURE FLAG: Uses SOLID services or legacy code
        /// </summary>
        public static HashSet<string> BuildWhitelist()
        {
            if (ParameterServiceOptimizationFlags.UseSolidParameterServices)
            {
                return _snapshotPolicy.GetWhitelist();
            }
            else
            {
                return BuildWhitelistLegacy();
            }
        }
        
        /// <summary>
        /// Capture whitelisted parameter values for an element
        /// ✅ FEATURE FLAG: Uses SOLID services or legacy code
        /// </summary>
        public static List<SerializableKeyValue> CaptureParams(
            Element element,
            HashSet<string> whitelist,
            Document doc,
            string docKey)
        {
            if (ParameterServiceOptimizationFlags.UseSolidParameterServices)
            {
                // ✅ SOLID: Delegate to composed service
                return _parameterCapture.CaptureParameters(element, _snapshotPolicy);
            }
            else
            {
                // ✅ LEGACY: Use original code
                return CaptureParamsLegacy(element, whitelist, doc, docKey);
            }
        }
        
        /// <summary>
        /// Convert a parameter value to string
        /// ✅ FEATURE FLAG: Uses SOLID services or legacy code
        /// </summary>
        public static string ConvertParameterToString(Element element, Parameter param)
        {
            if (ParameterServiceOptimizationFlags.UseSolidParameterServices)
            {
                return _parameterCapture.ConvertParameterToString(element, param);
            }
            else
            {
                return ConvertParameterToStringLegacy(element, param);
            }
        }
        
        /// <summary>
        /// Lookup parameter with fallbacks
        /// ✅ FEATURE FLAG: Uses SOLID services or legacy code
        /// </summary>
        public static Parameter LookupParam(Element element, string paramName)
        {
            if (ParameterServiceOptimizationFlags.UseSolidParameterServices)
            {
                return _parameterCapture.LookupParameter(element, paramName);
            }
            else
            {
                return LookupParamLegacy(element, paramName);
            }
        }
        
        /// <summary>
        /// Add a learned key
        /// ✅ FEATURE FLAG: Uses SOLID services or legacy code
        /// </summary>
        public static void AddLearnedKey(string key)
        {
            if (ParameterServiceOptimizationFlags.UseSolidParameterServices)
            {
                _keyStore.AddLearnedKey(key);
            }
            else
            {
                AddLearnedKeyLegacy(key);
            }
        }
        
        /// <summary>
        /// Get document key
        /// ✅ BACKWARD COMPATIBLE: No change needed
        /// </summary>
        public static string GetDocKey(Document doc)
        {
            return doc.IsLinked ? $"LINK_{doc.Title}" : "HOST";
        }
        
        // ========================================================================
        // ✅ LEGACY METHODS (Preserved for fallback)
        // Original code moved to separate partial class file
        // ========================================================================
        
        private static partial HashSet<string> BuildWhitelistLegacy();
        private static partial List<SerializableKeyValue> CaptureParamsLegacy(
            Element element, HashSet<string> whitelist, Document doc, string docKey);
        private static partial string ConvertParameterToStringLegacy(Element element, Parameter param);
        private static partial Parameter LookupParamLegacy(Element element, string paramName);
        private static partial void AddLearnedKeyLegacy(string key);
    }
}
