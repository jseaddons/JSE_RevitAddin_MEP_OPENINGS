using System.Collections.Generic;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.ClearanceProviders
{
    /// <summary>
    /// Singleton manager for UI clearance settings
    /// Provides global access to clearance overrides from UI
    /// </summary>
    public class ClearanceManager
    {
        private static ClearanceManager _instance;
        private static readonly object _lock = new object();
        
        private Dictionary<string, double> _uiClearances;
        private readonly ClearanceProviderFactory _providerFactory;
        
        public static ClearanceManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                            _instance = new ClearanceManager();
                    }
                }
                return _instance;
            }
        }
        
        private ClearanceManager()
        {
            _uiClearances = null;
            _providerFactory = new ClearanceProviderFactory();
        }
        
        /// <summary>
        /// Set UI clearance settings to override default clearance calculations
        /// </summary>
        /// <param name="clearances">Dictionary of clearance settings from UI</param>
        public void SetUIClearances(Dictionary<string, double> clearances)
        {
            DebugLogger.Info($"[CLEARANCE_DEBUG] ClearanceManager.SetUIClearances called with {clearances.Count} values:");
            foreach (var kvp in clearances)
            {
                DebugLogger.Info($"[CLEARANCE_DEBUG]   {kvp.Key} = {kvp.Value}mm");
            }
            
            _uiClearances = clearances;
            
            DebugLogger.Info($"[CLEARANCE_DEBUG] ClearanceManager._uiClearances set successfully");
        }
        
        /// <summary>
        /// Get clearance for a MEP element using appropriate provider
        /// </summary>
        /// <param name="mepElement">The MEP element</param>
        /// <returns>Clearance value in internal units</returns>
        public double GetClearance(Element mepElement)
        {
            var provider = _providerFactory.GetProvider(mepElement);
            return provider.GetClearance(mepElement, _uiClearances);
        }
        
        /// <summary>
        /// Clear UI clearance settings (revert to defaults)
        /// </summary>
        public void ClearUIClearances()
        {
            _uiClearances = null;
        }
        
        /// <summary>
        /// Check if UI clearances are currently set
        /// </summary>
        /// <returns>True if UI clearances are active</returns>
        public bool HasUIClearances()
        {
            return _uiClearances != null && _uiClearances.Count > 0;
        }

        /// <summary>
        /// Get the current UI clearance settings
        /// </summary>
        /// <returns>Dictionary of UI clearance settings</returns>
        public Dictionary<string, double> GetUIClearances()
        {
            var result = _uiClearances ?? new Dictionary<string, double>();
            
            DebugLogger.Info($"[CLEARANCE_DEBUG] ClearanceManager.GetUIClearances returning {result.Count} values:");
            foreach (var kvp in result)
            {
                DebugLogger.Info($"[CLEARANCE_DEBUG]   {kvp.Key} = {kvp.Value}mm");
            }
            
            return result;
        }
    }
}
