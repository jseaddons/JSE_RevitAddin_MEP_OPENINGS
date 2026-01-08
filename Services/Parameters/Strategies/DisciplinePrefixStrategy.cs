using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Strategies;
using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Strategies
{
    public class DisciplinePrefixStrategy : IPrefixStrategy
    {
        public string ResolvePrefix(string category, MarkPrefixSettings settings, ClashZone zone)
        {
            if (settings == null) return "OPN"; // Fallback

            // ✅ PRIORITIZE JSON: Use values from JSON blob first (matches User Expectation better)
            // The direct DB column "MepSystemType" often contains internal Revit names (e.g., "Mechanical Supply Air 70")
            // The JSON blob contains the parameter value (e.g., "Supply Air") which matches the UI settings.
            string systemType = GetParameterValue(zone, "System Type");
            string serviceType = GetParameterValue(zone, "Service Type");
            
            // Fallback to DB columns if JSON lookup failed
            if (string.IsNullOrEmpty(systemType)) systemType = zone.MepSystemType?.Trim();
            if (string.IsNullOrEmpty(serviceType)) serviceType = zone.MepServiceType?.Trim();

            NumberingDebugLogger.LogInfo($"[DisciplinePrefixStrategy] Category: {category}, SystemType: '{systemType}', ServiceType: '{serviceType}'");

            // Delegate to Settings logic (which handles the overrides)
            return settings.GetPrefixForElement(category, systemType, serviceType);
        }

        private string GetParameterValue(ClashZone zone, string paramName)
        {
            if (zone?.MepParameterValues == null) 
            {
                NumberingDebugLogger.LogInfo($"[DisciplinePrefixStrategy] Zone {zone?.ClashZoneId} has NULL MepParameterValues");
                return null;
            }
            
            // Try Case-Insensitive Key Lookup
            foreach (var kv in zone.MepParameterValues)
            {
                if (string.Equals(kv.Key, paramName, StringComparison.OrdinalIgnoreCase))
                {
                    return kv.Value;
                }
            }

            // Diagnostic: Log all keys if not found
            var keys = string.Join(", ", zone.MepParameterValues.Select(k => k.Key));
            NumberingDebugLogger.LogInfo($"[DisciplinePrefixStrategy] Parameter '{paramName}' NOT found. Available keys: {keys}");
            
            return null;
        }
    }
}
