using System;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    public class DisciplinePrefixStrategy : IPrefixStrategy
    {
        public string ResolvePrefix(string category, MarkPrefixSettings settings, ClashZone zone)
        {
            if (settings == null) return "OPN"; // Fallback

            // ✅ OPTIMIZED: Use pre-resolved properties from model (loaded from direct DB columns)
            string systemType = zone.MepSystemType?.Trim();
            string serviceType = zone.MepServiceType?.Trim();

            // Fallback to searching parameter list if DB columns are empty (legacy/backward compatibility)
            if (string.IsNullOrEmpty(systemType)) systemType = GetParameterValue(zone, "System Type");
            if (string.IsNullOrEmpty(serviceType)) serviceType = GetParameterValue(zone, "Service Type");

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
