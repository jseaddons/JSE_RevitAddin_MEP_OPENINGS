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

            // Extract System Type or Service Type from ClashZone parameters
            string systemType = GetParameterValue(zone, "System Type");
            string serviceType = GetParameterValue(zone, "Service Type");

            // Delegate to Settings logic (which handles the overrides)
            return settings.GetPrefixForElement(category, systemType, serviceType);
        }

        private string GetParameterValue(ClashZone zone, string paramName)
        {
            if (zone?.MepParameterValues == null) return null;
            
            // Try Case-Insensitive Key Lookup
            foreach (var kv in zone.MepParameterValues)
            {
                if (string.Equals(kv.Key, paramName, StringComparison.OrdinalIgnoreCase))
                {
                    return kv.Value;
                }
            }
            return null;
        }
    }
}
