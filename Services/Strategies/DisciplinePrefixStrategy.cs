using System;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    public class DisciplinePrefixStrategy : IPrefixStrategy
    {
        public string ResolvePrefix(string category, object settingsObj, ClashZone zone)
        {
             // DEPRECATED: Logic simplified as per user request.
             // Prefix strategy based on settings is disabled. Returning generic default.
             return "MEP";
        }
    }
}
