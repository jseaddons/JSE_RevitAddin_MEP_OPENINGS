using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Strategies
{
    public interface IPrefixStrategy
    {
        /// <summary>
        /// Determines the prefix for a specific ClashZone based on category, settings, and element parameters (System Type).
        /// </summary>
        string ResolvePrefix(string category, object settings, ClashZone zone);
    }
}
