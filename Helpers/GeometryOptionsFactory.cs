using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// Factory for consistent geometry extraction Options.
    /// 2024 build enables richer geometry + references to improve intersection detection reliability.
    /// 2023 build keeps conservative settings to avoid behavioral regression.
    /// </summary>
    public static class GeometryOptionsFactory
    {
        /// <summary>
        /// Options for intersection / clash oriented extraction.
        /// Replace direct 'new Options()' calls with this method.
        /// </summary>
        public static Options CreateIntersectionOptions()
        {
            var opt = new Options();
#if REVIT2024_OR_GREATER
            // Minimal adjustments for 2024 per migration doc feedback:
            // Fine detail only (retain solids fidelity). Leave other flags off.
            opt.DetailLevel = ViewDetailLevel.Fine;
            opt.IncludeNonVisibleObjects = false; // User request: do NOT load hidden/non-visible geometry.
            opt.ComputeReferences = false;       // Keep disabled to avoid overhead.
#else
            // 2023 unchanged baseline
            opt.DetailLevel = ViewDetailLevel.Medium;
            opt.IncludeNonVisibleObjects = false;
            opt.ComputeReferences = false;
#endif
            return opt;
        }
    }
}
