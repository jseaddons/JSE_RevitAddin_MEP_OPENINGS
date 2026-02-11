using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// Revit 2026 API compatibility: ElementId.IntegerValue was removed.
    /// Revit 2023-2025: ElementId.IntegerValue (int property)
    /// Revit 2026: ElementId.Value (long property), IntegerValue removed.
    /// This extension provides a unified accessor for all versions.
    /// </summary>
    internal static class ElementIdCompat
    {
        internal static int GetIntegerValue(this ElementId id)
        {
#if REVIT2026
            return (int)id.Value;
#else
            return id.IntegerValue;
#endif
        }
    }
}
