using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Helpers
{
    /// <summary>
    /// Revit 2026 API compatibility: ElementId.IntegerValue was removed.
    /// Revit 2023-2025: ElementId.IntegerValue (int property)
    /// Revit 2026: ElementId.Value (long property), IntegerValue removed.
    /// This extension provides a unified accessor for all versions.
    /// </summary>
    public static class ElementIdCompat
    {
        /// <summary>
        /// Unified way to get a numeric value from an ElementId.
        /// Returns long to accommodate Revit 2024+ IDs, while being backward compatible with int.
        /// </summary>
        public static long GetIdValue(this ElementId id)
        {
            if (id == null) return -1;
#if REVIT2024 || REVIT2025 || REVIT2026 || REVIT2024_OR_GREATER || NET8_0_OR_GREATER
            return id.Value;
#else
            return id.IntegerValue;
#endif
        }

        /// <summary>
        /// Unified way to create an ElementId from a long value.
        /// </summary>
        public static ElementId FromValue(long idValue)
        {
#if REVIT2024 || REVIT2025 || REVIT2026 || REVIT2024_OR_GREATER || NET8_0_OR_GREATER
            return new ElementId(idValue);
#else
            return new ElementId((int)idValue);
#endif
        }

        /// <summary>
        /// Unified way to cast a long/int value to a BuiltInParameter.
        /// </summary>
        public static BuiltInParameter GetBip(long bipValue)
        {
            return (BuiltInParameter)bipValue;
        }

        /// <summary>
        /// Legacy helper - redirects to GetIdValue and casts to int.
        /// Use with caution for extremely large IDs in R24+.
        /// </summary>
        public static int GetIntegerValue(this ElementId id)
        {
            return (int)id.GetIdValue();
        }

        /// <summary>
        /// Get the ElementId value as int (safe for DB storage and model properties).
        /// Works on both net48 (R23/R24) and net8.0 (R25/R26).
        /// </summary>
        public static int ToInt(this ElementId id)
        {
            return (int)GetIdValue(id);
        }

        /// <summary>
        /// Get the ElementId value as long (native type on R25/R26).
        /// Works on both net48 (R23/R24) and net8.0 (R25/R26).
        /// </summary>
        public static long ToLong(this ElementId id)
        {
            return GetIdValue(id);
        }

        /// <summary>
        /// Create an ElementId from an int value.
        /// Works on both net48 (R23/R24) and net8.0 (R25/R26).
        /// </summary>
        public static ElementId FromInt(int value)
        {
            return FromValue(value);
        }

        /// <summary>
        /// Create an ElementId from a long value.
        /// Works on both net48 (R23/R24) and net8.0 (R25/R26).
        /// </summary>
        public static ElementId FromLong(long value)
        {
            return FromValue(value);
        }
    }
}
