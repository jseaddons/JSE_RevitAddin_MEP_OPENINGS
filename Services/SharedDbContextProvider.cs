using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Data;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Provides a shared SleeveDbContext instance per session when enabled by flag.
    /// Falls back to new instances when reuse is disabled.
    /// </summary>
    public static class SharedDbContextProvider
    {
        private static SleeveDbContext? _shared;
        private static Document? _doc;

        /// <summary>
        /// Get a SleeveDbContext, reusing a single instance per session when the flag is enabled.
        /// </summary>
        public static SleeveDbContext GetOrCreate(Document document, Action<string>? logger = null)
        {
            if (!OptimizationFlags.ReuseDbContextDuringRefresh)
            {
                return new SleeveDbContext(document, logger);
            }

            if (_shared == null || _doc == null)
            {
                _doc = document;
                _shared = new SleeveDbContext(document, logger);
            }
            return _shared;
        }

        /// <summary>
        /// Reset the shared instance (optional; call between sessions if needed).
        /// </summary>
        public static void Reset()
        {
            try { _shared?.Dispose(); } catch { }
            _shared = null;
            _doc = null;
        }
    }
}
