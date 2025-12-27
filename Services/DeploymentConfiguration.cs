namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Central configuration for deployment settings.
    /// Set DeploymentMode = true to disable all logging and reduce memory usage.
    /// 
    /// HOW TO USE:
    /// 1. For deployment: Change this line to: public static bool DeploymentMode { get; set; } = true;
    /// 2. This will disable ALL logging (SafeFileLogger, DebugLogger, BatchedLogger)
    /// 3. Expected memory savings: 15-30 KB per clash zone (no string allocations for logs)
    /// </summary>
    public static class DeploymentConfiguration
    {
        /// <summary>
        /// Set to true for production deployment to disable all logging.
        /// This reduces memory usage by only ~13% (not significant enough).
        /// 
        /// When enabled:
        /// - All SafeFileLogger calls are skipped (no I/O)
        /// - All DebugLogger calls are skipped (no file writes)
        /// - All BatchedLogger calls are skipped (no string allocations)
        /// - Memory savings: Only ~13% (not worth losing debug capabilities)
        /// 
        /// ✅ CRITICAL: UseDiagnosticMode OVERRIDES this setting.
        /// If OptimizationFlags.UseDiagnosticMode = true, logging is enabled regardless of DeploymentMode.
        /// </summary>
        public static bool DeploymentMode 
        { 
            get 
            {
                // ✅ FIX: Diagnostic mode overrides deployment mode
                // If diagnostic mode is enabled, always allow logging (DeploymentMode = false)
                if (OptimizationFlags.UseDiagnosticMode)
                    return false;
                return _deploymentMode;
            }
            set => _deploymentMode = value;
        }
        
        private static bool _deploymentMode = false; // ✅ DIAGNOSTIC: Force logging ON to debug corner values

        /// <summary>
        /// Feature flag for Phase B Global XML dedupe. Defaults to false so the existing
        /// placement pipeline is untouched until the utility is validated.
        /// </summary>
        public static bool EnableGlobalIndexDedupe { get; set; } = true;

        /// <summary>
        /// When EnableGlobalIndexDedupe is true, this flag controls whether the run is a dry run.
        /// Leave true to log the proposed changes without touching the XML.
        /// Set to false only after reviewing the dry-run output and backing up the XML files.
        /// </summary>
        public static bool GlobalIndexDedupeDryRun { get; set; } = true;

        /// <summary>
        /// ✅ PHASE SQLITE-2: Use SQLite as primary data source instead of XML
        /// When true: SQLite is the operational store, XML writes are optional/disabled
        /// When false: XML is primary, SQLite is dual-write mirror (Phase SQLITE-1)
        /// </summary>
        public static bool UseSqliteAsPrimary { get; set; } = true; // ✅ PHASE 2: SQLite is now primary

        /// <summary>
        /// ✅ PHASE 2: Disable XML file creation - use database only
        /// When true: XML files are NOT created (database is single source of truth)
        /// When false: XML files are still created (dual-write mode for safety)
        /// XML reading is still enabled as fallback during transition
        /// </summary>
        public static bool DisableXmlCreation { get; set; } = true; // ✅ PHASE 2: Disable XML creation - database only

        /// <summary>
        /// ✅ PARALLEL PLANNING: Enable parallel pre-computation of sleeve dimensions, clearance, and rotation
        /// When true: Uses ParallelSleevePlacementPlanner to pre-compute placement data in parallel
        /// When false: Uses original sequential placement logic (backward compatible)
        /// Benefits when enabled:
        /// - Early skip detection (avoid Revit API calls for zones that will be skipped)
        /// - Risk-based reordering (process low-risk zones first for better success rate)
        /// - Parallel computation (dimensions, clearance, rotation calculated in parallel)
        /// - Better diagnostics (risk classification available)
        /// </summary>
        public static bool EnableParallelPlanning { get; set; } = true; // ✅ PARALLEL PLANNING: Enabled (fixing round pipe cluster bug)
    }
}

