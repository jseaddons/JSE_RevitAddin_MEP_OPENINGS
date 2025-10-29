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
        /// This reduces memory usage by 15-30 KB per clash zone (logging strings).
        /// 
        /// When enabled:
        /// - All SafeFileLogger calls are skipped (no I/O)
        /// - All DebugLogger calls are skipped (no file writes)
        /// - All BatchedLogger calls are skipped (no string allocations)
        /// - Memory per clash zone should drop from ~83 KB to ~55-65 KB
        /// </summary>
        public static bool DeploymentMode { get; set; } = true; // ✅ TESTING: Enabled to disable logging except memory profiling
    }
}

