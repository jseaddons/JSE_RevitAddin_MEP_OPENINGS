using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime;
using System.Runtime.InteropServices;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Memory management and timeout protection service
    /// Prevents crashes on large/unclean files by:
    /// - Monitoring memory usage
    /// - Enforcing timeouts (5 minutes default)
    /// - Gracefully exiting when limits are reached
    /// - Force garbage collection when needed
    /// </summary>
    public class MemoryManager : IDisposable
    {
        private readonly long _maxMemoryBytes;
        private readonly TimeSpan _maxExecutionTime;
        private readonly Stopwatch _executionTimer;
        private readonly PerformanceCounter _memoryCounter;
        private bool _disposed = false;
        private CancellationTokenSource _cancellationTokenSource;

        // Default limits
        // ⚠️ NOTE: Even with 64GB system RAM, individual Revit processes have practical limits
        // Modern Revit (64-bit) can typically use 6-8GB per process before issues
        // We set a conservative 6GB limit to prevent crashes while allowing large models
        private const long DEFAULT_MAX_MEMORY_MB = 6144; // 6GB limit (reasonable for large models)
        private const int DEFAULT_TIMEOUT_MINUTES = 5;

        /// <summary>
        /// Automatically calculates optimal memory limit based on system RAM
        /// Formula: 10% of total RAM, capped at 12GB (for Revit process limits)
        /// Minimum: 4GB (for systems with < 40GB RAM)
        /// Uses Win32 API to detect RAM without requiring System.Management assembly
        /// </summary>
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GlobalMemoryStatusEx(ref MEMORYSTATUSEX lpBuffer);

        [StructLayout(LayoutKind.Sequential)]
        private struct MEMORYSTATUSEX
        {
            public uint dwLength;
            public uint dwMemoryLoad;
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
        }

        public static long CalculateOptimalMemoryLimit()
        {
            try
            {
                // Get total system RAM in bytes using Win32 API (no external dependencies)
                MEMORYSTATUSEX memStatus = new MEMORYSTATUSEX();
                memStatus.dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX));
                
                if (GlobalMemoryStatusEx(ref memStatus))
                {
                    long totalRamBytes = (long)memStatus.ullTotalPhys;
                    long totalRamMB = totalRamBytes / (1024 * 1024);
                    
                    // Formula: 10% of total RAM, with reasonable bounds
                    long optimalMB = (long)(totalRamMB * 0.10);
                    
                    // Cap at 12GB (Revit process practical limit)
                    long maxMB = 12 * 1024; // 12GB
                    optimalMB = Math.Min(optimalMB, maxMB);
                    
                    // Minimum: 4GB for systems with < 40GB RAM
                    long minMB = 4 * 1024; // 4GB
                    optimalMB = Math.Max(optimalMB, minMB);
                    
                    SafeFileLogger.SafeAppendText("memory_manager.log", 
                        $"System RAM detected: {totalRamMB / 1024}GB. Calculated optimal limit: {optimalMB / 1024}GB (10% of RAM, capped at 12GB)");
                    
                    return optimalMB;
                }
                else
                {
                    SafeFileLogger.SafeAppendText("memory_manager.log", 
                        $"GlobalMemoryStatusEx failed. Using default {DEFAULT_MAX_MEMORY_MB / 1024}GB limit.");
                }
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"Failed to detect system RAM: {ex.Message}. Using default {DEFAULT_MAX_MEMORY_MB / 1024}GB limit.");
            }
            
            // Fallback to default
            return DEFAULT_MAX_MEMORY_MB;
        }

        public MemoryManager(long? maxMemoryMB = null, int timeoutMinutes = DEFAULT_TIMEOUT_MINUTES)
        {
            // Use provided limit, or auto-detect optimal limit based on system RAM
            long actualLimitMB = maxMemoryMB ?? CalculateOptimalMemoryLimit();
            _maxMemoryBytes = actualLimitMB * 1024 * 1024; // Convert MB to bytes
            _maxExecutionTime = TimeSpan.FromMinutes(timeoutMinutes);
            _executionTimer = Stopwatch.StartNew();
            _cancellationTokenSource = new CancellationTokenSource();

            // Try to initialize performance counter (may fail on some systems)
            try
            {
                _memoryCounter = new PerformanceCounter("Process", "Working Set", Process.GetCurrentProcess().ProcessName);
            }
            catch
            {
                _memoryCounter = null; // Continue without performance counter
            }

            SafeFileLogger.SafeAppendText("memory_manager.log", 
                $"MemoryManager initialized: MaxMemory={maxMemoryMB}MB, Timeout={timeoutMinutes} minutes");
        }

        /// <summary>
        /// Checks if memory or time limits have been exceeded. Throws OperationCanceledException if so.
        /// Call this periodically during long operations.
        /// </summary>
        public void CheckLimits()
        {
            if (_disposed)
                return;

            // Check timeout FIRST - most critical
            if (_executionTimer.Elapsed >= _maxExecutionTime)
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"⏰ TIMEOUT: Execution time exceeded {_maxExecutionTime.TotalMinutes} minutes (Elapsed: {_executionTimer.Elapsed:mm\\:ss})");
                
                _cancellationTokenSource?.Cancel();
                
                // Force cleanup before throwing
                try { ForceCleanup(); } catch { }
                
                throw new TimeoutException(
                    $"Operation exceeded maximum execution time of {_maxExecutionTime.TotalMinutes} minutes (Elapsed: {_executionTimer.Elapsed:mm\\:ss}). " +
                    $"This usually indicates a very large or problematic file. " +
                    $"Please try with a smaller file, reduce section box, or contact support.");
            }

            // Check memory - ENFORCE STRICT LIMITS
            long currentMemory = GetCurrentMemoryUsage();
            double memoryPercent = (double)currentMemory / _maxMemoryBytes * 100.0;
            
            // ⚠️ CRITICAL: Enforce memory limits strictly
            if (currentMemory > _maxMemoryBytes)
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"💾 MEMORY LIMIT EXCEEDED: Current={currentMemory / 1024 / 1024}MB ({memoryPercent:F1}%), Max={_maxMemoryBytes / 1024 / 1024}MB");
                
                // Try to free memory ONCE
                bool freed = TryFreeMemory();
                
                // Re-check after cleanup
                long memoryAfterCleanup = GetCurrentMemoryUsage();
                if (memoryAfterCleanup > _maxMemoryBytes * 0.95) // Still > 95% of limit
                {
                    SafeFileLogger.SafeAppendText("memory_manager.log", 
                        $"💾 MEMORY CLEANUP FAILED: Still at {memoryAfterCleanup / 1024 / 1024}MB after cleanup");
                    
                    // If we can't free enough memory, STOP operation immediately
                    _cancellationTokenSource?.Cancel();
                    
                    try { ForceCleanup(); } catch { }
                    
                    throw new OutOfMemoryException(
                        $"Memory usage exceeded limit of {_maxMemoryBytes / 1024 / 1024}MB. " +
                        $"Current: {currentMemory / 1024 / 1024}MB ({memoryPercent:F1}%). " +
                        $"After cleanup: {memoryAfterCleanup / 1024 / 1024}MB. " +
                        $"Note: Even with 64GB system RAM, Revit processes have practical limits (~6-8GB per process). " +
                        $"Please try with a smaller file, reduce section box, close other applications, or contact support.");
                }
                else
                {
                    SafeFileLogger.SafeAppendText("memory_manager.log", 
                        $"💾 MEMORY CLEANUP SUCCESS: Freed from {currentMemory / 1024 / 1024}MB to {memoryAfterCleanup / 1024 / 1024}MB");
                }
            }
            // Warn at 80% to give early warning
            else if (memoryPercent > 80.0)
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"⚠️ MEMORY WARNING: {currentMemory / 1024 / 1024}MB ({memoryPercent:F1}%) - approaching limit");
            }

            // Check if cancellation was requested
            if (_cancellationTokenSource.Token.IsCancellationRequested)
            {
                throw new OperationCanceledException(
                    "Operation was cancelled due to resource limits.");
            }
        }

        /// <summary>
        /// Gets the cancellation token - pass this to long-running operations
        /// </summary>
        public CancellationToken CancellationToken => _cancellationTokenSource.Token;

        /// <summary>
        /// Gets current memory usage in bytes
        /// </summary>
        private long GetCurrentMemoryUsage()
        {
            try
            {
                // Method 1: GC.GetTotalMemory (managed memory)
                long managedMemory = GC.GetTotalMemory(false);
                
                // Method 2: Process.WorkingSet (total process memory)
                long processMemory = Process.GetCurrentProcess().WorkingSet64;
                
                // Use the larger value for safety
                return Math.Max(managedMemory, processMemory);
            }
            catch
            {
                // Fallback: Use GC memory only
                return GC.GetTotalMemory(false);
            }
        }

        /// <summary>
        /// Tries to free memory by forcing garbage collection
        /// </summary>
        private bool TryFreeMemory()
        {
            try
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    "🧹 Attempting memory cleanup...");
                
                // Force full garbage collection
                GC.Collect(2, GCCollectionMode.Forced, true);
                GC.WaitForPendingFinalizers();
                GC.Collect(2, GCCollectionMode.Forced, true);
                
                long memoryAfter = GetCurrentMemoryUsage();
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"🧹 Memory after cleanup: {memoryAfter / 1024 / 1024}MB");
                
                // Check if we're still over limit
                if (memoryAfter > _maxMemoryBytes * 0.9) // Still using > 90% of limit
                {
                    return false; // Cleanup didn't help enough
                }
                
                return true; // Cleanup successful
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"❌ Memory cleanup failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Forces immediate cleanup and memory release
        /// </summary>
        public void ForceCleanup()
        {
            try
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    "🔧 Force cleanup initiated...");
                
                // Full garbage collection
                GC.Collect(2, GCCollectionMode.Forced, true);
                GC.WaitForPendingFinalizers();
                GC.Collect(2, GCCollectionMode.Forced, true);
                
                // Compact large object heap (LOH)
                GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect(2, GCCollectionMode.Forced, true);
                
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"🔧 Force cleanup completed. Memory: {GetCurrentMemoryUsage() / 1024 / 1024}MB");
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"❌ Force cleanup error: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets current execution time elapsed
        /// </summary>
        public TimeSpan ElapsedTime => _executionTimer.Elapsed;

        /// <summary>
        /// Gets remaining time before timeout
        /// </summary>
        public TimeSpan RemainingTime => _maxExecutionTime - _executionTimer.Elapsed;

        /// <summary>
        /// Gets current memory usage in MB
        /// </summary>
        public long CurrentMemoryMB => GetCurrentMemoryUsage() / 1024 / 1024;

        /// <summary>
        /// Gets memory usage percentage (0-100)
        /// </summary>
        public double MemoryUsagePercent => (double)GetCurrentMemoryUsage() / _maxMemoryBytes * 100.0;

        /// <summary>
        /// Checks limits and returns status message
        /// </summary>
        public string GetStatus()
        {
            long memoryMB = CurrentMemoryMB;
            double percent = MemoryUsagePercent;
            TimeSpan remaining = RemainingTime;
            
            return $"Memory: {memoryMB}MB ({percent:F1}%), " +
                   $"Time: {_executionTimer.Elapsed:mm\\:ss}/{_maxExecutionTime:mm\\:ss}, " +
                   $"Remaining: {remaining:mm\\:ss}";
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            try
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"MemoryManager disposed. Final memory: {CurrentMemoryMB}MB, Elapsed: {ElapsedTime:mm\\:ss}");
                
                // Cancel any ongoing operations
                _cancellationTokenSource?.Cancel();
                
                // Cleanup
                _cancellationTokenSource?.Dispose();
                _memoryCounter?.Dispose();
                _executionTimer?.Stop();
                
                // Final cleanup
                ForceCleanup();
            }
            catch (Exception ex)
            {
                SafeFileLogger.SafeAppendText("memory_manager.log", 
                    $"❌ Error during MemoryManager disposal: {ex.Message}");
            }
            finally
            {
                _disposed = true;
            }
        }
    }
}

