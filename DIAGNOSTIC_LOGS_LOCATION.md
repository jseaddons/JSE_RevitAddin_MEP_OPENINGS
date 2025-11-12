# Diagnostic Logs Location Guide

## ✅ ALL LOGGING NOW GOES THROUGH SafeFileLogger (WRAPPED)

All diagnostic logging has been converted from `Debug.WriteLine` (which only appears in Visual Studio Debug Output) to proper log files using `SafeFileLogger.SafeAppendText()`.

## Diagnostic Log Files Location

All diagnostic logs are written to: **`%AppData%\Roaming\JSE_MEP_Openings\Logs\`**

### Main Diagnostic Log Files

1. **`safefilelogger_diagnostic.log`**
   - **Purpose**: Tracks every attempt to write logs through `SafeFileLogger`
   - **Content**: All `SafeFileLogger` operations (init, create directory, write attempts, successes, failures)
   - **Always Written**: Yes (even if deployment mode is ON, to debug logging issues)
   - **Format**: `[timestamp] [SAFEFILELOGGER] Operation=xxx, Message=xxx, Path=xxx, Success=true/false`

2. **`safefilelogger_errors.log`**
   - **Purpose**: Captures exceptions and errors from `SafeFileLogger`
   - **Content**: Stack traces, exception types, error messages when logging fails
   - **Always Written**: Yes (when errors occur)

3. **`refresh_YYYY-MM-DD_HH-mm-ss.log`**
   - **Purpose**: Main refresh operation log
   - **Content**: All refresh operation diagnostics, memory profiling, intersection detection
   - **Written When**: Deployment mode OFF OR Diagnostic mode ON

4. **`refresh_memory_profiling_YYYY-MM-DD_HH-mm-ss.log`**
   - **Purpose**: Detailed memory profiling snapshots
   - **Content**: Memory usage per phase, clash zone memory analysis
   - **Written When**: Deployment mode OFF OR Diagnostic mode ON

5. **`performance.log`**
   - **Purpose**: Performance timing data
   - **Content**: Start/end times for refresh operations
   - **Written When**: Deployment mode OFF OR Diagnostic mode ON

## How to Check Diagnostic Logs

1. **Open File Explorer**
2. **Navigate to**: `%AppData%\Roaming\JSE_MEP_Openings\Logs\`
   - Or manually: `C:\Users\<YourUsername>\AppData\Roaming\JSE_MEP_Openings\Logs\`
3. **Check these files**:
   - `safefilelogger_diagnostic.log` - Shows all logging attempts
   - `safefilelogger_errors.log` - Shows any logging errors
   - `refresh_*.log` - Shows refresh operation logs (if deployment mode OFF or diagnostic mode ON)
   - `refresh_memory_profiling_*.log` - Shows memory profiling (if deployment mode OFF or diagnostic mode ON)

## What Changed

### Before (❌ BAD)
- `Debug.WriteLine()` calls → Only visible in Visual Studio Debug Output window
- Direct `File.AppendAllText()` calls → Bypassed deployment mode checks
- No diagnostic logging → Couldn't debug why logs weren't appearing

### After (✅ GOOD)
- All diagnostic logging → `SafeFileLogger.SafeAppendText()` (wrapped)
- All `Debug.WriteLine()` → Removed or replaced with file logging
- Diagnostic log file → Always written to `safefilelogger_diagnostic.log`
- Error log file → Always written to `safefilelogger_errors.log` when errors occur

## Files Modified

1. **`Services/SafeFileLogger.cs`**
   - Removed all `Debug.WriteLine()` calls
   - Added `WriteDiagnosticLogInternal()` method for internal diagnostics
   - All diagnostics now write to `safefilelogger_diagnostic.log` file

2. **`Services/RefreshService.cs`**
   - Removed all `Debug.WriteLine()` calls
   - All diagnostics now use `SafeFileLogger.SafeAppendText()` to write to refresh log file

## Example Diagnostic Log Entry

```
[2025-01-04 14:30:45.123] [SAFEFILELOGGER] Operation=INIT, Message=Log directory initialized: C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs, Directory exists: True, Path=C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs, Success=True
```

## Troubleshooting

If you don't see refresh logs:
1. Check `safefilelogger_diagnostic.log` - Shows all logging attempts
2. Check `safefilelogger_errors.log` - Shows any errors
3. Verify `DeploymentConfiguration.DeploymentMode` is `false` OR `OptimizationFlags.UseDiagnosticMode` is `true`
4. Check file permissions on the Logs directory

