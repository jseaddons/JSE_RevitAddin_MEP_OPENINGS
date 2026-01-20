# Refresh Log Debugging Guide

## Where Refresh Logs Are Written

### Primary Location
**`C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs`**

### Expected Log Files
1. **Refresh Logs**: `Refresh_YYYY-MM-DD_HH-mm-ss.log`
   - Example: `Refresh_2025-01-15_14-30-45.log`
   - Contains: All refresh operations, flag syncing, clash zone detection

2. **Memory Profiling Logs**: `refresh_memory_profiling_YYYY-MM-DD_HH-mm-ss.log`
   - Example: `refresh_memory_profiling_2025-01-15_14-30-45.log`
   - Contains: Memory usage analysis per clash zone

3. **Error Logs**: `safefilelogger_errors.log`
   - Contains: Errors from SafeFileLogger if file writes fail

4. **Performance Log**: `performance.log`
   - Contains: Performance timing data

---

## How to Debug Missing Logs

### Step 1: Check Debug Output Window
When refresh runs, check **Visual Studio Debug Output Window** for:
- `[SafeFileLogger] ATTEMPTING: fileName=...`
- `[SafeFileLogger] SUCCESS: Written to ...`
- `[SafeFileLogger] SKIPPED: ...`
- `[REFRESH] Log directory: ...`
- `[REFRESH] File exists after write: ...`

### Step 2: Check DeploymentMode
Verify `DeploymentConfiguration.DeploymentMode` is `false`:
- File: `Services/DeploymentConfiguration.cs`
- Line 24: Should be `public static bool DeploymentMode { get; set; } = false;`
- If `true`, ALL logging is disabled

### Step 3: Check Log Directory
Run this PowerShell command to check:
```powershell
$logDir = [System.IO.Path]::Combine([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData), "JSE_MEP_Openings", "Logs")
Write-Host "Log Directory: $logDir"
Write-Host "Directory Exists: $(Test-Path $logDir)"
Write-Host "Directory Writable: $([System.IO.Directory]::Exists($logDir))"
Get-ChildItem $logDir -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 10 Name, LastWriteTime
```

### Step 4: Check Error Log
Look for `safefilelogger_errors.log` in the log directory:
- This file contains detailed error information if SafeFileLogger fails
- Includes: Exception type, message, stack trace, deployment mode, log path

---

## Diagnostic Logging Added

### SafeFileLogger Diagnostic Output
Every call to `SafeFileLogger.SafeAppendText()` now logs to Debug Output:
- `[SafeFileLogger] ATTEMPTING: fileName=..., messageLength=..., DeploymentMode=...`
- `[SafeFileLogger] LogPath=..., DirectoryExists=...`
- `[SafeFileLogger] SUCCESS: Written to ...` (if successful)
- `[SafeFileLogger] SKIPPED: ...` (if skipped)
- `[SafeFileLogger] No permission to write to log: ...` (if permission error)
- `[SafeFileLogger] Error writing to ...` (if other error)

### RefreshService Diagnostic Output
After writing first log entries, RefreshService checks if file exists:
- `[REFRESH] Log directory: ...`
- `[REFRESH] Full log path: ...`
- `[REFRESH] File exists after write: True/False`
- `[REFRESH] ✅ Log file exists! Size: ... bytes` (if exists)
- `[REFRESH] ❌ Log file does NOT exist: ...` (if missing)

---

## Common Issues and Solutions

### Issue 1: DeploymentMode = true
**Symptoms**: No logs appear, Debug output shows `[SafeFileLogger] SKIPPED: DeploymentMode=true`
**Solution**: Set `DeploymentConfiguration.DeploymentMode = false`

### Issue 2: Permission Error
**Symptoms**: Debug output shows `[SafeFileLogger] No permission to write to log: ...`
**Solution**: Check folder permissions for `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs`

### Issue 3: Directory Not Created
**Symptoms**: Debug output shows `[SafeFileLogger] Failed to create directory ...`
**Solution**: Check AppData folder permissions

### Issue 4: Log File Not Created
**Symptoms**: Debug output shows `[REFRESH] ❌ Log file does NOT exist: ...`
**Solution**: Check Debug output for SafeFileLogger errors, check `safefilelogger_errors.log`

---

## Next Steps When Refresh Runs

1. **Open Visual Studio Debug Output Window** (View → Output → Show output from: Debug)
2. **Run Refresh** in Revit
3. **Look for these messages**:
   - `[REFRESH] === REFRESH METHOD STARTED ===`
   - `[SafeFileLogger] ATTEMPTING: fileName=Refresh_...`
   - `[SafeFileLogger] SUCCESS: Written to ...` OR `[SafeFileLogger] SKIPPED: ...`
   - `[REFRESH] File exists after write: True/False`
4. **Check log directory** for files created
5. **If files don't exist**, check `safefilelogger_errors.log` for error details

---

## Summary

**Log Location**: `C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Logs`

**Log Files**:
- `Refresh_YYYY-MM-DD_HH-mm-ss.log` - Main refresh log
- `refresh_memory_profiling_YYYY-MM-DD_HH-mm-ss.log` - Memory profiling
- `safefilelogger_errors.log` - SafeFileLogger errors
- `performance.log` - Performance timing

**Debug Output**: Check Visual Studio Debug Output Window for diagnostic messages

**DeploymentMode**: Must be `false` for logs to write

**Diagnostic Logging**: All logging attempts are now logged to Debug Output for troubleshooting

