# Refresh Logs Troubleshooting - File Investigation List

## 🔴 CRITICAL FILES - Check These First

### 1. `Services/SafeFileLogger.cs` ⚠️ PRIMARY SUSPECT
**Why it matters:** This is the main logging wrapper that all refresh logs go through.

**Check these lines:**
- **Line 370**: `if (DeploymentConfiguration.DeploymentMode) return;` - This SKIPS all logging if DeploymentMode is true
- **Line 431**: `if (!DeploymentConfiguration.DeploymentMode)` - Double check inside lock
- **Line 403**: `GetLogDirectory()` - May fail silently if directory creation fails
- **Line 223-290**: `WriteDiagnosticLogInternal` - This writes diagnostic logs but may fail silently
- **Line 296-362**: `WriteDiagnosticLog` - Another diagnostic log writer that may fail

**Potential Issues:**
- Directory creation fails silently (lines 190-217)
- Exceptions are caught and swallowed (lines 443-495)
- Diagnostic logs may fail if AppData directory is inaccessible
- Fallback to Desktop may not be working

**What to check:**
1. Does `%AppData%\Roaming\JSE_MEP_Openings\Logs\` exist?
2. Is it writable?
3. Check Desktop for `JSE_MEP_Openings_Diagnostic.log` (fallback location)
4. Check for `safefilelogger_diagnostic.log` in AppData Logs directory
5. Check for `safefilelogger_errors.log` in AppData Logs directory

---

### 2. `Services/DeploymentConfiguration.cs` ⚠️ CRITICAL SWITCH
**Why it matters:** If `DeploymentMode = true`, ALL logging is disabled.

**Check these lines:**
- **Line 24**: `public static bool DeploymentMode { get; set; } = false;`
- **Current status:** Should be `false` for logs to appear

**What to check:**
1. Is `DeploymentMode` set to `false`?
2. Is it being changed elsewhere in code?
3. Search codebase for `DeploymentConfiguration.DeploymentMode = true`

---

### 3. `Services/RefreshService.cs` ⚠️ WHERE LOGS ARE WRITTEN
**Why it matters:** This is where refresh logs are actually written.

**Check these lines:**
- **Line 481**: `string refreshLogName = $"Refresh_{timestamp}.log";` - Log file name
- **Line 485-492**: Early log writes to verify SafeFileLogger is working
- **Line 640**: Memory profiler initialization logging
- **Line 636**: `_memoryProfiler = new MemoryProfiler(...)` - May fail silently

**What to check:**
1. Are `SafeFileLogger.SafeAppendText()` calls happening?
2. Is `refreshLogName` variable being passed correctly?
3. Check if exceptions are being caught and swallowed around logging calls
4. Look for try-catch blocks that silently swallow exceptions around logging

**Search for:**
- `catch (Exception` blocks that might swallow logging errors
- Any code that calls `SafeFileLogger.SafeAppendText` with wrong parameters

---

### 4. `Services/OptimizationFlags.cs` ⚠️ DIAGNOSTIC MODE
**Why it matters:** Memory profiling logs depend on `UseDiagnosticMode`.

**Check these lines:**
- **Line 79**: `public static bool UseDiagnosticMode { get; set; } = true;`
- **Current status:** Should be `true` for memory profiling logs

**What to check:**
1. Is `UseDiagnosticMode` set to `true`?
2. Is it being changed elsewhere?

---

### 5. `Services/MemoryProfiler.cs` ⚠️ MEMORY LOGS
**Why it matters:** Memory profiling logs go through this class.

**Check these lines:**
- Look for `TakeSnapshot` method - does it check `DeploymentMode`?
- Look for `WriteToFile` or similar methods - do they use `SafeFileLogger`?
- Check if exceptions are caught and swallowed

**What to check:**
1. Does `MemoryProfiler.TakeSnapshot()` call `SafeFileLogger`?
2. Does it check `DeploymentMode` or `UseDiagnosticMode`?
3. Are exceptions being caught and swallowed?

---

## 🟡 SECONDARY FILES - Check These Next

### 6. `Services/LoggingConfiguration.cs`
**Why it matters:** Has additional logging switches.

**Check these lines:**
- **Line 79**: `if (DeploymentConfiguration.DeploymentMode) return;`
- **Line 12**: `DisableAllHardcodedLogging` - Should be `false`
- **Line 17**: `EnableRefreshButton` - Should be checked if refresh logs are disabled

**What to check:**
1. Is `DisableAllHardcodedLogging = false`?
2. Is `EnableRefreshButton = true`? (if refresh logs go through this)

---

### 7. `Services/DebugLogger.cs`
**Why it matters:** Some refresh logs may go through DebugLogger.

**Check these lines:**
- **Line 19**: `public static bool IsEnabled = true;`
- **Line 78**: `IsLoggingEnabledForCurrentService()` - Checks deployment mode
- Look for `DeploymentConfiguration.DeploymentMode` checks

**What to check:**
1. Is `DebugLogger.IsEnabled = true`?
2. Are DebugLogger calls being made from RefreshService?
3. Check if `IsLoggingEnabledForCurrentService()` returns false

---

## 🔍 INVESTIGATION STEPS

### Step 1: Check Deployment Mode
```csharp
// Check DeploymentConfiguration.cs
public static bool DeploymentMode { get; set; } = false; // Should be FALSE
```

### Step 2: Check Diagnostic Mode
```csharp
// Check OptimizationFlags.cs
public static bool UseDiagnosticMode { get; set; } = true; // Should be TRUE
```

### Step 3: Check Log Directory
```powershell
# Run in PowerShell
Test-Path "$env:APPDATA\Roaming\JSE_MEP_Openings\Logs"
# Should return True

# Check if writable
Test-Path "$env:APPDATA\Roaming\JSE_MEP_Openings\Logs\test.tmp"
New-Item "$env:APPDATA\Roaming\JSE_MEP_Openings\Logs\test.tmp" -Force
Remove-Item "$env:APPDATA\Roaming\JSE_MEP_Openings\Logs\test.tmp"
```

### Step 4: Check Diagnostic Logs
Look for these files:
- `%AppData%\Roaming\JSE_MEP_Openings\Logs\safefilelogger_diagnostic.log`
- `%AppData%\Roaming\JSE_MEP_Openings\Logs\safefilelogger_errors.log`
- `%Desktop%\JSE_MEP_Openings_Diagnostic.log` (fallback location)

### Step 5: Check Refresh Logs
Look for:
- `%AppData%\Roaming\JSE_MEP_Openings\Logs\Refresh_*.log`
- `%AppData%\Roaming\JSE_MEP_Openings\Logs\refresh_memory_profiling_*.log`

### Step 6: Add Debugging Code
Add this at the START of `RefreshService.ExecuteRefreshInternal`:
```csharp
// Add after line 468 (after settings declaration)
try
{
    string testPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "REFRESH_LOG_TEST.txt");
    File.WriteAllText(testPath, $"Refresh started at {DateTime.Now}\nDeploymentMode={DeploymentConfiguration.DeploymentMode}\nDiagnosticMode={OptimizationFlags.UseDiagnosticMode}\n");
}
catch (Exception testEx)
{
    // If this fails, file system access is the problem
}
```

---

## 🐛 COMMON FAILURE POINTS

### Failure Point 1: DeploymentMode Check
**Location:** `SafeFileLogger.cs:370`
**Issue:** If `DeploymentMode = true`, logs are skipped
**Fix:** Ensure `DeploymentConfiguration.DeploymentMode = false`

### Failure Point 2: Directory Creation Fails
**Location:** `SafeFileLogger.cs:190-217`
**Issue:** `TryCreateDirectory` returns false silently
**Fix:** Check AppData directory permissions

### Failure Point 3: Exception Swallowed
**Location:** `SafeFileLogger.cs:443-495`
**Issue:** Exceptions are caught and logged to diagnostic log, but diagnostic log write also fails
**Fix:** Check diagnostic log files (Desktop fallback)

### Failure Point 4: Memory Profiler Not Initialized
**Location:** `RefreshService.cs:636`
**Issue:** `_memoryProfiler` may not be initialized if constructor fails
**Fix:** Check if `MemoryProfiler` constructor throws exceptions

### Failure Point 5: Log File Name Mismatch
**Location:** `RefreshService.cs:481`
**Issue:** `refreshLogName` may not match what's being searched for
**Fix:** Verify log file name pattern matches search pattern

---

## 📋 CHECKLIST FOR USER INVESTIGATION

- [ ] `DeploymentConfiguration.cs` line 24: `DeploymentMode = false`
- [ ] `OptimizationFlags.cs` line 79: `UseDiagnosticMode = true`
- [ ] `SafeFileLogger.cs` line 370: No early return
- [ ] AppData directory exists: `%AppData%\Roaming\JSE_MEP_Openings\Logs\`
- [ ] AppData directory is writable
- [ ] Desktop fallback log exists: `%Desktop%\JSE_MEP_Openings_Diagnostic.log`
- [ ] Diagnostic log exists: `%AppData%\Roaming\JSE_MEP_Openings\Logs\safefilelogger_diagnostic.log`
- [ ] Error log exists: `%AppData%\Roaming\JSE_MEP_Openings\Logs\safefilelogger_errors.log`
- [ ] No exceptions in RefreshService around logging calls
- [ ] `MemoryProfiler` is initialized correctly
- [ ] Search codebase for `DeploymentConfiguration.DeploymentMode = true`

---

## 🔗 RELATED FILES (Less Likely But Worth Checking)

- `Services/BatchedLogger.cs` - May have its own logging switches
- `Commands/UniversalSleevePlacementCommand.cs` - May have logging switches
- Any file that modifies `DeploymentConfiguration.DeploymentMode` at runtime

---

## 💡 RECOMMENDATION

**Start with these 3 files in order:**
1. `Services/DeploymentConfiguration.cs` - Verify `DeploymentMode = false`
2. `Services/SafeFileLogger.cs` - Verify directory creation and exception handling
3. `Services/RefreshService.cs` - Verify logging calls are actually being made

**Then check:**
4. AppData directory permissions
5. Diagnostic log files (Desktop fallback)
6. Exception handling that might swallow errors

