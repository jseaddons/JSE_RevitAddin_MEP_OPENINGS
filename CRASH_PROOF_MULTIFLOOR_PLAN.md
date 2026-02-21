# Crash-Proof Multi-Floor Implementation Plan

## Current State Analysis

### Multi-Floor Architecture
- **FloorBatchProcessor**: Orchestrates processing across multiple floors with chunking
- **SingleFloorProcessor**: Processes one floor per transaction
- **BulkPlacementService**: Handles actual sleeve placement

### Existing Safety Mechanisms
1. **Per-floor transaction with rollback**
2. **Checkpoint/resume capability** (checkpoint saved after each chunk)
3. **Basic exception handling** with rollback
4. **Memory management** (GC forced between chunks)

### Crash Risk Areas Identified
1. **No IFailuresPreprocessor** - Revit warnings/errors can crash the operation
2. **No timeout protection** - Long-running floors can hang indefinitely
3. **No element validation** - Elements may be deleted/modified during processing
4. **Memory pressure** - Large floors can cause OOM
5. **No pre-flight checks** - Invalid state not detected before processing

---

## Implementation Plan

### Phase 1: Transaction Safety (Critical)

#### 1.1 Create MultiFloorFailurePreprocessor
```csharp
public class MultiFloorFailurePreprocessor : IFailuresPreprocessor
{
    private readonly string _floorName;
    private readonly Action<string> _logger;
    
    public FailureProcessingResult PreprocessFailures(FailuresAccessor fa)
    {
        var failures = fa.GetFailureMessages();
        foreach (var f in failures)
        {
            var severity = f.GetSeverity();
            var desc = f.GetDescriptionText();
            
            if (severity == FailureSeverity.Warning)
            {
                _logger?.Invoke($"[Floor {_floorName}] Swallowing warning: {desc}");
                fa.DeleteWarning(f);
            }
            else if (severity == FailureSeverity.Error)
            {
                // Handle known harmless errors
                if (IsHarmlessError(desc))
                {
                    _logger?.Invoke($"[Floor {_floorName}] Treating harmless error as warning: {desc}");
                    fa.DeleteWarning(f);
                }
                else
                {
                    _logger?.Invoke($"[Floor {_floorName}] CRITICAL ERROR: {desc}");
                    // Let it fail - will be caught by outer exception handler
                }
            }
        }
        return FailureProcessingResult.Continue;
    }
    
    private bool IsHarmlessError(string desc)
    {
        var harmlessPatterns = new[]
        {
            "duplicate",
            "coincident",
            "slightly off axis",
            "very small arc",
            "instability detected"
        };
        return harmlessPatterns.Any(p => desc.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);
    }
}
```

#### 1.2 Apply to SingleFloorProcessor
```csharp
public FloorProcessingResult ProcessFloor(Level level, OpeningFilter filter)
{
    using (var transaction = new Transaction(_doc, $"Process Floor {level.Name}"))
    {
        // Set up failure handling
        var failureOptions = transaction.GetFailureHandlingOptions();
        failureOptions.SetFailuresPreprocessor(new MultiFloorFailurePreprocessor(level.Name, msg => 
            SafeFileLogger.SafeAppendText("multifloor.log", $"[{DateTime.Now}] {msg}\n")));
        failureOptions.SetClearAfterRollback(true);
        transaction.SetFailureHandlingOptions(failureOptions);
        
        transaction.Start();
        // ... rest of method
    }
}
```

---

### Phase 2: Timeout Protection (Critical)

#### 2.1 Enhance CrashSafeExecutor for Multi-Floor
```csharp
public class MultiFloorCrashProtection
{
    private readonly Stopwatch _stopwatch = new();
    private readonly int _floorTimeoutMs;
    private readonly int _totalTimeoutMs;
    
    public MultiFloorCrashProtection(int floorTimeoutMinutes = 1, int totalTimeoutMinutes = 10)
    {
        _floorTimeoutMs = floorTimeoutMinutes * 60 * 1000;
        _totalTimeoutMs = totalTimeoutMinutes * 60 * 1000;
    }
    
    public bool CheckFloorTimeout(string floorName)
    {
        if (_stopwatch.ElapsedMilliseconds > _floorTimeoutMs)
        {
            SafeFileLogger.SafeAppendText("multifloor.log",
                $"[{DateTime.Now}] ⏱ FLOOR TIMEOUT: {floorName} exceeded {_floorTimeoutMs/60000} minutes\n");
            return true;
        }
        return false;
    }
    
    public bool CheckTotalTimeout()
    {
        if (_stopwatch.ElapsedMilliseconds > _totalTimeoutMs)
        {
            SafeFileLogger.SafeAppendText("multifloor.log",
                $"[{DateTime.Now}] ⏱ TOTAL TIMEOUT: Exceeded {_totalTimeoutMs/60000} minutes\n");
            return true;
        }
        return false;
    }
}
```

#### 2.2 Integrate into FloorBatchProcessor
```csharp
public MultiFloorResult ProcessFloors(List<Level> levels, OpeningFilter filter, int chunkSize = 5)
{
    var crashProtection = new MultiFloorCrashProtection(floorTimeoutMinutes: 1, totalTimeoutMinutes: 10);
    
    foreach (var level in chunk)
    {
        if (crashProtection.CheckTotalTimeout())
        {
            throw new TimeoutException("Total multi-floor processing time exceeded");
        }
        
        var floorTask = Task.Run(() =>
        {
            var cts = new CancellationTokenSource(crashProtection.FloorTimeoutMs);
            // Process with timeout
        });
    }
}
```

---

### Phase 3: Element Validation (High Priority)

#### 3.1 Pre-Flight Validation
```csharp
public class MultiFloorValidation
{
    public ValidationResult ValidateFloor(Level level, Document doc)
    {
        var errors = new List<string>();
        
        // Check level is valid
        if (level == null || level.Id == null || level.Id.IntegerValue <= 0)
        {
            errors.Add("Invalid level object");
        }
        
        // Check level has geometry
        var levelElevation = level.ProjectElevation;
        if (double.IsNaN(levelElevation) || double.IsInfinity(levelElevation))
        {
            errors.Add("Level has invalid elevation");
        }
        
        // Check document is valid
        if (doc == null || doc.IsReadOnly)
        {
            errors.Add("Document is null or read-only");
        }
        
        return new ValidationResult { IsValid = errors.Count == 0, Errors = errors };
    }
    
    public ValidationResult ValidateElements(List<ClashZone> clashes, Document doc)
    {
        var validClashes = new List<ClashZone>();
        
        foreach (var clash in clashes)
        {
            // Verify MEP element still exists
            if (clash.MepElementIds?.Any() == true)
            {
                bool allExist = clash.MepElementIds.All(id => 
                    doc.GetElement(new ElementId(id)) != null);
                if (!allExist)
                {
                    SafeFileLogger.SafeAppendText("multifloor.log",
                        $"[{DateTime.Now}] ⚠️ Clash {clash.Id} has deleted MEP elements, skipping\n");
                    continue;
                }
            }
            validClashes.Add(clash);
        }
        
        return new ValidationResult { IsValid = true, ValidClashes = validClashes };
    }
}
```

---

### Phase 4: Memory Pressure Handling (Medium Priority)

#### 4.1 Memory Monitor
```csharp
public class MemoryPressureMonitor
{
    private readonly long _memoryThresholdBytes;
    
    public MemoryPressureMonitor(int thresholdMB = 2048)
    {
        _memoryThresholdBytes = thresholdMB * 1024 * 1024;
    }
    
    public bool IsUnderMemoryPressure()
    {
        var usedMemory = GC.GetTotalMemory(false);
        return usedMemory > _memoryThresholdBytes;
    }
    
    public void CheckAndPauseIfNeeded(string operation)
    {
        if (IsUnderMemoryPressure())
        {
            SafeFileLogger.SafeAppendText("multifloor.log",
                $"[{DateTime.Now}] 🛑 Memory pressure detected during {operation}. Forcing GC...\n");
            
            GC.Collect(2, GCCollectionMode.Aggressive, blocking: true);
            GC.WaitForPendingFinalizers();
            
            // Brief pause to let system stabilize
            Thread.Sleep(500);
        }
    }
}
```

---

### Phase 5: Checkpoint & Recovery Enhancements (Medium Priority)

#### 5.1 Enhanced Checkpoint with Error Tracking
```csharp
public class EnhancedCheckpointManager : CheckpointManager
{
    public void SaveCheckpointWithErrors(MultiFloorProgress progress, List<string> failedFloors)
    {
        var checkpoint = new MultiFloorCheckpoint
        {
            ProcessedFloors = progress.CompletedFloors.Select(f => f.Name).ToList(),
            FailedFloors = failedFloors,
            Timestamp = DateTime.Now,
            Version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown"
        };
        
        SaveCheckpoint(checkpoint);
    }
    
    public List<Level> GetRemainingLevels(List<Level> allLevels, out List<string> failedFloorNames)
    {
        var checkpoint = LoadCheckpoint();
        failedFloorNames = checkpoint?.FailedFloors ?? new List<string>();
        
        if (checkpoint == null) return allLevels;
        
        return allLevels
            .Where(l => !checkpoint.ProcessedFloors.Contains(l.Name))
            .Where(l => !failedFloorNames.Contains(l.Name)) // Skip previously failed
            .ToList();
    }
}
```

---

### Phase 6: Circuit Breaker Pattern (Low Priority)

#### 6.1 Fail-Fast on Repeated Failures
```csharp
public class CircuitBreaker
{
    private readonly int _failureThreshold;
    private int _consecutiveFailures = 0;
    
    public bool ShouldProcess(string floorName)
    {
        if (_consecutiveFailures >= _failureThreshold)
        {
            SafeFileLogger.SafeAppendText("multifloor.log",
                $"[{DateTime.Now}] 🚫 CIRCUIT BREAKER: Stopping after {_consecutiveFailures} consecutive failures\n");
            return false;
        }
        return true;
    }
    
    public void RecordSuccess() => _consecutiveFailures = 0;
    public void RecordFailure() => _consecutiveFailures++;
}
```

---

## Implementation Order

### Week 1: Critical (Must Have)
1. **MultiFloorFailurePreprocessor** - Prevents Revit warnings from crashing
2. **Transaction safety enhancements** - SetClearAfterRollback, proper options
3. **Timeout protection** - Per-floor and total timeout

### Week 2: High Priority
4. **Element validation** - Pre-flight checks, element existence validation
5. **Enhanced logging** - Better diagnostics for crashes

### Week 3: Medium Priority
6. **Memory pressure handling** - GC management, pauses
7. **Enhanced checkpointing** - Error tracking, skip failed floors

### Week 4: Polish
8. **Circuit breaker** - Fail-fast on repeated failures
9. **User feedback** - Progress dialogs, cancel buttons

---

## Testing Strategy

1. **Stress Test**: Process 50+ floors with large datasets
2. **Failure Injection**: Deliberately cause warnings/errors
3. **Timeout Test**: Simulate slow operations
4. **Memory Test**: Process with limited memory
5. **Recovery Test**: Interrupt and resume

---

## Success Criteria

- [ ] No crashes due to Revit warnings/errors
- [ ] Graceful timeout handling with partial results saved
- [ ] Automatic retry for transient failures
- [ ] Resume capability after crash
- [ ] Memory usage stays below 2GB
- [ ] Failed floors don't block entire batch
