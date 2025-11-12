# Coding Standards and Best Practices

## Purpose
This document establishes mandatory coding standards that MUST be followed before making any code changes. Always refer to this document FIRST before editing code.

---

## 🚨 CRITICAL: ALWAYS GET USER CONSENT AND DISCUSS BEFORE CODE EDITS

**MANDATORY RULE**: Before making ANY code changes, you MUST:
1. ✅ **DISCUSS** the proposed solution with the user
2. ✅ **EXPLAIN** what changes you plan to make and why
3. ✅ **GET EXPLICIT CONSENT** from the user before proceeding
4. ✅ **WAIT** for user approval before implementing

**NEVER**:
- ❌ Edit code without discussing first
- ❌ Assume what the user wants
- ❌ Make changes "because you think it's right"
- ❌ Skip consent and discussion steps
- ❌ **ADD FALLBACKS OR DEFENSIVE CHECKS WITHOUT USER CONSENT** - User is tired of fallbacks

**Exception**: Only when user explicitly says "yes" or "proceed" or "implement" after discussion.

---

## 🚨 STRICT: NO FALLBACKS WITHOUT USER CONSENT

**MANDATORY RULE**: Before adding ANY fallback logic, defensive checks, or error handling that continues execution:
1. ✅ **MUST ASK USER FIRST** - "Should I add a fallback for X?"
2. ✅ **GET EXPLICIT CONSENT** - User must approve fallback behavior
3. ✅ **STOP EXECUTION IF USER SAYS NO** - If user says no fallback, return error/empty result
4. ✅ **RESPECT USER DECISIONS** - If user says "no fallback", don't add fallback logic

**Examples of FALLBACKS that require consent**:
- ❌ Continuing without section box (user said section box is always present)
- ❌ Using default values when data is missing
- ❌ Falling back to entire model when section box is null
- ❌ Continuing execution when validation fails
- ❌ Using alternative methods when primary method fails

**User's Explicit Instruction**: "Before any fallback you want to check with me and get my consent - strict instruction"

**NEVER**:
- ❌ Add fallback logic without asking
- ❌ Assume fallback is needed
- ❌ Add "defensive" checks that continue execution
- ❌ Add error handling that falls back to alternative behavior

**Example - WRONG**:
```csharp
// ❌ BAD: Added fallback without user consent
if (sectionBox == null)
{
    _logger("WARNING: No section box, using entire model");
    sectionBox = document.GetBoundingBox(); // FALLBACK - requires consent!
}
```

**Example - CORRECT**:
```csharp
// ✅ GOOD: No fallback - stop execution as user requested
if (sectionBox == null)
{
    _logger("ERROR: Section box is REQUIRED");
    return new List<>(); // Stop - no fallback
}
```

---

## 🚨 MANDATORY: Check Before Editing Code

Before making ANY code change, verify:
1. ✅ Have I discussed the changes with the user and gotten consent?
2. ✅ Am I following OOP principles?
3. ✅ Am I using existing infrastructure instead of creating new?
4. ✅ Am I understanding the architecture before changing it?
5. ✅ Am I fixing root causes, not symptoms?
6. ✅ Am I respecting user requirements?

---

## 1. OOP (Object-Oriented Programming) Principles

### ✅ DO: Use Existing Service Classes
- **FlagManager**: Use `FlagManager.SyncFlagsFromGlobal()` instead of inline flag syncing
- **GlobalIndexService**: Use `GlobalIndexService.GetResolvedGuids()` instead of direct XML reads
- **SafeFileLogger**: Use `SafeFileLogger.SafeAppendText()` instead of `File.AppendAllText()`
- **FlagManager**: Use `FlagManager.ResetFlagsForDeletedSleeves()` instead of manual flag resets

### ❌ DON'T: Create Duplicate Logic
- Don't iterate through categories inline - use service methods
- Don't write flag sync logic inline - use `FlagManager`
- Don't write file operations inline - use `SafeFileLogger`
- Don't create new methods without checking if they exist in service classes

### Example - WRONG:
```csharp
// ❌ BAD: Inline logic, duplicates FlagManager functionality
foreach (var category in categories)
{
    var globalIndex = GlobalIndexService.LoadOrCreate(_document, category);
    foreach (var clashZone in clashZones)
    {
        var entry = globalIndex.Entries.FirstOrDefault(e => e.Id == clashZone.Id.ToString());
        if (entry != null)
        {
            clashZone.IsResolved = entry.IsResolved; // Duplicate logic!
        }
    }
}
```

### Example - CORRECT:
```csharp
// ✅ GOOD: Use existing OOP method
_flagManager.SyncFlagsFromGlobal(clashZones, category);
```

---

## 2. Logging Standards

### ✅ DO: Always Use SafeFileLogger Wrapper
- Use `SafeFileLogger.SafeAppendText(fileName, message)` for ALL file logging
- Use `SafeFileLogger.GetLogFilePath(fileName)` for log paths
- Use `SafeFileLogger.GetLogDirectory()` for log directory
- Respect `DeploymentConfiguration.DeploymentMode` flag

### ❌ DON'T: Bypass SafeFileLogger
- Never use `File.AppendAllText()` directly
- Never use `File.WriteAllText()` directly
- Never bypass deployment mode checks
- Never create new logging infrastructure

### Example - WRONG:
```csharp
// ❌ BAD: Direct file write, bypasses SafeFileLogger
if (!DeploymentConfiguration.DeploymentMode)
{
    File.AppendAllText(logPath, "Log entry\n");
}
```

### Example - CORRECT:
```csharp
// ✅ GOOD: Use SafeFileLogger wrapper
SafeFileLogger.SafeAppendText("refresh.log", "Log entry\n");
```

---

## 3. Flag Management Standards

### ✅ DO: Use FlagManager for All Flag Operations
- Use `FlagManager.SyncFlagsFromGlobal()` to sync flags from Global XML
- Use `FlagManager.ResetFlagsForDeletedSleeves()` to reset flags for deleted sleeves
- Use `FlagManager.UpdateFlagsForPlacement()` to update flags after placement
- Understand flag hierarchy: `IsClusterResolved` takes precedence over `IsResolved`

### ❌ DON'T: Manipulate Flags Directly
- Never set `clashZone.IsResolved = true/false` directly without FlagManager
- Never set `clashZone.IsClusterResolved = true/false` directly without FlagManager
- Never sync flags inline - use FlagManager methods
- Never reset flags inline - use FlagManager methods

### Example - WRONG:
```csharp
// ❌ BAD: Direct flag manipulation, bypasses FlagManager
clashZone.IsResolved = true;
clashZone.SleeveInstanceId = sleeveId;
```

### Example - CORRECT:
```csharp
// ✅ GOOD: Use FlagManager
_flagManager.UpdateFlagsForPlacement(clashZone, sleeveId, isCluster: false, categoryName);
```

---

## 4. Global XML Standards

### ✅ DO: Use GlobalIndexService for All Global XML Operations
- Use `GlobalIndexService.LoadOrCreate()` to load Global XML
- Use `GlobalIndexService.GetResolvedGuids()` for single category
- Use `GlobalIndexService.GetResolvedGuidsForCategories()` for multiple categories
- Use `GlobalIndexService.UpsertFlagsWithIds()` to update flags

### ❌ DON'T: Access Global XML Directly
- Never deserialize Global XML files directly
- Never iterate through categories inline - use service methods
- Never bypass GlobalIndexService for Global XML operations

### Example - WRONG:
```csharp
// ❌ BAD: Direct XML access, bypasses GlobalIndexService
var serializer = new XmlSerializer(typeof(CategoryGlobalIndex));
using (var reader = new StreamReader(xmlPath))
{
    var index = (CategoryGlobalIndex)serializer.Deserialize(reader);
    // Process index...
}
```

### Example - CORRECT:
```csharp
// ✅ GOOD: Use GlobalIndexService
var globalIndex = GlobalIndexService.LoadOrCreate(_document, categoryName);
var entry = globalIndex.Entries.FirstOrDefault(e => e.Id == clashZone.Id.ToString());
```

---

## 5. Performance Standards

### ✅ DO: Avoid Expensive Operations
- Never add expensive Revit API calls (like `GetElement()` in loops) without user consent
- Use existing data structures (flags, caches) instead of API calls
- Check flags/caches before expensive operations
- Use parameter caching to avoid repeated `LookupParameter()` calls

### ❌ DON'T: Add Expensive Operations Without User Consent
- Never add `OpeningDuplicationChecker.IsAnySleeveAtLocation()` without explicit user approval
- Never add expensive loops without optimization
- Never bypass existing performance optimizations

### Example - WRONG:
```csharp
// ❌ BAD: Expensive operation without user consent
if (!OpeningDuplicationChecker.IsAnySleeveAtLocation(_doc, point, tolerance))
{
    // Place sleeve
}
```

### Example - CORRECT:
```csharp
// ✅ GOOD: Use existing flags (no expensive API calls)
if (!clashZone.IsResolved && !clashZone.IsClusterResolved)
{
    // Place sleeve
}
```

---

## 6. Architecture Understanding Standards

### ✅ DO: Understand Architecture Before Changing
- Read existing code to understand data flow
- Understand flag hierarchy: Global XML → Filter XML → Memory
- Understand service responsibilities (FlagManager, GlobalIndexService, SafeFileLogger)
- Check existing OOP methods before creating new ones

### ❌ DON'T: Make Changes Without Understanding
- Never bypass existing infrastructure without understanding why it exists
- Never create duplicate methods without checking existing ones
- Never assume existing code is wrong - investigate first

### Example - WRONG:
```csharp
// ❌ BAD: Created new method without checking if FlagManager already has it
public void SyncFlagsFromGlobal(List<ClashZone> clashZones, string category)
{
    // Duplicate logic - FlagManager already has this!
}
```

### Example - CORRECT:
```csharp
// ✅ GOOD: Use existing FlagManager method
_flagManager.SyncFlagsFromGlobal(clashZones, category);
```

---

## 7. Root Cause Fixing Standards

### ✅ DO: Fix Root Causes, Not Symptoms
- Investigate why a problem exists before fixing
- Use diagnostic logging to understand the issue
- Fix the actual problem, not workarounds
- Use existing debugging tools instead of creating new ones

### ❌ DON'T: Create Quick Fixes That Cause New Problems
- Never bypass existing infrastructure to fix symptoms
- Never add expensive operations to fix symptoms
- Never create workarounds - fix root causes

### Example - WRONG:
```csharp
// ❌ BAD: Bypassing SafeFileLogger to fix missing logs (symptom fix)
File.AppendAllText(logPath, "Log entry\n"); // Creates new problems!
```

### Example - CORRECT:
```csharp
// ✅ GOOD: Enhanced SafeFileLogger error logging to diagnose root cause
catch (Exception ex)
{
    // Enhanced error logging to diagnose why SafeFileLogger fails
    System.Diagnostics.Debug.WriteLine($"[SafeFileLogger] Error: {ex.Message}");
}
```

---

## 8. Code Review Checklist

Before submitting code changes, verify:

### OOP Principles
- [ ] Am I using existing service classes (FlagManager, GlobalIndexService, SafeFileLogger)?
- [ ] Am I creating duplicate logic instead of using existing methods?
- [ ] Am I following single responsibility principle?

### Logging
- [ ] Am I using `SafeFileLogger.SafeAppendText()` instead of direct file writes?
- [ ] Am I respecting `DeploymentConfiguration.DeploymentMode`?
- [ ] Am I bypassing SafeFileLogger wrapper?

### Flag Management
- [ ] Am I using `FlagManager` methods instead of direct flag manipulation?
- [ ] Am I understanding flag hierarchy (`IsClusterResolved` > `IsResolved`)?
- [ ] Am I syncing flags from Global XML properly?

### Performance
- [ ] Am I adding expensive operations without user consent?
- [ ] Am I using existing data structures (flags, caches) instead of API calls?
- [ ] Am I checking flags before expensive operations?

### Architecture
- [ ] Do I understand the existing architecture before changing it?
- [ ] Am I checking for existing methods before creating new ones?
- [ ] Am I following the established data flow?

### Root Cause
- [ ] Am I fixing root causes, not symptoms?
- [ ] Am I investigating before fixing?
- [ ] Am I creating workarounds instead of proper fixes?

---

## 9. Common Anti-Patterns to Avoid

### ❌ Anti-Pattern 1: Bypassing Instead of Fixing
**Problem**: When system doesn't work → bypass it
**Correct**: Investigate why it doesn't work → fix root cause

### ❌ Anti-Pattern 2: Adding Expensive Operations
**Problem**: Add expensive API calls to solve problem quickly
**Correct**: Use existing data structures/flags/caches

### ❌ Anti-Pattern 3: Not Understanding Architecture
**Problem**: Make changes without understanding data flow
**Correct**: Read code, understand architecture, then fix

### ❌ Anti-Pattern 4: Creating Duplicate Methods
**Problem**: Create new methods without checking existing ones
**Correct**: Check existing service classes first, use existing methods

### ❌ Anti-Pattern 5: Direct File Writes
**Problem**: Use `File.AppendAllText()` instead of `SafeFileLogger`
**Correct**: Always use `SafeFileLogger.SafeAppendText()`

---

## 10. Service Class Responsibilities

### FlagManager
- **Responsibility**: All flag operations (sync, reset, update)
- **Methods**: 
  - `SyncFlagsFromGlobal()` - Sync flags from Global XML
  - `ResetFlagsForDeletedSleeves()` - Reset flags for deleted sleeves
  - `UpdateFlagsForPlacement()` - Update flags after placement
- **Never**: Manipulate flags directly, bypass FlagManager

### GlobalIndexService
- **Responsibility**: All Global XML operations
- **Methods**:
  - `LoadOrCreate()` - Load or create Global XML
  - `GetResolvedGuids()` - Get resolved GUIDs for single category
  - `GetResolvedGuidsForCategories()` - Get resolved GUIDs for multiple categories
  - `UpsertFlagsWithIds()` - Update flags in Global XML
- **Never**: Access Global XML files directly, bypass GlobalIndexService

### SafeFileLogger
- **Responsibility**: All file logging operations
- **Methods**:
  - `SafeAppendText()` - Append text to log file
  - `GetLogFilePath()` - Get full log file path
  - `GetLogDirectory()` - Get log directory
- **Never**: Use `File.AppendAllText()` or `File.WriteAllText()` directly

---

## 11. Flag Hierarchy and Flow

### Flag Hierarchy
1. **IsClusterResolved = true**: Highest priority - cluster sleeve exists
2. **IsResolved = true**: Lower priority - individual sleeve exists
3. **Both false**: Unresolved - needs placement

### Data Flow
```
Global XML (source of truth)
    ↓
Filter XML (filter-specific state)
    ↓
Memory (runtime state)
```

### Sync Flow
- **Refresh**: Sync flags FROM Global XML TO Filter XML
- **Placement**: Update flags FROM Memory TO Global XML
- **Filter Load**: Sync flags FROM Global XML TO Memory

---

## 12. When Adding New Categories or Linked Files

### ✅ DO: Check Global XML for ALL Categories
- Load existing clash zones FIRST to get all categories
- Check Global XML for ALL categories (not just selectedMepCategories)
- Use `GlobalIndexService.GetResolvedGuidsForCategories()` for multiple categories
- Sync flags AFTER merging new clash zones

### ❌ DON'T: Only Check Selected Categories
- Never build `__globalsResolved` only from `selectedMepCategories`
- Never skip categories from existing clash zones
- Never assume new categories don't have Global XML entries

---

## 13. Error Handling Standards

### ✅ DO: Use Existing Error Handling
- Use `SafeFileLogger` for error logging (handles exceptions internally)
- Use `FlagManager` error handling (throws exceptions properly)
- Log errors but continue processing when appropriate

### ❌ DON'T: Create New Error Handling
- Never create new try-catch blocks for file operations (SafeFileLogger handles it)
- Never bypass existing error handling
- Never ignore errors silently

---

## 14. Memory and Performance Standards

### ✅ DO: Use Existing Optimizations
- Use parameter caching to avoid repeated `LookupParameter()` calls
- Use `FlagManager` instead of inline flag operations
- Use `GlobalIndexService` instead of direct XML reads
- Respect existing memory management (MemoryManager, MemoryProfiler)

### ❌ DON'T: Create New Optimizations That Break Existing Ones
- Never bypass existing caching mechanisms
- Never add expensive operations that break performance
- Never ignore existing memory management

---

## 15. Documentation Standards

### ✅ DO: Document Changes Properly
- Add comments explaining WHY changes were made
- Reference existing OOP methods in comments
- Document root cause fixes
- Update this document if new patterns emerge

### ❌ DON'T: Skip Documentation
- Never make changes without comments
- Never skip explaining root cause fixes
- Never forget to update standards if patterns change

---

## 16. Testing Standards

### ✅ DO: Verify Changes Don't Break Existing Functionality
- Test that existing OOP methods still work
- Test that flags are synced correctly
- Test that logging works properly
- Test that no new performance issues are introduced

### ❌ DON'T: Assume Changes Work
- Never skip testing after changes
- Never assume existing functionality still works
- Never ignore user feedback about issues

---

## Summary: Mandatory Checklist Before Code Changes

1. ✅ **GET USER CONSENT** - ALWAYS discuss and get explicit approval before editing code
2. ✅ **NO FALLBACKS WITHOUT CONSENT** - Never add fallback logic without asking user first - STRICT INSTRUCTION
3. ✅ **STOP ON ERRORS** - If user says no fallback, return error/empty result - don't continue execution
4. ✅ **DON'T EDIT WITHOUT CONSENT** - Especially TestProfileManagementCommand or other entry points - ask first
5. ✅ **Read existing code** - Understand architecture before changing
6. ✅ **Check existing service classes** - Use existing OOP methods
7. ✅ **Use SafeFileLogger** - Never bypass logging wrapper
8. ✅ **Use FlagManager** - Never manipulate flags directly
9. ✅ **Use GlobalIndexService** - Never access Global XML directly
10. ✅ **Fix root causes** - Never create workarounds
11. ✅ **Respect user requirements** - Never add expensive operations without consent
12. ✅ **Follow OOP principles** - Single responsibility, use existing infrastructure
13. ✅ **Check for duplicates** - Don't create duplicate methods
14. ✅ **Test thoroughly** - Verify changes don't break existing functionality

---

**Remember**: If you're about to bypass existing infrastructure or create duplicate logic, STOP and check this document first. There's almost always an existing OOP method that does what you need.

---

*This document should be reviewed and updated when new patterns emerge or new service classes are created.*

