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

## 🚨 ABSOLUTE: NEVER BREAK PRINCIPLES, DOCTRINES, OR ARCHITECTURE

**MANDATORY RULE**: You MUST NEVER compromise, bypass, or modify core principles, doctrines, or established code architecture on your own initiative.

**IMMUTABLE PRINCIPLES** (DO NOT MODIFY):
1. ✅ **3-Point Validation** - MEP+Host+Point validation is CORE and cannot be compromised
2. ✅ **Architectural Patterns** - Service layer separation, OOP structure, dependency injection patterns
3. ✅ **Data Flow Principles** - Database-first, XML fallback, flag hierarchy
4. ✅ **Validation Rules** - Established validation logic, duplicate detection methods
5. ✅ **Core Algorithms** - Intersection detection logic, placement strategies, sizing calculations

**MANDATORY PROCESS**:
1. ✅ **READ CODING STANDARDS FIRST** - Always read this document before making ANY changes
2. ✅ **IDENTIFY PRINCIPLES** - Determine which principles/doctrines apply to the area you're working on
3. ✅ **DISCUSS WITH USER** - If a fix seems to require breaking a principle, DISCUSS with user first
4. ✅ **PROPOSE ALTERNATIVES** - If principle conflict exists, propose alternative solutions that preserve principles
5. ✅ **GET EXPLICIT APPROVAL** - Never modify principles without explicit user approval

**NEVER**:
- ❌ Compromise on core principles to "fix" a problem
- ❌ Bypass validation rules (e.g., skip 3-point validation)
- ❌ Modify architectural patterns without discussion
- ❌ Change established algorithms without user consent
- ❌ Assume "it's just stale data" and change code - discuss data cleanup first
- ❌ Modify doctrines based on "whims" or assumptions

**Examples of PRINCIPLE VIOLATIONS**:
- ❌ **WRONG**: Skipping 3-point validation (MEP+Host+Point) for dampers because intersection points changed
- ❌ **WRONG**: Modifying duplicate detection to use MEP+Host only when point matching fails
- ❌ **WRONG**: Changing architecture to work around stale data instead of cleaning data
- ❌ **WRONG**: Compromising validation principles to handle edge cases

**Examples of CORRECT APPROACH**:
- ✅ **CORRECT**: Identify that stale data is causing duplicates → Suggest deleting DB and refreshing
- ✅ **CORRECT**: Maintain 3-point validation → Fix intersection point calculation properly
- ✅ **CORRECT**: Preserve architecture → Work within established patterns
- ✅ **CORRECT**: Discuss principle conflicts → Propose alternatives that preserve principles

**CRITICAL REMINDER**: 
> "We cannot compromise on 3-point validation. Don't play around with code to your whims and fancy. If this is due to stale data, I can delete the DB and start afresh without discussion. Why would you change principles?"

**Read this section time and again** - Principles, doctrines, and architecture are IMMUTABLE without explicit user approval and discussion.

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

## 🚨 MANDATORY: ALWAYS CREATE BACKUP BEFORE RUNNING PYTHON SCRIPTS

**CRITICAL SAFETY RULE**: Before running ANY Python script, you MUST create a backup of all files/data that the script will modify or access.

**MANDATORY PROCESS**:
1. ✅ **IDENTIFY FILES** - Determine which files, databases, or data the Python script will access or modify
2. ✅ **CREATE BACKUP** - Make a complete backup copy of all identified files/data BEFORE running the script
3. ✅ **VERIFY BACKUP** - Ensure backup is complete and can be restored if needed
4. ✅ **THEN RUN SCRIPT** - Only after backup is confirmed, proceed with running the Python script

**BACKUP REQUIREMENTS**:
- ✅ **Database Files** - Backup SQLite databases (.db files) before any Python script that modifies them
- ✅ **Configuration Files** - Backup config files, XML files, JSON files before Python scripts that read/write them
- ✅ **Data Files** - Backup any data files the script processes or modifies
- ✅ **Project Files** - If script modifies project structure, backup entire project directory

**NEVER**:
- ❌ Run Python scripts without creating backups first
- ❌ Assume "it will be fine" - Python scripts can cause unforeseen havoc
- ❌ Skip backup "because it's just a small change"
- ❌ Run destructive operations without restore capability

**RATIONALE**:
> "Always keep a backup before running any Python script. So even if Python code creates havoc unforeseen, we can restore."

**Examples**:
- ✅ **CORRECT**: Before running a Python script that modifies a SQLite database → Create backup copy of .db file
- ✅ **CORRECT**: Before running a Python script that processes XML files → Backup all XML files first
- ✅ **CORRECT**: Before running a Python script that modifies project structure → Create full project backup
- ❌ **WRONG**: Running Python script immediately without backup → Risk of data loss if script fails

**Recovery Plan**: Always know HOW to restore from backup before running the script.

---

## 🚨 MANDATORY: Check Before Editing Code

Before making ANY code change, verify:
1. ✅ **HAVE I READ THE CODING STANDARDS?** - Especially the "NEVER BREAK PRINCIPLES" section above
2. ✅ **AM I PRESERVING PRINCIPLES?** - Am I maintaining 3-point validation, architectural patterns, core algorithms?
3. ✅ **IS THIS A DATA ISSUE?** - Should I suggest data cleanup (delete DB) instead of changing code principles?
4. ✅ **IF RUNNING PYTHON SCRIPT: HAVE I CREATED BACKUP?** - Always backup all files/data before running Python scripts
5. ✅ Have I discussed the changes with the user and gotten consent?
6. ✅ Am I following OOP principles?
7. ✅ Am I using existing infrastructure instead of creating new?
8. ✅ Am I understanding the architecture before changing it?
9. ✅ Am I fixing root causes, not symptoms?
10. ✅ Am I respecting user requirements?

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
- **❌ NEVER CREATE 2 SERVICES FOR ONE FUNCTION** - This makes troubleshooting very hard. Always extend existing services instead of creating duplicates (e.g., use `RefreshPathDeterminer` for both refresh and placement path determination, don't create separate `SleevePlacementPathDeterminer`)

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

### 🚨 MANDATORY: NEVER COMPROMISE EFFICIENCY
**CRITICAL PRINCIPLE**: Do NOT violate the principle of compromising efficiency under ANY circumstances.
- ❌ **NEVER** add expensive operations that slow down execution
- ❌ **NEVER** create duplicate services that add overhead
- ❌ **NEVER** bypass existing optimizations for "convenience"
- ❌ **NEVER** add unnecessary checks or loops that impact performance
- ✅ **ALWAYS** prioritize efficiency over "defensive" coding
- ✅ **ALWAYS** use existing optimized methods instead of creating new ones
- ✅ **ALWAYS** check performance impact before adding new code paths

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

### ❌ Anti-Pattern 5: Creating 2 Services for One Function
**Problem**: Create duplicate services (e.g., `SleevePlacementPathDeterminer` when `RefreshPathDeterminer` exists)
**Correct**: Extend existing services - makes troubleshooting much easier, avoids code duplication

### ❌ Anti-Pattern 6: Compromising Efficiency
**Problem**: Add "defensive" code or duplicate services that slow down execution
**Correct**: Never compromise efficiency - prioritize performance over convenience

### ❌ Anti-Pattern 7: Direct File Writes
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

1. ✅ **NEVER BREAK PRINCIPLES** - Read coding standards first - Principles, doctrines, and architecture are IMMUTABLE without explicit user approval
2. ✅ **CREATE BACKUP BEFORE PYTHON SCRIPTS** - Always backup all files/data before running ANY Python script - prevents unforeseen havoc
3. ✅ **GET USER CONSENT** - ALWAYS discuss and get explicit approval before editing code
4. ✅ **NO FALLBACKS WITHOUT CONSENT** - Never add fallback logic without asking user first - STRICT INSTRUCTION
5. ✅ **STOP ON ERRORS** - If user says no fallback, return error/empty result - don't continue execution
6. ✅ **DON'T EDIT WITHOUT CONSENT** - Especially TestProfileManagementCommand or other entry points - ask first
7. ✅ **PRESERVE PRINCIPLES** - Never compromise on core principles (3-point validation, architectural patterns) to "fix" issues
8. ✅ **DISCUSS DATA CLEANUP** - If issue is stale data, suggest deleting DB - don't change code principles
9. ✅ **Read existing code** - Understand architecture before changing
10. ✅ **Check existing service classes** - Use existing OOP methods
11. ✅ **Use SafeFileLogger** - Never bypass logging wrapper
12. ✅ **Use FlagManager** - Never manipulate flags directly
13. ✅ **Use GlobalIndexService** - Never access Global XML directly
14. ✅ **Fix root causes** - Never create workarounds
15. ✅ **Respect user requirements** - Never add expensive operations without consent
16. ✅ **Follow OOP principles** - Single responsibility, use existing infrastructure
17. ✅ **Check for duplicates** - Don't create duplicate methods or services
18. ✅ **NEVER CREATE 2 SERVICES FOR ONE FUNCTION** - Makes troubleshooting very hard - extend existing services instead
19. ✅ **NEVER COMPROMISE EFFICIENCY** - Do not violate efficiency principles under any circumstances
20. ✅ **Test thoroughly** - Verify changes don't break existing functionality

---

**Remember**: If you're about to bypass existing infrastructure or create duplicate logic, STOP and check this document first. There's almost always an existing OOP method that does what you need.

---

*This document should be reviewed and updated when new patterns emerge or new service classes are created.*

