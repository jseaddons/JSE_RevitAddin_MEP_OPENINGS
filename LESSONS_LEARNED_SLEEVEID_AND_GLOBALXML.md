# Lessons Learned: SleeveInstanceId Reset & Global XML Path Bugs

## Date: Current Session
## Bugs Fixed:
1. **BUG #3**: `SleeveInstanceId` reset to 0 on rerun (losing values from fresh run)
2. **BUG #4**: `pipes_global.xml` created in wrong folder (Default instead of project-specific)

---

## 🎯 Key Lessons Learned

### 1. **ALWAYS Preserve Existing XML Data When Updating**
**Problem:** `UpdateExistingClashZone` in `ClashZoneService.cs` was blindly overwriting `SleeveInstanceId = -1` even when the value was already loaded from XML.

**Lesson:**
- ✅ **CHECK BEFORE OVERWRITING**: Always verify if a value is already set before overwriting it
- ✅ **PRESERVE FROM PERSISTENCE**: Data loaded from XML files should be treated as authoritative until explicitly invalidated
- ✅ **DEFENSIVE CODING**: Add conditional checks: `if (existingValue <= 0)` before setting default values

**Code Pattern:**
```csharp
// ❌ BAD: Blindly overwrites
existingZone.SleeveInstanceId = -1;

// ✅ GOOD: Preserves existing values
if (existingZone.SleeveInstanceId <= 0)
{
    existingZone.SleeveInstanceId = -1; // Only set if not already populated
}
```

---

### 2. **Use Centralized Path Services - Never Hardcode Paths**
**Problem:** `GlobalFlagManager` was using hardcoded `"Default"` folder path instead of project-specific directory.

**Lesson:**
- ✅ **CENTRALIZE PATH LOGIC**: Use `ProjectPathService.GetFiltersDirectory(doc)` instead of hardcoding paths
- ✅ **CONSISTENCY ACROSS SERVICES**: All services (`GlobalIndexService`, `GlobalFlagManager`, etc.) should use the same path resolution mechanism
- ✅ **NO FALLBACK TO "Default"**: Unless explicitly intended for fallback scenarios, never hardcode "Default" as it creates files in wrong locations

**Code Pattern:**
```csharp
// ❌ BAD: Hardcoded path
var filtersDirectory = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
    "JSE_MEP_Openings", 
    "Projects", 
    "Default",  // ❌ WRONG!
    "Filters");

// ✅ GOOD: Use centralized service
string filtersDirectory = ProjectPathService.GetFiltersDirectory(doc);
```

---

### 3. **Cache Keys Must Include All Relevant Context**
**Problem:** `GlobalFlagManager` singleton cache used only `categoryName` as key, causing cross-project contamination.

**Lesson:**
- ✅ **INCLUDE PROJECT CONTEXT IN CACHE KEYS**: When caching instances that are project-specific, include the project path/directory in the cache key
- ✅ **PREVENT CROSS-PROJECT CONTAMINATION**: Different projects need different instances (different XML files)
- ✅ **CACHE KEY PATTERN**: Use `"{categoryName}|{projectPath}"` or similar composite keys

**Code Pattern:**
```csharp
// ❌ BAD: Single project key
string cacheKey = categoryName; // Only category - wrong project!

// ✅ GOOD: Composite key with project context
string cacheKey = doc != null ? $"{categoryName}|{filtersDirectory}" : categoryName;
return _instanceCache.GetOrAdd(cacheKey, key => {
    // Parse composite key and create instance with correct project path
});
```

---

### 4. **When Adding Document Parameter, Update All Callers**
**Problem:** After adding `Document` parameter to `GlobalFlagManager.GetOrCreate()`, callers needed to be updated.

**Lesson:**
- ✅ **FIND ALL CALLERS FIRST**: Use `grep` to find all call sites before making breaking changes
- ✅ **UPDATE ALL CALLERS**: Ensure every caller passes the new parameter (use `replace_all` when possible)
- ✅ **VERIFY SERVICES HAVE DOCUMENT ACCESS**: Check if all calling services have access to `Document` object

**Files Updated:**
- `Services/UniversalClusterService.cs` (had `_doc` field)
- `Services/UniversalSleevePlacerService.cs` (had `_doc` field)
- `Services/ClashZoneService.cs` (had `document` parameter)

---

### 5. **Distinguish Between "Not Set" and "Explicitly Zero"**
**Problem:** The check `SleeveInstanceId <= 0` correctly handles both `-1` (not set) and `0` (invalid/explicitly zero).

**Lesson:**
- ✅ **UNDERSTAND DEFAULT VALUES**: Know what values represent "not set" (`-1`) vs "invalid" (`0`) vs "valid" (`> 0`)
- ✅ **PRESERVE SEMANTICS**: When preserving values, respect the meaning of each numeric value
- ✅ **DOCUMENT VALUE SEMANTICS**: Add comments explaining what each value means

**Value Semantics:**
- `SleeveInstanceId = -1`: Not set yet (default, will be populated during placement)
- `SleeveInstanceId = 0`: Invalid or explicitly cleared (rare, usually indicates error state)
- `SleeveInstanceId > 0`: Valid Revit element ID (loaded from XML or just placed)

---

## 🔍 Debugging Process Lessons

### 1. **Identify the Data Flow**
- Trace where data comes from (XML → in-memory → operations → back to XML)
- Identify all places where values can be overwritten
- Check both "set" operations and "reset" operations

### 2. **Check for Hardcoded Paths**
- Search for hardcoded strings like `"Default"`, `"Filters"`, project paths
- Compare with how other similar services handle paths
- Look for inconsistency between services

### 3. **Verify Cache Scope**
- Check if singleton caches are correctly scoped (per-project vs global)
- Verify cache keys include all relevant context
- Test with multiple projects to ensure no cross-contamination

---

## 📋 Prevention Checklist

Before making changes that modify persistent data:

- [ ] **Preserve existing values**: Check if value is already set before overwriting
- [ ] **Use centralized services**: Never hardcode paths, use `ProjectPathService`
- [ ] **Include context in cache keys**: Project-specific caches must include project path
- [ ] **Update all callers**: Find and update all callers when changing method signatures
- [ ] **Understand value semantics**: Know what each numeric value means (default, invalid, valid)
- [ ] **Test with multiple projects**: Verify no cross-project contamination
- [ ] **Verify XML persistence**: Ensure values are correctly saved to and loaded from XML

---

## 🚨 Red Flags to Watch For

1. **Hardcoded paths** with `"Default"` or project names
2. **Singleton caches** without project context in keys
3. **Update methods** that don't check existing values before overwriting
4. **Method signature changes** without updating all callers
5. **Services** that create files in different locations than other services

---

## ✅ Success Criteria

After fixes:
- ✅ `SleeveInstanceId` values are preserved on rerun (loaded from XML)
- ✅ `*_global.xml` files are created in project-specific folders
- ✅ No cross-project contamination in singleton caches
- ✅ All services use `ProjectPathService` consistently
- ✅ All callers updated to pass `Document` parameter

---

## 📝 Related Code Files

- `Services/ClashZoneService.cs` - `UpdateExistingClashZone` method (Bug #3 fix)
- `Services/GlobalFlagManager.cs` - Constructor and `GetOrCreate` method (Bug #4 fix)
- `Services/ProjectPathService.cs` - Centralized path resolution
- `Services/GlobalIndexService.cs` - Reference implementation (already uses `ProjectPathService`)

---

**Remember:** When working with persistent data and singleton services, always consider:
1. **Where the data comes from** (XML, memory, user input)
2. **What preserves it** (defensive checks, proper caching)
3. **What can overwrite it** (update methods, cache misses, incorrect paths)

