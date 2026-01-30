# 🔧 3 PREFIX FIXES FOR PARAMETER SERVICE (CORRECTED - USER-DEFINED ONLY)

## 📋 REQUIREMENTS (CLARIFIED)

### Fix #1: Cluster Prefix Should Respect Service Type
**Issue:** When a cluster has zones with the same service type (e.g., 2 ducts with "Chilled Water Supply"), the cluster should use the service type prefix from user settings.

**Rule:** If all zones in cluster have same service type → use **user-defined** service type prefix (same as individual sleeve)

### Fix #2: Multi-Link Cluster/Combined Gets "MEP" Prefix  
**Issue:** Clusters or combined sleeves containing zones from different linked files should always use "MEP" prefix.

**Rule:** If zones are from different links → use **"MEP"** prefix (this is hardcoded)

### Fix #3: Combined Sleeve with "Chilled" - Exception to Rule #2
**Issue:** Combined sleeves are normally "MEP", BUT if zones are from **single link** AND contain "chilled water" system type, use the **user-defined system type prefix** instead.

**Rule:** 
- Normal combined sleeve → "MEP" (default)
- **EXCEPTION:** Single link + "chilled water" system → use **user-defined prefix** from system type overrides

---

## 📍 FILE TO MODIFY

**File:** `C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Services\MarkParameterService.cs`

**Method:** `ResolveDisciplinePrefixForElement` (around line 1430)

---

## ✅ THE CORRECTED FIX

### Step 1: Add New Helper Method (Add after line 1430)

```csharp
/// <summary>
/// ✅ FIX #1, #2, #3: Enhanced prefix resolution for clusters and combined sleeves
/// Checks for:
/// - Multi-link zones → "MEP" prefix (hardcoded)
/// - Same service type across all zones → user-defined service type prefix
/// - Combined sleeve with "chilled" in single link → user-defined system type prefix (EXCEPTION)
/// </summary>
private string ResolveDisciplinePrefixForClusterOrCombined(
    FamilyInstance sleeve,
    string category,
    string defaultPrefix,
    MarkPrefixSettings markPrefixes,
    Document doc,
    SleeveDbContext sharedContext)
{
    if (sleeve == null || sharedContext == null || markPrefixes == null)
        return defaultPrefix;
    
    try
    {
        int sleeveId = sleeve.Id.IntegerValue;
        List<ClashZone> zones = new List<ClashZone>();
        
        // Determine if this is a cluster or combined sleeve
        bool isCluster = false;
        bool isCombined = IsCombinedSleeve(sleeve);
        
        if (isCombined)
        {
            // ✅ FIX #3: For combined sleeves, get all zones from CombinedSleeves table
            using (var cmd = sharedContext.Connection.CreateCommand())
            {
                cmd.CommandText = @"
                    SELECT cz.* FROM ClashZones cz
                    INNER JOIN CombinedSleeveZones csz ON csz.ClashZoneGuid = cz.Id
                    WHERE csz.CombinedInstanceId = @CombinedId";
                cmd.Parameters.AddWithValue("@CombinedId", sleeveId);
                
                using (var reader = cmd.ExecuteReader())
                {
                    var repository = new ClashZoneRepository(sharedContext);
                    while (reader.Read())
                    {
                        // Use reflection to call private MapClashZone method
                        var mapMethod = repository.GetType().GetMethod("MapClashZone", 
                            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                        if (mapMethod != null)
                        {
                            var zone = mapMethod.Invoke(repository, new object[] { reader }) as ClashZone;
                            if (zone != null)
                                zones.Add(zone);
                        }
                    }
                }
            }
        }
        else
        {
            // ✅ FIX #1: For clusters, get all zones from ClashZones table by ClusterSleeveInstanceId
            var repository = new ClashZoneRepository(sharedContext);
            var allZones = repository.GetClashZonesByCategory(category);
            zones = allZones?.Where(z => z.ClusterSleeveInstanceId == sleeveId && z.IsClusterResolved).ToList() 
                    ?? new List<ClashZone>();
            isCluster = zones.Count > 0;
        }
        
        if (zones.Count == 0)
            return defaultPrefix; // No zones found, use default
        
        // ✅ FIX #2: Check if zones are from different links
        var distinctLinks = zones
            .Select(z => z.HostLinkInstanceId ?? 0)
            .Distinct()
            .ToList();
        
        bool isMultiLink = distinctLinks.Count > 1;
        bool isSingleLink = distinctLinks.Count == 1;
        
        if (isMultiLink)
        {
            // Multiple linked files → use "MEP" prefix (HARDCODED)
            string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            if (!DeploymentConfiguration.DeploymentMode)
            {
                File.AppendAllText(mepmarkLogPath,
                    $"[PREFIX-FIX-2] ✅ {(isCombined ? "Combined" : "Cluster")} sleeve {sleeveId} has zones from {distinctLinks.Count} different links → using 'MEP' prefix (hardcoded)\n");
            }
            return "MEP";
        }
        
        // ✅ FIX #3: EXCEPTION - Combined sleeve with single link + "chilled" system type
        // This overrides the normal "MEP" default for combined sleeves
        if (isCombined && isSingleLink)
        {
            // Check if any zone contains "chilled water" system type
            var chilledZones = zones.Where(z => 
            {
                var systemType = GetClashParameterValue(z, "System Type", "MEP System Type", "System Classification");
                var serviceName = GetClashParameterValue(z, "System Name", "Service Name");
                return (systemType?.Contains("Chill", StringComparison.OrdinalIgnoreCase) ?? false) ||
                       (systemType?.Contains("CHW", StringComparison.OrdinalIgnoreCase) ?? false) ||
                       (serviceName?.Contains("Chill", StringComparison.OrdinalIgnoreCase) ?? false);
            }).ToList();
            
            if (chilledZones.Count > 0)
            {
                // Get the system type from the first chilled zone
                var systemType = GetClashParameterValue(chilledZones[0], "System Type", "MEP System Type", "System Classification");
                
                // ✅ USE USER-DEFINED PREFIX from system type overrides
                var resolvedPrefix = markPrefixes.GetPrefixForElement(category, systemType, null);
                
                if (!string.IsNullOrWhiteSpace(resolvedPrefix))
                {
                    string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath,
                            $"[PREFIX-FIX-3] ✅ EXCEPTION: Combined sleeve {sleeveId} is single-link + contains 'chilled water' system type '{systemType}' → using USER-DEFINED prefix '{resolvedPrefix}'\n");
                    }
                    return resolvedPrefix;
                }
                else
                {
                    // No user-defined override found, log warning and continue to default logic
                    string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        File.AppendAllText(mepmarkLogPath,
                            $"[PREFIX-FIX-3] ⚠️ Combined sleeve {sleeveId} has 'chilled water' system type '{systemType}' but NO user-defined override found → will use default MEP\n");
                    }
                }
            }
        }
        
        // ✅ FIX #1: Check if all zones have the same service type (applies to CLUSTERS and NON-CHILLED COMBINED)
        var systemTypes = zones
            .Select(z => GetClashParameterValue(z, "System Type", "MEP System Type", "System Classification"))
            .Where(st => !string.IsNullOrWhiteSpace(st))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        
        if (systemTypes.Count == 1)
        {
            // All zones have same system type → use USER-DEFINED service type prefix
            string systemType = systemTypes[0];
            
            // ✅ USE USER-DEFINED PREFIX from system type overrides
            var resolvedPrefix = markPrefixes.GetPrefixForElement(category, systemType, null);
            var defaultCategoryPrefix = markPrefixes.GetDisciplinePrefix(category);
            
            if (!string.IsNullOrWhiteSpace(resolvedPrefix) && 
                !resolvedPrefix.Equals(defaultCategoryPrefix, StringComparison.OrdinalIgnoreCase))
            {
                // User-defined service type override exists and is different from default category prefix
                string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    File.AppendAllText(mepmarkLogPath,
                        $"[PREFIX-FIX-1] ✅ {(isCombined ? "Combined" : "Cluster")} sleeve {sleeveId} has all zones with same system type '{systemType}' → using USER-DEFINED prefix '{resolvedPrefix}'\n");
                }
                return resolvedPrefix;
            }
        }
        
        // ✅ DEFAULT FOR COMBINED SLEEVES: Use "MEP" if no exception applies
        if (isCombined)
        {
            string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            if (!DeploymentConfiguration.DeploymentMode)
            {
                File.AppendAllText(mepmarkLogPath,
                    $"[PREFIX-DEFAULT] Combined sleeve {sleeveId} → using default 'MEP' prefix (no exception applied)\n");
            }
            return "MEP";
        }
    }
    catch (Exception ex)
    {
        if (!DeploymentConfiguration.DeploymentMode)
        {
            string mepmarkLogPath = SafeFileLogger.GetLogFilePath("mepmark_debug.log");
            File.AppendAllText(mepmarkLogPath,
                $"[PREFIX-FIX] ❌ Error resolving cluster/combined prefix for sleeve {sleeve.Id}: {ex.Message}\n");
        }
    }
    
    // Default: use standard resolution for clusters or fallback
    return defaultPrefix;
}
```

### Step 2: Modify ResolveDisciplinePrefixForElement (around line 283)

Find this code in the main marking loop:

```csharp
// ✅ ENHANCED: Resolve element-specific prefix (e.g., System Type override)
var elementPrefix = ResolveDisciplinePrefixForElement(category, disciplinePrefix, markPrefixes, clashZone);
if (!string.IsNullOrWhiteSpace(elementPrefix))
{
    candidatePrefixes.Add(elementPrefix);
}
```

**Replace with:**

```csharp
// ✅ ENHANCED: Resolve element-specific prefix (e.g., System Type override)
// ✅ FIX #1, #2, #3: Check if this is a cluster or combined sleeve first
var elementPrefix = disciplinePrefix;

// Check if sleeve is cluster or combined
bool isCluster = (clashZone != null && clashZone.IsClusterResolved && clashZone.ClusterSleeveInstanceId == sleeveId);
bool isCombined = IsCombinedSleeve(sleeve);

if (isCluster || isCombined)
{
    // ✅ Apply cluster/combined-specific prefix logic (user-defined prefixes only, no hardcoding)
    elementPrefix = ResolveDisciplinePrefixForClusterOrCombined(
        sleeve, category, disciplinePrefix, markPrefixes, doc, sharedContext);
}
else
{
    // ✅ Individual sleeve: use standard resolution
    elementPrefix = ResolveDisciplinePrefixForElement(category, disciplinePrefix, markPrefixes, clashZone);
}

if (!string.IsNullOrWhiteSpace(elementPrefix))
{
    candidatePrefixes.Add(elementPrefix);
}
```

---

## 🧪 CORRECTED TESTING SCENARIOS

### Test #1: Cluster with Same Service Type
**Setup:**
- Create cluster with 2 ducts: both "Chilled Water Supply"
- **UI Setting:** System Type override: "Chilled Water Supply" → "CHW" (user-defined)

**Expected:**
- Individual sleeves: "CHW001", "CHW002"
- Cluster sleeve: "CHW003" ✓ (uses **user-defined** "CHW", not "DCT")

**Verification Log:**
```
[PREFIX-FIX-1] ✅ Cluster sleeve 123456 has all zones with same system type 'Chilled Water Supply' → using USER-DEFINED prefix 'CHW'
```

### Test #2: Multi-Link Cluster
**Setup:**
- Create cluster with 2 ducts: 1 from Link A, 1 from Link B
- Both same category (Ducts)

**Expected:**
- Cluster sleeve: "MEP001" ✓ (hardcoded "MEP" for multi-link, this is correct)

**Verification Log:**
```
[PREFIX-FIX-2] ✅ Cluster sleeve 123456 has zones from 2 different links → using 'MEP' prefix (hardcoded)
```

### Test #3: Combined Sleeve with "Chilled" (Single Link) - EXCEPTION
**Setup:**
- Create combined sleeve with: 1 duct + 1 pipe (both chilled water)
- **All zones in SINGLE link**
- **UI Setting:** System Type override: "Chilled Water Supply" → "CHW" (user-defined)

**Expected:**
- Combined sleeve: "CHW001" ✓ (uses **user-defined** "CHW", NOT "MEP" - this is the EXCEPTION!)

**Verification Log:**
```
[PREFIX-FIX-3] ✅ EXCEPTION: Combined sleeve 123456 is single-link + contains 'chilled water' system type 'Chilled Water Supply' → using USER-DEFINED prefix 'CHW'
```

### Test #4: Combined Sleeve with "Chilled" (Single Link) - NO USER SETTING
**Setup:**
- Create combined sleeve with: 1 duct + 1 pipe (both chilled water)
- **All zones in SINGLE link**
- **UI Setting:** NO system type override configured

**Expected:**
- Combined sleeve: "MEP001" ✓ (no user override → falls back to "MEP")

**Verification Log:**
```
[PREFIX-FIX-3] ⚠️ Combined sleeve 123456 has 'chilled water' system type 'Chilled Water Supply' but NO user-defined override found → will use default MEP
[PREFIX-DEFAULT] Combined sleeve 123456 → using default 'MEP' prefix (no exception applied)
```

### Test #5: Combined Sleeve Mixed Services (Same Link, NO "Chilled")
**Setup:**
- Create combined sleeve with: 1 duct (supply air) + 1 pipe (hot water)
- All zones in same link
- NO "chilled" keyword anywhere

**Expected:**
- Combined sleeve: "MEP001" ✓ (default for combined)

**Verification Log:**
```
[PREFIX-DEFAULT] Combined sleeve 123456 → using default 'MEP' prefix (no exception applied)
```

### Test #6: Combined Sleeve from Different Links (with "Chilled")
**Setup:**
- Create combined sleeve with: 1 chilled water duct from Link A + 1 pipe from Link B
- **Zones from DIFFERENT links**

**Expected:**
- Combined sleeve: "MEP001" ✓ (multi-link overrides everything, hardcoded "MEP")

**Verification Log:**
```
[PREFIX-FIX-2] ✅ Combined sleeve 123456 has zones from 2 different links → using 'MEP' prefix (hardcoded)
```

---

## 📊 CORRECTED PRIORITY ORDER

The fix applies checks in this order:

1. **FIX #2** (Multi-Link): **HIGHEST PRIORITY**
   - Zones from different links → Use **"MEP"** (hardcoded)
   - This overrides everything else!

2. **FIX #3** (Combined + "Chilled" + Single Link): **EXCEPTION CASE**
   - Combined sleeve + single link + "chilled water" → Use **user-defined** system type prefix
   - This is the ONLY exception to combined sleeves using "MEP"

3. **FIX #1** (Same Service Type): **THIRD PRIORITY**
   - All zones same system type → Use **user-defined** service type prefix
   - Applies to both clusters AND combined sleeves (if no exception)

4. **DEFAULT (Combined)**: **FALLBACK**
   - Combined sleeve with no exception → Use **"MEP"** (hardcoded default)

5. **DEFAULT (Other)**: **FINAL FALLBACK**
   - Use standard category prefix resolution

---

## 🔑 KEY DIFFERENCES FROM PREVIOUS VERSION

### ❌ WRONG (Previous):
- Hardcoded "CHW" prefix → **REMOVED**
- Hardcoded "MEP" prefix for all combined → **CHANGED**

### ✅ CORRECT (New):
- **All prefixes are user-defined** (from UI system type overrides)
- **Only "MEP" is hardcoded** (for multi-link and default combined)
- **Exception added:** Single-link combined with "chilled" can use user-defined prefix

---

## 🔍 VERIFICATION CHECKLIST

After applying fix, verify these scenarios work:

- [ ] Cluster with same service type uses **user-defined** service type prefix (Fix #1)
- [ ] Cluster with zones from different links uses **"MEP"** (hardcoded) (Fix #2)
- [ ] Combined sleeve from different links uses **"MEP"** (hardcoded) (Fix #2 - overrides everything)
- [ ] Combined sleeve with single link + "chilled" uses **user-defined** system type prefix (Fix #3 - EXCEPTION)
- [ ] Combined sleeve with single link + "chilled" but NO user setting → falls back to **"MEP"** (Fix #3 fallback)
- [ ] Combined sleeve with single link + NO "chilled" uses **"MEP"** (default for combined)
- [ ] Individual sleeves still work correctly (no regression)
- [ ] Logs show which fix was applied and which prefix source (user-defined vs hardcoded)

---

## 🚨 CRITICAL CLARIFICATIONS

### What Gets Hardcoded:
- ✅ **"MEP"** for multi-link sleeves (Fix #2)
- ✅ **"MEP"** for combined sleeves by default

### What's User-Defined:
- ✅ **All system type prefixes** (e.g., "CHW", "HW", "SA", etc.)
- ✅ **All service type prefixes** (e.g., "PLU", "DCT", "ELE", etc.)
- ✅ **Exception case:** Combined + single link + "chilled" → uses user-defined prefix

### The Exception Logic:
```
Combined sleeve logic:
├─ Multiple links? → "MEP" (hardcoded)
├─ Single link + "chilled"? → User-defined system type prefix (EXCEPTION)
└─ Otherwise → "MEP" (hardcoded default)
```

---

## 📝 SUMMARY

This **CORRECTED** fix ensures:
- ✅ **No hardcoded prefixes** except "MEP" for specific cases
- ✅ **All service/system type prefixes come from user UI settings**
- ✅ **Exception added:** Single-link combined with "chilled" can use user prefix
- ✅ **Multi-link always uses "MEP"** (this is correct and expected)
- ✅ **Clear logging** showing user-defined vs hardcoded prefix source
- ✅ **Fallback to "MEP"** if user hasn't configured system type override

The key insight: **Combined sleeves default to "MEP", with ONE exception - single-link + chilled water gets user-defined prefix if configured!**
