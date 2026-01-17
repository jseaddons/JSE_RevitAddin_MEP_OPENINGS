# 🎯 COMPLETE DEBUGGING GUIDE: Why All 6 Sleeves Are Clustered Together

**Date:** January 16, 2026  
**Issue:** Expected 4 sleeves clustered + 2 individual, but all 6 are marked as clustered  
**Goal:** Identify which component is causing this behavior

---

## THE DECISION TREE

```
Did you intend to place 6 sleeves that should form:
├─ 1 cluster of 4 sleeves (close together)
├─ 1 individual sleeve 5 (far away)
└─ 1 individual sleeve 6 (far away)

But instead all 6 are marked as clustered?

Then follow this debugging guide →
```

---

## STEP 1: Check the ClusterSleeves_v2 Table (Database)

**This is the SOURCE OF TRUTH.**

### The Question:
How many cluster records exist, and what do they contain?

### SQL Query:
```sql
SELECT 
    ClusterGUID,
    Status,
    ConstituentZoneGuids,
    (SELECT COUNT(*) FROM (
        SELECT value FROM json_each('[' || 
        REPLACE(REPLACE(ConstituentZoneGuids, ', ', '","'), ',', '","') ||
        ']')
    )) as constituent_count
FROM ClusterSleeves_v2
ORDER BY CalculatedAt DESC
LIMIT 10;
```

### What You'll See:

**CASE A: One cluster with all 6 sleeves**
```
ClusterGUID: xxx
Status: Pending (or Placed)
ConstituentZoneGuids: guid1,guid2,guid3,guid4,guid5,guid6
constituent_count: 6  ← ALL 6 IN ONE!
```
→ **Problem is in CLUSTERING ALGORITHM** (FormClusters())
→ Go to **SECTION A** below

**CASE B: Three clusters (4, 1, 1)**
```
ClusterGUID: aaa
ConstituentZoneGuids: guid1,guid2,guid3,guid4
constituent_count: 4  ← Cluster of 4

ClusterGUID: bbb
ConstituentZoneGuids: guid5
constituent_count: 1  ← Individual

ClusterGUID: ccc
ConstituentZoneGuids: guid6
constituent_count: 1  ← Individual
```
→ **Problem is in PERSISTENCE/FLAGS** (BatchUpdateFlags() or SaveToClusterSleevesLegacy())
→ Go to **SECTION B** below

---

## SECTION A: Clustering Algorithm Issue (All 6 in One Cluster)

**File:** `ClusterAlgorithmService.cs`

### Root Cause:
The `FormClusters()` method uses BFS and is grouping sleeves too aggressively.

### Why:
- Tolerance distance too large, OR
- Host ID validation too weak, OR
- Proximity checker returning true when it shouldn't

### Diagnosis Code:

Add this to `FormClusters()` method:

```csharp
public Dictionary<SleeveGroupKey, List<List<dynamic>>> FormClusters(...)
{
    // ✅ ADD AT START:
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"\n╔════════════════════════════════════════════════╗\n");
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"║ FormClusters START                             ║\n");
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"║ Tolerance: {toleranceDist * 304.8,8:F1}mm                        ║\n");
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"╚════════════════════════════════════════════════╝\n");
    
    // ... existing code ...
    
    // ✅ ADD AFTER clusters.Add(cluster):
    var clusterSummary = string.Join(",", cluster.Select(c => 
        ((c?.ClashZone as ClashZone)?.Id.ToString() ?? "null").Substring(0, 8)));
    
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"[{DateTime.Now:HH:mm:ss}] Cluster #{clusters.Count}: Size={cluster.Count} | GUIDs=[{clusterSummary}...]\n");
    
    // ✅ ADD BEFORE RETURN:
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"╔════════════════════════════════════════════════╗\n");
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"║ FormClusters COMPLETE                          ║\n");
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"║ Total Groups: {result.Count,-3} Total Clusters: {result.Sum(r => r.Value.Count),-3}           ║\n");
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"║ Cluster Sizes: {string.Join(",", result.SelectMany(r => r.Value).Select(c => c.Count))}                         ║\n");
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"╚════════════════════════════════════════════════╝\n");
}
```

### Root Cause Solutions:

#### Solution A1: Reduce Tolerance Distance
```csharp
// Check what toleranceDist is passed
SafeFileLogger.SafeAppendText("cluster_debug.log",
    $"Tolerance: {toleranceDist * 304.8:F1}mm (feet={toleranceDist:F3})\n");

// If > 2 feet (0.61m), it's too large!
// Typical recommendation: 1-6 inches (25-150mm = 0.082-0.49 feet)
if (toleranceDist > 0.5)  // More than 6 inches
{
    SafeFileLogger.SafeAppendText("placement_errors.log",
        $"⚠️ WARNING: Tolerance {toleranceDist * 304.8:F1}mm is very large - may cause over-clustering!\n");
}
```

#### Solution A2: Strengthen Host Validation
```csharp
// IN ShouldClusterSleeves():

// ❌ BEFORE:
if (host1 > 0 && host2 > 0)  // Only if BOTH valid
{
    if (host1 != host2) return false;
}

// ✅ AFTER: If either is valid, they MUST match
if ((host1 > 0 || host2 > 0) && host1 != host2)
{
    SafeFileLogger.SafeAppendText("cluster_debug.log",
        $"❌ REJECT: Different hosts (host1={host1}, host2={host2})\n");
    return false;
}
```

### How to Fix:

**Option 1: Reduce the tolerance distance (RECOMMENDED)**
- Find where `FormClusters()` is called
- Check the `toleranceDist` parameter being passed
- If > 1 foot, reduce it to 6 inches (0.5 feet)

**Option 2: Tighten host validation**
- Apply Solution A2 above
- Re-test clustering

**Option 3: Add clustering strength parameter**
```csharp
// New parameter: How many proximities must sleeves share to cluster?
// Currently: 1 proximity match = cluster
// Could require: 2+ proximity matches = cluster (stricter)

if (proximityScore < 2)  // Require at least 2 proximity matches
{
    return false;  // Don't cluster
}
```

---

## SECTION B: Persistence/Flag Setting Issue (Clusters Calculated Correctly, But All Marked)

**Files:**
- `BatchClusterPlacementService.cs` - Updates flags
- `ClusterSleeveRepository.cs` - Saves to database

### Root Cause:
ClusterSleeves_v2 is correct (3 clusters of 4, 1, 1) but:
1. BatchUpdateFlags() is marking ALL constituent sleeves as clustered
2. SaveToClusterSleevesLegacy() is failing silently
3. ClusterSleeves table remains empty

### Diagnosis:

Check `batch_v2.log`:

```
Look for these messages:
[HH:mm:ss] PerformSwapDeletion START
[HH:mm:ss]   ConstituentZoneGuids: guid1,guid2,guid3,guid4
[HH:mm:ss]   Found 4 zones in database ← Should match cluster size
[HH:mm:ss]   Updated 4 zones with cluster data ← Should match
[HH:mm:ss] 💾 SAVED to ClusterSleeves (legacy): ClusterInstanceId=xxx ← If missing = fail!
```

### How to Fix:

#### Fix B1: Verify SaveToClusterSleevesLegacy() is Not Silently Failing
```csharp
// Add at START of SaveToClusterSleevesLegacy():
SafeFileLogger.SafeAppendText("batch_v2.log",
    $"[{DateTime.Now:HH:mm:ss}] 🔍 SaveToClusterSleevesLegacy START\n" +
    $"    ClusterGUID: {cluster.ClusterGUID}\n" +
    $"    ClusterInstanceId: {clusterInstanceId}\n");

// Change SILENT return to EXCEPTION:
if (comboId <= 0 || filterId <= 0)
{
    SafeFileLogger.SafeAppendText("placement_errors.log",
        $"[{DateTime.Now:HH:mm:ss}] ❌ CRITICAL: Cannot save - Missing ComboId ({comboId}) or FilterId ({filterId})\n");
    
    // ✅ THROW so caller knows:
    throw new InvalidOperationException(
        $"Cannot save to ClusterSleeves: ComboId={comboId}, FilterId={filterId}");
}
```

#### Fix B2: Verify ClusterSleeves Table Insert
```csharp
// ADD AFTER repo.SaveClusterSleeve():
SafeFileLogger.SafeAppendText("batch_v2.log",
    $"[{DateTime.Now:HH:mm:ss}] ✅ SaveClusterSleeve completed\n");

// Verify it actually inserted:
using (var cmd = dbContext.Connection.CreateCommand())
{
    cmd.CommandText = "SELECT COUNT(*) FROM ClusterSleeves WHERE ClusterInstanceId = @id";
    cmd.Parameters.AddWithValue("@id", clusterInstanceId);
    int count = (int)cmd.ExecuteScalar();
    
    if (count == 0)
    {
        SafeFileLogger.SafeAppendText("placement_errors.log",
            $"❌ CRITICAL: ClusterSleeves table insert failed! No record found for ClusterInstanceId={clusterInstanceId}\n");
    }
    else
    {
        SafeFileLogger.SafeAppendText("batch_v2.log",
            $"✅ Verified: ClusterSleeves table has {count} row(s) for ClusterInstanceId={clusterInstanceId}\n");
    }
}
```

---

## STEP 2: Check the ClashZones Table (After Placement)

### The Question:
How many sleeves have `MarkedForClusterProcess = 1` and `IsClusterResolvedFlag = 1`?

### SQL Query:
```sql
SELECT 
    COUNT(*) as total,
    SUM(CASE WHEN MarkedForClusterProcess = 1 THEN 1 ELSE 0 END) as marked_for_cluster,
    SUM(CASE WHEN IsClusterResolvedFlag = 1 THEN 1 ELSE 0 END) as is_cluster_resolved
FROM ClashZones
WHERE ComboId = YOUR_COMBO_ID;
```

### Expected vs Actual:

**EXPECTED:**
```
total: 6
marked_for_cluster: 4  ← Only the 4 in the cluster
is_cluster_resolved: 4 ← Only the 4 in the cluster
```

**ACTUAL (BUG):**
```
total: 6
marked_for_cluster: 6  ← ALL 6! ❌
is_cluster_resolved: 6 ← ALL 6! ❌
```

---

## STEP 3: Determine Which Component Failed

### Decision Logic:

```
IF ClusterSleeves_v2 has:
├─ CASE A: 1 cluster with all 6 sleeves
│  └─→ PROBLEM: Clustering algorithm (Section A)
│      FIX: Reduce tolerance or tighten host validation
│
└─ CASE B: 3 clusters (4, 1, 1)
   └─→ PROBLEM: Flag/persistence layer (Section B)
       └─ Check logs for "Cannot save" or "Failed to save"
       └─ FIX: Apply Fix B1 and B2 above
```

---

## STEP 4: Add Complete Logging

**File:** `BatchClusterPlacementService.cs`

```csharp
private void PerformSwapDeletion(Document doc, BatchClusterData cluster, int clusterElementId)
{
    // ✅ LOG: Entry point
    SafeFileLogger.SafeAppendText("debug_db.log",
        $"\n╔════════════════════════════════════════╗\n" +
        $"║ PerformSwapDeletion START              ║\n" +
        $"║ ClusterGUID: {cluster.ClusterGUID}  ║\n" +
        $"║ ClusterInstanceId: {clusterElementId,-28}║\n" +
        $"╚════════════════════════════════════════╝\n");

    if (string.IsNullOrEmpty(cluster.ConstituentZoneGuids))
    {
        SafeFileLogger.SafeAppendText("debug_db.log", 
            "❌ EARLY RETURN: No ConstituentZoneGuids\n");
        return;
    }

    var guids = new List<Guid>();
    var guidStrings = cluster.ConstituentZoneGuids.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

    foreach (var guidStr in guidStrings)
    {
        var trimmed = guidStr.Trim();
        if (Guid.TryParse(trimmed, out Guid parsedGuid))
        {
            guids.Add(parsedGuid);
        }
    }

    SafeFileLogger.SafeAppendText("debug_db.log",
        $"Parsed GUIDs: {guids.Count} (from {guidStrings.Length} strings)\n");

    var zones = _repository.GetClashZonesByGuids(guids);

    SafeFileLogger.SafeAppendText("debug_db.log",
        $"Found zones: {zones.Count}\n" +
        $"Zones: {string.Join(",", zones.Select(z => z.Id.ToString().Substring(0, 8)))}\n");

    // ... rest of code ...

    SafeFileLogger.SafeAppendText("debug_db.log",
        $"✅ Updated {updates.Count} zones with cluster flags\n\n");
}

private void SaveToClusterSleevesLegacy(Document doc, BatchClusterData cluster, int clusterInstanceId)
{
    SafeFileLogger.SafeAppendText("batch_v2.log",
        $"\n[{DateTime.Now:HH:mm:ss}] 🔍 SaveToClusterSleevesLegacy START\n" +
        $"    ClusterGUID: {cluster.ClusterGUID}\n" +
        $"    ClusterInstanceId: {clusterInstanceId}\n");

    try
    {
        // ... existing code ...

        if (comboId <= 0 || filterId <= 0)
        {
            string error = $"Cannot save: ComboId={comboId}, FilterId={filterId}";
            SafeFileLogger.SafeAppendText("placement_errors.log",
                $"[{DateTime.Now:HH:mm:ss}] ❌ {error}\n");
            
            // ✅ THROW instead of silent return
            throw new InvalidOperationException(error);
        }

        // ... save code ...

        SafeFileLogger.SafeAppendText("batch_v2.log",
            $"[{DateTime.Now:HH:mm:ss}] ✅ SaveToClusterSleevesLegacy COMPLETE\n");
    }
    catch (Exception ex)
    {
        SafeFileLogger.SafeAppendText("placement_errors.log",
            $"[{DateTime.Now:HH:mm:ss}] ❌ SaveToClusterSleevesLegacy FAILED: {ex.Message}\n");
        
        // ✅ RETHROW so caller knows
        throw;
    }
}
```

---

## STEP 5: Run the Test Again

With all logging in place:

```
1. Add the logging code above
2. Rebuild the project
3. Place 6 sleeves that should cluster as 4+1+1
4. Check logs in: %APPDATA%\JSE_MEP_Openings\Logs\R2023\
5. Review batch_v2.log and debug_db.log
```

---

## FINAL CHECKLIST

```
[ ] Step 1: Check ClusterSleeves_v2 table (1 cluster with all 6, or 3 clusters?)
    
    [ ] If 1 cluster with all 6:
        [ ] Apply Section A fixes (reduce tolerance or tighten validation)
        [ ] Re-test

    [ ] If 3 clusters (4,1,1):
        [ ] Check batch_v2.log for "Cannot save" messages
        [ ] Apply Section B fixes
        [ ] Re-test

[ ] Step 2: Check ClashZones table (how many marked as clustered?)
    [ ] If all 6: That confirms ClusterSleeves_v2 has 1 cluster
    [ ] If 4: That means persistence is partially working

[ ] Step 3: Add comprehensive logging
    [ ] Add logging to FormClusters()
    [ ] Add logging to PerformSwapDeletion()
    [ ] Add logging to SaveToClusterSleevesLegacy()

[ ] Step 4: Re-run test with logging enabled
    [ ] Check which section the problem is in

[ ] Step 5: Apply appropriate fix
    [ ] Section A: Algorithm fix
    [ ] Section B: Persistence fix
```

---

**Status:** Ready for targeted debugging ✅

**Next Action:** Run the SQL queries to determine if problem is in cluster CALCULATION or cluster PERSISTENCE.
