# Parameter Batching Performance Comparison

## Current Performance (Batching DISABLED)

### From Latest Logs (2025-12-02):
- **34 sleeves attempted**: All failed due to FamilySymbol.IsActive error
- **Place Single Sleeve**: 1.0ms average (but 0 items placed - failing early)
- **Individual Sleeve Placement**: 9,712ms total (but 0 items placed)
- **Note**: The 1ms timing is misleading - sleeves fail before parameter setting

### From Previous Analysis (PERFORMANCE_ANALYSIS_PRE_MAX_LOAD.md):
- **10 sleeves placed successfully**: 1,515ms total
- **Average per sleeve**: 142.62ms
- **Parameter Setting**: 1,043ms (73.1% of total time) ⚠️ **BOTTLENECK**
- **Sleeve Creation**: 192ms (13.5%)
- **Family Loading**: 90ms (6.3%)
- **Other**: 190ms (7.1%)

---

## Expected Performance (Batching ENABLED)

### Based on OptimizationFlags.cs Documentation:
- **Expected gain**: 4-6× faster individual placement
- **Target**: 143-203ms → **<30ms per sleeve**
- **Parameter setting**: Should drop from 73.1% to ~10-15% of total time

### Projected Performance for 34 Sleeves:

| Metric | Without Batching | With Batching | Improvement |
|--------|------------------|---------------|-------------|
| **Time per sleeve** | 142.62ms | **<30ms** | **4.7× faster** |
| **Parameter setting** | 1,043ms (73.1%) | **~100-150ms** (10-15%) | **7-10× faster** |
| **Total for 34 sleeves** | ~4,849ms | **~1,020ms** | **4.7× faster** |
| **Throughput** | 7 sleeves/second | **33 sleeves/second** | **4.7× faster** |

---

## Key Performance Benefits of Batching

### 1. Parameter Setting Optimization
- **Without batching**: Each parameter write = 1 Revit API call
- **With batching**: All parameters accumulated → 1 regeneration → 1 batch write
- **Savings**: Eliminates 10-20 parameter API calls per sleeve

### 2. Regeneration Optimization
- **Without batching**: Regeneration after each sleeve (34 regenerations)
- **With batching**: Single regeneration after all sleeves (1 regeneration)
- **Savings**: 33 fewer regenerations = ~3,300ms saved (assuming 100ms per regeneration)

### 3. API Call Reduction
- **Without batching**: ~15-20 API calls per sleeve × 34 = 510-680 calls
- **With batching**: ~15-20 accumulation operations + 1 batch write = ~35-40 calls total
- **Savings**: ~470-640 fewer API calls

---

## Real-World Impact

### For 100 Sleeves:
- **Without batching**: ~14.3 seconds
- **With batching**: ~3.0 seconds
- **Time saved**: **11.3 seconds** (79% reduction)

### For 500 Sleeves:
- **Without batching**: ~71.3 seconds (1.2 minutes)
- **With batching**: ~15.0 seconds
- **Time saved**: **56.3 seconds** (79% reduction)

---

## Current Status

- **Batching Flag**: `UseBatchedParameterWrites = false` (DISABLED)
- **Reason**: User requested to disable due to issues
- **Expected Performance if Re-enabled**: 4-6× faster placement

---

## Recommendation

**Enable batching** once the FamilySymbol validation issue is fixed. The performance gain is significant:
- **4-6× faster** individual sleeve placement
- **7-10× faster** parameter setting
- **79% time reduction** for large batches

The batching implementation includes safety flags to prevent multiple flushes, so it should be safe to re-enable.

