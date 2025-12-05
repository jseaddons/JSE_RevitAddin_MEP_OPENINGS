# ClusterPlacementService Comparison Report
## Working Version (Desktop) vs Current Codebase

**Date:** 2025-01-XX  
**Purpose:** Identify missing placement logic that prevents cluster sleeves from being placed

---

## ✅ **STRUCTURAL COMPARISON**

Both versions are **structurally identical**:
- Same class structure
- Same method signatures
- Same dependency injection pattern
- Same caching mechanisms

---

## 🔍 **DETAILED METHOD COMPARISON**

### 1. **PlaceClusterSleeve() Method**
**Status:** ✅ **IDENTICAL**

Both versions have:
- ✅ Input validation
- ✅ Family loading logic
- ✅ Reference level retrieval
- ✅ Wall/Framing midpoint override with Z-delta protection (>1000mm threshold)
- ✅ Family instance creation with profiling
- ✅ Immediate ID capture
- ✅ Size parameter setting
- ✅ Rotation logic (X-wall, Y-wall, floor)
- ✅ Metadata setting
- ✅ Cluster resolution marking

**No differences found.**

---

### 2. **SetSizeParameters() Method**
**Status:** ✅ **IDENTICAL**

Both versions have:
- ✅ RCS dimension mapping for walls/framing
- ✅ Wall thickness override for depth
- ✅ Deferred parameter support (batching)
- ✅ Immediate parameter setting (non-batching)

**No differences found.**

---

### 3. **SetMetadata() Method** ⚠️ **CRITICAL FOR PLACEMENT**
**Status:** ✅ **IDENTICAL**

Both versions have:
- ✅ MEP_Category parameter setting (deferred/immediate)
- ✅ Filter Name parameter setting (deferred/immediate)
- ✅ **IMMEDIATE** setting of "Sleeve Instance ID" = -1 (critical for cleanup)
- ✅ **IMMEDIATE** setting of "Cluster Sleeve Instance ID" (critical for cleanup)
- ✅ Verification logic with `Regenerate()` call
- ✅ Diagnostic logging

**No differences found.**

---

### 4. **ApplyRotation() Method**
**Status:** ✅ **IDENTICAL**

Both versions have:
- ✅ Protected code flag check
- ✅ Input validation
- ✅ Zero-angle skip logic
- ✅ Category-specific rotation (cable tray, duct, pipe)
- ✅ Wall vs floor rotation logic
- ✅ 90° offset handling (skip for cable trays/ducts, apply for pipes)
- ✅ Angle normalization

**No differences found.**

---

### 5. **ComputeClusterMidpoint() Method**
**Status:** ✅ **IDENTICAL**

Both versions have:
- ✅ Corner-based midpoint calculation (primary)
- ✅ WCS bounding box fallback
- ✅ Z-delta protection (>1000mm threshold)
- ✅ Diagnostic logging

**No differences found.**

---

### 6. **ComputeCornerBasedClusterMidpoint() Method**
**Status:** ✅ **IDENTICAL**

Both versions have:
- ✅ Corner collection from ClashZone (SleeveCorner1X/Y/Z through SleeveCorner4X/Y/Z)
- ✅ Validation (minimum 4 corners)
- ✅ Centroid calculation
- ✅ Fallback to WCS midpoint

**No differences found.**

---

### 7. **ComputeWcsClusterMidpoint() Method**
**Status:** ✅ **IDENTICAL**

Both versions have:
- ✅ Bounding box union logic
- ✅ Uninitialized bbox detection
- ✅ Centroid calculation

**No differences found.**

---

## 🚨 **CRITICAL FINDINGS**

### **NO DIFFERENCES FOUND IN ClusterPlacementService.cs**

The working version and current version are **functionally identical**. This suggests:

1. **The issue is NOT in ClusterPlacementService.cs itself**
2. **The problem is likely in:**
   - How `ClusterPlacementService` is being **called** (calling code)
   - How **deferred parameters are flushed** (if batching is enabled)
   - How **cluster sleeves are identified** after placement (cleanup service)
   - **Transaction management** (when parameters are committed)
   - **Element validity** (sleeves becoming invalid after creation)

---

## 🔍 **POTENTIAL ROOT CAUSES (Outside ClusterPlacementService)**

### 1. **Deferred Parameter Flushing**
**Location:** `RefactoredClusterService.cs` or calling code
- If `UseBatchedParameterWrites` is enabled, parameters are deferred
- Critical parameters (Sleeve Instance ID, Cluster Sleeve Instance ID) are set **immediately** in `SetMetadata()`
- But other parameters (Width, Height, Depth, MEP_Category) are deferred
- **Check:** Is `FlushDeferredClusterParameters()` being called **before** cleanup service runs?

### 2. **Transaction Commit Timing**
**Location:** `RefactoredClusterService.cs` or `OpeningCommandOrchestrator.cs`
- Cluster sleeve is created in a transaction
- Parameters are set (immediate or deferred)
- Transaction must be **committed** before cleanup service can find the sleeve
- **Check:** Is transaction committed **after** `PlaceClusterSleeve()` returns?

### 3. **Element Reference Validity**
**Location:** `RefactoredClusterService.cs`
- `PlaceClusterSleeve()` returns `placedClusterSleeve` and `capturedClusterSleeveId`
- If the element reference becomes stale, cleanup service won't find it
- **Check:** Is the calling code using `capturedClusterSleeveId` to look up the sleeve **freshly** from the document?

### 4. **Cleanup Service Misidentification**
**Location:** `CleanupService.cs` or similar
- Cleanup service identifies cluster sleeves by checking "Sleeve Instance ID" = -1
- If this parameter is not set **before** cleanup runs, sleeve will be misidentified
- **Check:** Is cleanup service running **after** `SetMetadata()` completes?

### 5. **Flag Management**
**Location:** `FlagManager` or `IFlagManager`
- Cluster sleeves must be registered in flag management
- If flags are not updated correctly, sleeves may be deleted
- **Check:** Is `BatchUpdateFlagsForPlacement()` being called with `isCluster: true`?

---

## 📋 **RECOMMENDATIONS**

### **Immediate Actions:**

1. **Check `RefactoredClusterService.PlaceClusterForGroup()`:**
   - Verify `PlaceClusterSleeve()` return value is checked
   - Verify `capturedClusterSleeveId` is used to look up sleeve **freshly** from document
   - Verify transaction is committed **after** placement

2. **Check Deferred Parameter Flushing:**
   - Verify `FlushDeferredClusterParameters()` is called **before** cleanup service
   - Verify it's called **within the same transaction** as placement

3. **Check Cleanup Service:**
   - Verify cleanup service checks "Sleeve Instance ID" = -1 **correctly**
   - Verify cleanup service runs **after** transaction commit
   - Verify cleanup service uses **fresh element lookup** (not stale references)

4. **Check Flag Management:**
   - Verify `BatchUpdateFlagsForPlacement()` is called with correct parameters
   - Verify `isCluster: true` is passed for cluster sleeves
   - Verify flags are persisted to database **before** cleanup runs

5. **Add Diagnostic Logging:**
   - Log immediately after `PlaceClusterSleeve()` returns
   - Log after transaction commit
   - Log before/after cleanup service runs
   - Log element lookup results (found vs. not found)

---

## ✅ **CONCLUSION**

**ClusterPlacementService.cs is NOT the problem.** The working version and current version are identical.

**The issue is in the CALLING CODE or INTEGRATION POINTS:**
- Transaction management
- Deferred parameter flushing
- Element reference handling
- Cleanup service timing
- Flag management

---

## 🔍 **CALLING CODE ANALYSIS**

### **RefactoredClusterService.PlaceClusterForGroup() Flow:**

1. ✅ Calls `PlaceClusterSleeve()` (line 1286)
2. ✅ Verifies placement success (line 1313)
3. ✅ Stores rotation data (line 1330)
4. ✅ Updates flags using `BatchUpdateFlagsForPlacement()` (line 1394)
5. ✅ Verifies cluster sleeve exists BEFORE deleting individual sleeves (line 1444)
6. ✅ Deletes individual sleeves (line 1474+)
7. ✅ Verifies cluster sleeve exists AFTER deleting individual sleeves (line 1557)

### **Critical Observations:**

1. **Deferred Parameter Flushing:**
   - `FlushDeferredClusterParameters()` is called **AFTER all clusters are placed** (line 610, 1822)
   - **NOT called within `PlaceClusterForGroup()` itself**
   - Critical parameters (Sleeve Instance ID, Cluster Sleeve Instance ID) are set **IMMEDIATELY** in `SetMetadata()`, so this should not be the issue
   - **However:** If batching is enabled, Width/Height/Depth/MEP_Category are deferred and won't be set until flush

2. **Transaction Management:**
   - `PlaceClusterSleeve()` is called **within a transaction** (assumed, not explicitly shown in snippet)
   - Transaction must be **committed** before cleanup service runs
   - **Check:** Is transaction committed **after** `PlaceClusterForGroup()` returns?

3. **Element Reference Validity:**
   - Code uses `capturedClusterSleeveId` to look up sleeve **freshly** from document (line 1444, 1559)
   - This is correct - avoids stale reference issues
   - **However:** If element is deleted between placement and verification, it will be lost

4. **Flag Management:**
   - `BatchUpdateFlagsForPlacement()` is called with `isCluster: true` (line 1396)
   - This should mark clash zones as cluster-resolved
   - **Check:** Are flags persisted to database **before** cleanup runs?

5. **Cleanup Service:**
   - Cleanup service is called **after** `FlushDeferredClusterParameters()` (line 647, 1825)
   - This is correct - parameters should be set before cleanup
   - **However:** If cleanup service misidentifies cluster sleeves, they may be deleted

---

## 🚨 **POTENTIAL ROOT CAUSES**

### **1. Transaction Not Committed**
**Symptom:** Cluster sleeve created but not visible/accessible after placement  
**Check:** Verify transaction is committed **after** `PlaceClusterForGroup()` returns  
**Location:** `OpeningCommandOrchestrator.cs` or transaction wrapper

### **2. Deferred Parameters Not Flushed Before Cleanup**
**Symptom:** Cluster sleeve exists but parameters are wrong/missing  
**Check:** Verify `FlushDeferredClusterParameters()` is called **before** cleanup service  
**Status:** ✅ **CORRECT** - Flush is called before cleanup (line 610, 1822)

### **3. Cleanup Service Misidentification**
**Symptom:** Cluster sleeve deleted by cleanup service  
**Check:** Verify cleanup service checks "Sleeve Instance ID" = -1 correctly  
**Location:** `CleanupService.cs` or similar

### **4. Element Deleted During Individual Sleeve Deletion**
**Symptom:** Cluster sleeve exists before deletion, missing after  
**Check:** Verify individual sleeve deletion doesn't accidentally delete cluster sleeve  
**Status:** ✅ **PROTECTED** - Code checks `id.IntegerValue != capturedClusterSleeveId.Value` (line 1537)

### **5. Flag Management Not Persisting**
**Symptom:** Flags not saved to database, causing duplicate placement attempts  
**Check:** Verify `BatchUpdateFlagsForPlacement()` persists to database  
**Location:** `FlagManagerService.cs` or `IFlagManager` implementation

---

## 📋 **RECOMMENDATIONS**

### **Immediate Actions:**

1. **Check Transaction Commit:**
   - Verify transaction is committed **after** `PlaceClusterForGroup()` returns
   - Add logging: "Transaction committed" after commit

2. **Check Cleanup Service:**
   - Verify cleanup service checks "Sleeve Instance ID" = -1 correctly
   - Add logging: "Cleanup service checking cluster sleeve {id}"

3. **Check Flag Persistence:**
   - Verify `BatchUpdateFlagsForPlacement()` saves to database
   - Add logging: "Flags persisted for cluster sleeve {id}"

4. **Add Diagnostic Logging:**
   - Log immediately after `PlaceClusterSleeve()` returns
   - Log after transaction commit
   - Log before/after cleanup service runs
   - Log element lookup results (found vs. not found)

**Next Steps:**
1. Review `RefactoredClusterService.cs` (calling code) - ✅ **DONE** (see above)
2. Review `OpeningCommandOrchestrator.cs` (orchestration) - **TODO**
3. Review cleanup service (if separate) - **TODO**
4. Review flag management integration - **TODO**

