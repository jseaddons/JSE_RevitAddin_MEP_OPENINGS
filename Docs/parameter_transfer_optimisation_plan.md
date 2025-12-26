# Parameter Transfer & Mark/Remark Optimisation Plan

## Overview
This document outlines a refactoring and optimisation plan for the parameter transfer, mark, and remark operations in the Revit add-in. The goal is to dramatically improve performance, especially for large models with 1000+ sleeves, by addressing major bottlenecks in `ParameterTransferService.cs` and `SleeveParameterService.cs`.

---

## Major Performance Issues Identified

### 1. Logging Inside Loops
- **Problem:** Diagnostic and file logging occurs inside the main processing loop, causing thousands of file writes and significant I/O overhead.
- **Solution:** Buffer all log messages in memory (e.g., `StringBuilder`) and flush to file only once after the loop or periodically in large batches.

### 2. Inefficient Parameter Caching
- **Problem:** Parameter cache is initialized by eagerly looking up all parameters for all elements upfront, defeating the purpose of caching.
- **Solution:** Implement a lazy cache that only looks up and stores parameters when first needed.

### 3. Deferred Parameters Dictionary Structure
- **Problem:** Uses nested dictionaries with string keys and boxed `object` values, causing slow lookups and boxing/unboxing overhead.
- **Solution:** Replace with a struct-based approach, using `int` IDs and typed fields for parameter values.

### 4. Excessive LINQ Usage in Hot Paths
- **Problem:** Heavy use of LINQ in performance-critical loops creates unnecessary allocations and intermediate collections.
- **Solution:** Replace with simple `foreach` loops for hot paths.

### 5. File I/O in Critical Paths
- **Problem:** Multiple file writes per element processed.
- **Solution:** Use a log buffer and flush only after processing a batch or the entire operation.

### 6. Snapshot Lookup Fallback Chains
- **Problem:** Multiple dictionary lookups per element due to fallback logic.
- **Solution:** Build a unified lookup dictionary at the start and use a single lookup per element.

---

## Refactoring Steps

### Step 1: Remove/Buffer Logging Inside Loops
- Replace all direct file logging inside loops with a `StringBuilder` buffer.
- Flush the buffer to file after the loop or every N entries for very large batches.
- Disable all diagnostic logging in deployment mode.

### Step 2: Implement Lazy Parameter Cache
- Refactor parameter cache to only perform lookups when a parameter is first needed.
- Use a composite key (element ID + parameter name) for the cache.

### Step 3: Struct-Based Deferred Parameters
- Replace nested dictionary structure with a struct or class for deferred parameter values.
- Use `int` for element IDs and strongly-typed fields for parameter values.

### Step 4: Replace LINQ in Hot Paths
- Identify all LINQ usage in main processing loops and replace with explicit `foreach` loops.

### Step 5: Batch File I/O
- Use a log buffer for all file writes in critical paths.
- Flush buffer after processing or at regular intervals.

### Step 6: Unified Snapshot Lookup
- Build a single dictionary mapping all relevant IDs to snapshot views at the start of processing.
- Use this for all snapshot lookups in the main loop.

---

## Estimated Gains
- **Batch logging:** 8-12x faster
- **Lazy parameter cache:** 4x faster
- **Batch file I/O:** 50-100x fewer writes
- **Unified snapshot lookup:** 3-4x faster
- **Struct-based deferred params:** 2-3x faster
- **Combined:** 20-50x overall speedup for large models

---

## Next Steps
1. Review and approve this plan.
2. Refactor `ParameterTransferService.cs` and `SleeveParameterService.cs` according to the steps above.
3. Profile performance before and after refactoring.
4. Document results and update this plan as needed.

---

*Prepared: 2025-12-24*
