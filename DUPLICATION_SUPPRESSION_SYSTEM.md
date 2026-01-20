# Duplication Suppression System Optimization

This document outlines the plan to optimize the duplication checking system for MEP openings by combining section box filtering with a spatial grid.

## Problem

The current duplication checking mechanism iterates through all existing sleeves in the model, which is inefficient and causes performance bottlenecks in large models.

## Proposed Solution

To optimize the duplication checking process, I will implement a multi-layered filtering system that combines the strengths of both section box filtering and spatial partitioning.

### Plan

1.  **Filter by Section Box**:
    *   At the start of the sleeve placement process, all existing sleeves in the model will be collected.
    *   This list will be filtered using the active 3D view's section box, immediately discarding any sleeves outside the user's area of interest.

2.  **Build a `SleeveSpatialGrid`**:
    *   A `SleeveSpatialGrid` instance will be created using only the sleeves that passed the section box filter. This makes the grid smaller, faster to build, and more relevant to the current view.

3.  **Optimized Duplication Check**:
    *   For each new sleeve placement, the `SleeveSpatialGrid` will be queried to get a small, pre-filtered list of nearby sleeves.
    *   This localized list will then be passed to the `OpeningDuplicationChecker.IsAnySleeveAtLocationOptimized` method for a final, high-precision check.

4.  **Update Placer Services**:
    *   This new workflow will be implemented in `DuctSleevePlacerService`, `PipeSleevePlacerService`, and `CableTraySleeveCommand` to ensure consistent performance improvements across all sleeve placement tools.

This approach will significantly reduce the complexity of the duplication check from O(k) to nearly O(1) for each new sleeve, where 'k' is the total number of existing sleeves. This will result in a substantial performance improvement, especially in large and dense models.
