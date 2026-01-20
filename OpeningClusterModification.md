# Opening Cluster Modification

This document describes the logic behind the pipe sleeve clustering and rectangular opening replacement feature. These clustered openings are **linked to wall** geometry (no direct wall probing), and operate purely on sleeve proximity.

## Purpose
- Prevent multiple closely-spaced circular pipe sleeves from weakening a wall.
- Merge any group of sleeves within a configurable tolerance into one larger rectangular opening.

## Key Concepts
1. **Sleeve Collection**  
   - Filter all placed pipe sleeves (family instances beginning with `PS#OpeningOnWall`).
   - Operate on their insertion points in world XY coordinates.

2. **Proximity Clustering**  
   - Use a planar (XY) tolerance (default 100 mm).  
   - Build clusters: each sleeve in a cluster has at least one neighbor within tolerance.

3. **Cluster Replacement**  
   - For each cluster of size ≥ 2:
     - Compute the min/max X and Y of all sleeve origins.
     - Center point = average of min/max extents (Z from average sleeve elevation).
     - Width = max(spanX, spanY); Height = sleeve diameter.
     - Place one `PipeOpeningOnWallRect` family instance at center with calculated Width/Height.
     - Delete all original circular sleeves in the cluster.

4. **Linked Wall Note**  
   - The opening families remain **linked to the host wall** (no direct wall intersection logic).
   - Clustering and placement are driven by sleeve proximity, not by exact wall face geometry.

## Configuration
- **Tolerance**: default 100 mm, adjustable via `UnitUtils.ConvertToInternalUnits(…)` in the command code.
- **Family Symbol**: looks for `PipeOpeningOnWallRect` in active document. Must be loaded and activated before placement.

## Logging
- Detailed diagnostic logs are written to `PipeOpeningsRectCommand.log` in the user’s Documents folder.
- Logs include sleeve counts, cluster summaries, placement IDs, and deletion counts.

Refer to the **`PipeOpeningsRectCommand`** implementation for full code details and comments.
