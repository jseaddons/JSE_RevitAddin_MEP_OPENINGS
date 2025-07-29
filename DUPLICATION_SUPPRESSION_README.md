# Duplication Suppression System for MEP Openings

## Overview
This document explains the comprehensive duplication suppression system that prevents conflicting sleeve placements across different MEP opening commands. The system ensures that individual sleeve placement commands and cluster replacement commands work together without creating overlapping or duplicate openings.

## Core Principle
**Each command type only checks for and avoids its own type of openings - no cross-type interference.**

## Command Types and Their Domains

### 1. Individual Sleeve Placement Commands
These commands place individual sleeves for specific MEP element types:

#### PipeSleeveCommand
- **Places**: Individual pipe sleeves (`PS#` families)
- **Must avoid**: Existing rectangular pipe cluster openings (`PipeOpeningOnWallRect`, `PipeOpeningOnSlabRect`)
- **Logic**: Before placing a `PS#` sleeve, check if there's already a `PipeOpeningOnWallRect/PipeOpeningOnSlabRect` in that area

#### DuctSleeveCommand  
- **Places**: Individual duct sleeves (`DS#` families)
- **Must avoid**: Existing rectangular duct cluster openings (`DuctOpeningOnWall`, `DuctOpeningOnSlab`)
- **Logic**: Before placing a `DS#` sleeve, check if there's already a `DuctOpeningOnWall/DuctOpeningOnSlab` in that area

#### CableTraySleeveCommand
- **Places**: Individual cable tray sleeves (`CT#` families)
- **Must avoid**: Existing rectangular cable tray cluster openings (`CableTrayOpeningOnWall`, `CableTrayOpeningOnSlab`)
- **Logic**: Before placing a `CT#` sleeve, check if there's already a `CableTrayOpeningOnWall/CableTrayOpeningOnSlab` in that area

### 2. Cluster Replacement Commands
These commands replace groups of individual sleeves with larger rectangular openings:

#### PipeOpeningsRectCommand
- **Processes**: Groups of circular pipe sleeves (`PS#` families)
- **Creates**: Rectangular pipe openings (`PipeOpeningOnWallRect` family)
- **Must avoid**: Existing rectangular pipe openings of the same type
- **Logic**: Before placing a `PipeOpeningOnWallRect`, check if there's already another `PipeOpeningOnWallRect` in that area

#### RectangularSleeveClusterCommand
- **Processes**: Existing rectangular sleeves (`DS#`, `DMS#`, `CT#` families)
- **Creates**: Larger clustered versions of the same family types
- **Must avoid**: Existing cluster openings of the same specific type
- **Logic**: Before clustering `DS#` sleeves, check if there's already a clustered `DS#` opening in that area

## Duplication Suppression Implementation

### Tolerance Distance
- **Standard tolerance**: 100mm (configurable)
- **Used for**: Proximity checking between existing and proposed openings
- **Applied consistently**: Across all commands for uniform behavior

### Individual Sleeve Command Logic
```csharp
// Example for PipeSleeveCommand
var existingClusterOpenings = new FilteredElementCollector(doc)
    .OfClass(typeof(FamilyInstance))
    .Cast<FamilyInstance>()
    .Where(fi => fi.Symbol?.Family?.Name != null &&
           (fi.Symbol.Family.Name.Contains("PipeOpeningOnWallRect") ||
            fi.Symbol.Family.Name.Contains("PipeOpeningOnSlabRect")))
    .Where(fi => IsWithinTolerance(proposedLocation, fi.Location, toleranceDistance))
    .ToList();

if (existingClusterOpenings.Any())
{
    // Skip placing individual sleeve - cluster opening already exists
    return;
}
```

### Cluster Command Logic
```csharp
// Example for RectangularSleeveClusterCommand
var existingSameTypeCluster = new FilteredElementCollector(doc)
    .OfClass(typeof(FamilyInstance))
    .Cast<FamilyInstance>()
    .Where(fi => fi.Symbol?.Family?.Name == clusterSymbol.Family.Name &&
                 fi.Symbol.Name == clusterSymbol.Name)
    .Where(fi => IsWithinTolerance(clusterCenter, fi.Location, toleranceDistance))
    .ToList();

if (existingSameTypeCluster.Any())
{
    // Skip clustering - same type cluster already exists
    return;
}
```

## Execution Order and Coordination

### ClusterSleevesReplacementCommand Workflow
1. **PipeOpeningsRectCommand** executes first
   - Processes `PS#` → creates `PipeOpeningOnWallRect`
   - Has internal duplication suppression for its own type

2. **RectangularSleeveClusterCommand** executes second  
   - Processes `DS#`, `DMS#`, `CT#` → creates clustered versions
   - Has internal duplication suppression for each specific type
   - Does NOT interfere with pipe openings (different domain)

### Re-execution Safety
When commands are re-run:
- **Individual sleeve commands** detect existing cluster openings and skip placement in those areas
- **Cluster commands** detect existing clusters of the same type and skip re-clustering
- **No cross-interference** between different MEP types (pipe, duct, cable tray)

## Why This System Works

### Domain Separation
- **Pipes**: `PS#` sleeves ↔ `PipeOpeningOnWallRect` clusters
- **Ducts**: `DS#` sleeves ↔ `DuctOpeningOnWall` clusters  
- **Cable Trays**: `CT#` sleeves ↔ `CableTrayOpeningOnWall` clusters
- **Each domain is independent** - no cross-checking between types

### Hierarchical Relationship
- Individual sleeves are **replaced by** cluster openings of the same type
- Individual sleeve commands **respect** existing cluster openings
- Cluster commands **avoid duplicating** existing clusters of the same type

### Logging and Debugging
All duplication suppression actions are logged with:
- Element IDs involved
- Distances calculated  
- Reasons for skipping placement
- Family names and types checked

## Expected Behavior Examples

### Scenario 1: First Run
1. `PipeSleeveCommand` places individual `PS#` sleeves
2. `DuctSleeveCommand` places individual `DS#` sleeves  
3. `ClusterSleevesReplacementCommand` runs:
   - `PipeOpeningsRectCommand` clusters `PS#` → `PipeOpeningOnWallRect`
   - `RectangularSleeveClusterCommand` clusters `DS#` → larger `DS#`

### Scenario 2: Re-run Individual Commands
1. `PipeSleeveCommand` detects existing `PipeOpeningOnWallRect` and skips those areas
2. `DuctSleeveCommand` detects existing clustered `DS#` and skips those areas
3. Only new/unclustered areas get individual sleeves

### Scenario 3: Re-run Cluster Commands
1. `PipeOpeningsRectCommand` detects existing `PipeOpeningOnWallRect` and skips re-clustering
2. `RectangularSleeveClusterCommand` detects existing clustered `DS#` and skips re-clustering
3. No duplicate cluster openings are created

## Implementation Status

### ✅ Currently Implemented
- RectangularSleeveClusterCommand: Same-type duplication suppression
- PipeOpeningsRectCommand: Same-type duplication suppression  
- ClusterSleevesReplacementCommand: Coordination and logging

### 🔄 Needs Implementation
- PipeSleeveCommand: Check for existing `PipeOpeningOnWallRect/PipeOpeningOnSlabRect`
- DuctSleeveCommand: Check for existing `DuctOpeningOnWall/DuctOpeningOnSlab`
- CableTraySleeveCommand: Check for existing `CableTrayOpeningOnWall/CableTrayOpeningOnSlab`

### 🎯 Success Criteria
- ✅ No duplicate openings of the same type
- ✅ Individual sleeves don't place where clusters exist
- ✅ Clusters don't re-process already clustered areas
- ✅ Different MEP types work independently
- ✅ Safe for multiple re-executions
- ✅ Clear logging for debugging and verification

## Structural Framing Sleeve Depth Rule (2025-07-05)

- When placing sleeves in structural framing (beams, etc), the add-in will set the sleeve 'Depth' parameter to the value of the host's type parameter 'b' (or 'Width').
- If neither 'b' nor 'Width' is set on the host, the add-in will throw an error and notify the user: **Cannot determine structural framing depth: 'b' or 'Width' parameter not set for element ID ...**
- There is no fallback to 'h' or 'Depth' for structural framing. This ensures the sleeve always matches the true physical depth of the beam/host.
- For floors, the sleeve 'Depth' is set from the 'Default Thickness' parameter as before.
