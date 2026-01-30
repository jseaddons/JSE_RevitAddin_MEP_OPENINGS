# Architecture: Unified Placement & Code Reuse

## Executive Summary
To address the lack of robustness in current "Fast Path" implementations and eliminate code duplication, we are moving to a **Unified Planning-Execution Pattern**. This architecture ensures that every sleeve, whether placed individually or in bulk, follows the exact same decision logic derived from centralized services.

---

## 1. Centralized "Decision" Services (The Working Code)
We will strictly reuse these existing services to avoid "reinventing the wheel":

| Service | Responsibility | Logic Source |
| :--- | :--- | :--- |
| **ConfigurationResolutionService** | Resolves Opening Shape (Circular vs Rectangular) | **Global Rules** (e.g., Pipes > 200mm → Rectangular) |
| **InsulationAwareSizingService** | Calculates Dimensions (Width/Height/Diameter) | **Insulation + Clearance + Rounding** |
| **FamilyManager** | Selects Revit Family | **Host Type + Element Shape** mapping |
| **SleeveRotationService** | Calculates Rotation Angle | **Host Orientation + Element Direction** |
| **PlacementPointAdjustmentService** | Calculates XYZ Point | **Host Centerline + Wall Face Alignment** |

---

## 2. The Unified Planning Phase
Instead of having different logic for "Normal" vs "Saved Data" placement, both paths will now use the **ParallelSleevePlacementPlanner** (or its sequential equivalent) to generate a standardized instruction set.

### **The Single Source of Truth: `SleevePlacementPlanningDto`**
The DTO (Data Transfer Object) will now carry all decisions made by the services above:
- `SleeveFamilyName` (From `FamilyManager`)
- `TargetWidth`, `TargetHeight`, `TargetDiameter` (From `SizingService`)
- `PlacementPoint` (From `AdjustmentService`)
- `RotationRadians` (From `RotationService`)
- `IsCircular` (From `ConfigurationService`)

---

## 3. Deployment Strategy: "Plan Once, Execute Anywhere"

### **Robust "Fast Path" (Saved Data)**
Previously, the fast path used raw DB values and bypassed business rules. 
**New Logic**:
1. Check if Saved Data is available.
2. If yes, the **Planner** validates the saved data against the latest **Configuration Rules**.
3. If rules have changed (e.g., a pipe size now exceeds the 200mm limit), the Planner **overrides** the saved data using the centralized services.
4. Execution consumes the resulting `SleevePlacementPlanningDto`.

### **Bulk Placement**
The bulk path (`BulkPlacementService`) currently lacks family names and dimensions.
**New Logic**:
1. `OpeningCommandOrchestrator` calls the **Parallel Planner** for the entire batch.
2. The Planner uses the same services as individual placement.
3. Batch execution consumes the DTOs provided by the Planner.

---

## 4. Why this is Robust
- **Consistency**: A Pipe > 200mm will *always* get a Rectangular Opening, regardless of the placement mode.
- **Maintainability**: Changing a rounding rule in `InsulationAwareSizingService` instantly updates all placement paths.
- **Reusability**: `NewSleevePlacerService` becomes a thin orchestrator that simply "executes the plan" rather than "calculating the plan."

---

## 5. Implementation Roadmap
1. [x] **Static Logic Extraction**: Make `FamilyManager` methods static for shared utility.
2. [x] **DTO Enhancement**: Add full coordinate/rotation properties to `SleevePlacementPlanningDto`.
3. [ ] **Planner Update**: Integrate all 5 centralized services into `ParallelSleevePlacementPlanner`.
4. [ ] **Service Refactor**: Update `NewSleevePlacerService` to call the Planner for both its internal loops.
5. [ ] **Orchestrator Update**: Ensure Bulk path uses the updated Planner before execution.
