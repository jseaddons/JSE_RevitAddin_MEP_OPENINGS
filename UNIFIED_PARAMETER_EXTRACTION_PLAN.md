## Parameter Names to Capture (As in Working Code)

### For Dampers (MEP Elements)
- **Level-related (first available, prioritized):**
  - Reference Level
  - Level
  - Schedule Level
  - Schedule of Level
  - Reference Level Elevation
- **Essential MEP parameters:**
  - Size
  - Diameter
  - Width
  - Height
  - MEP Size
  - Nominal Size
  - Actual Size
  - Duct Size
  - Damper Size
  - Outside Diameter
  - System Type
  - System Name
  - System Abbreviation
  - Service Type
  - System Classification
  - Mark
  - (Optionally: Damper Width, Damper Height, DIMENSIONS_WIDTH, DIMENSIONS_HEIGHT, DIMENSION_DIAMETER)

### For Host Elements (Walls, etc.)
- Level
- Fire Rating
- (Optionally: Host orientation, thickness, centerline, etc. if required for downstream logic)

**Note:**
- The extraction logic should use prioritized/fallback lists for parameters like width/height/diameter to ensure robust capture across different family naming conventions.
- All captured parameters should be included in the `AllParameters` dictionary for full traceability and future-proofing.


# Unified Parameter Extraction and DB Dumping Plan (SOLID & Optimized)

## Overview
This document details a robust, maintainable, and optimized plan to consolidate all parameter extraction and database (DB) dumping logic for both MEP and host elements into a single, unified service. The plan is explicitly designed to follow SOLID principles and enable future optimizations.

---

## Why Consolidate?
- **Reduce duplication:** Avoid multiple helpers/services doing similar work.
- **Consistency:** Ensure all parameter extraction (for MEP and host) follows the same logic and structure.
- **Maintainability:** Make it easier to update, debug, or extend parameter handling.
- **Performance:** Enable efficient batch extraction and DB dumping.
- **SOLID compliance:** Ensure the design is modular, extensible, and testable.

---

## Step-by-Step Implementation Plan (with SOLID Principles)

### 1. Design a Unified Parameter Snapshot Structure (Single Responsibility, Open/Closed)
- Create a class (e.g., `ElementParameterSnapshot`) with fields for:
  - Common parameters: `Width`, `Height`, `Diameter`, `Level`, `FireRating`, etc.
  - `AllParameters` dictionary for raw key-value pairs.
  - Optional: Host orientation, thickness, centerline, etc.
- **Single Responsibility:** This class only represents a snapshot of parameters.
- **Open/Closed:** Add new fields or logic by extending, not modifying, the class.

### 2. Implement a Unified Extraction Service (Single Responsibility, Dependency Inversion, Liskov Substitution)
- Create a service (e.g., `IElementParameterExtractor` interface and `ElementParameterSnapshotService` implementation) that:
  - Accepts any `Element` (MEP or host).
  - Extracts all required parameters using prioritized/fallback logic.
  - Handles both MEP and host-specific needs (e.g., damper width/height, host level).
  - Returns an `ElementParameterSnapshot`.
- **Dependency Inversion:** Consumers depend on the interface, not the implementation.
- **Liskov Substitution:** Any new extractor (for new element types) can be swapped in without breaking consumers.

### 3. Use Strategy/Factory for Extensibility (Open/Closed, Single Responsibility)
- Use a factory or strategy pattern to select the correct extraction logic for different element types (MEP, host, etc.).
- This allows easy extension for new categories without modifying existing code.

### 4. Refactor All Extraction Points (Single Responsibility, Interface Segregation)
- Replace all usages of `MepParameterHelper`, `HostLevelHelper`, etc., with the new service.
- Ensure all parameter extraction in the refresh phase uses the unified service.
- **Interface Segregation:** Only expose the minimal interface needed for consumers.

### 5. Centralize DB Dumping/Serialization (Single Responsibility, Open/Closed)
- Refactor DB/XML/JSON dumping logic (e.g., in `ClashZonePersistenceService`, `UpdateDbService`) to use the new snapshot structure.
- Ensure all data written to the DB comes from the unified snapshot.
- Use serialization interfaces for flexibility (e.g., `IParameterSnapshotSerializer`).

### 6. Remove Redundant Helpers
- Delete or deprecate old helpers/services once all usages are migrated.

### 7. Add Logging, Fallbacks, and Diagnostics (Single Responsibility)
- Add robust logging for missing/ambiguous parameters.
- Ensure fallback logic is clear and maintainable.
- Add diagnostics for performance and error tracking.

### 8. Optimize for Performance
- **Batch Extraction:** Extract parameters for all elements in a single pass where possible.
- **Caching:** Cache parameter snapshots to avoid repeated extraction.
- **Asynchronous/Bulk DB Writes:** Use bulk or batched DB writes to minimize I/O overhead.
- **Configurable Extraction:** Allow configuration of which parameters to extract/dump for different workflows.
- **Parallel Processing:** Where safe, use parallelism for extraction and serialization.

### 9. Test and Validate (Single Responsibility)
- Test extraction and DB dump for both MEP and host elements.
- Validate that all required data (including host level) is available for downstream use.
- Add unit and integration tests for extractors and serializers.

---

## Example: Host Level Extraction
- The unified service extracts and stores the host’s level (and other host parameters) in the snapshot.
- Downstream code (placement, DB dump, etc.) uses the snapshot, not direct API calls.

---

## Benefits
- **Single source of truth** for all parameter extraction.
- **Easier debugging** and extension.
- **Consistent data** for all downstream consumers (DB, XML, placement, etc.).
- **SOLID-compliant, modular, and extensible design.**
- **Optimized for performance and future scalability.**

---

## Next Steps
1. Design and implement `ElementParameterSnapshot`, `IElementParameterExtractor`, and `ElementParameterSnapshotService`.
2. Refactor all parameter extraction and DB dump logic to use the new service.
3. Remove/deprecate old helpers.
4. Add logging, diagnostics, and configuration options.
5. Test and validate end-to-end.

---

*For questions or to start implementation, contact the project maintainer.*
