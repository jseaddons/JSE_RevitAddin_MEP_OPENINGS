# IMPLEMENTATION PLAN: Combined Sleeve Parameter Transfer (Database-Driven)

## Overview
This plan outlines the architecture for transferring parameters to Combined Sleeves. The process is strictly database-driven, relying on `SleeveSnapshots` and `CombinedSleeveConstituents` tables, avoiding live geometric intersection checks during the transfer phase.

**Core Concept:**
A Combined Sleeve in Revit (`CombinedInstanceId`) is linked to multiple constituent sleeves (Individual or Cluster) in the database. Parameter transfer works by aggregating the data stored in the snapshots of these constituents.

## Data Flow Architecture

### 1. Source of Truth
*   **Individual Sleeves**: Data stored in `SleeveSnapshots` table.
    *   **Key**: `ClashZoneGuid` (Primary & Deterministic) OR `SleeveInstanceId`.
*   **Cluster Sleeves**: Data stored in `SleeveSnapshots` table.
    *   **Key**: `ClusterInstanceId`.
*   **Combined Sleeves**: Structure stored in `CombinedSleeves` and `CombinedSleeveConstituents` tables.
    *   **Key**: `CombinedInstanceId` (Revit Element ID).

### 2. Lookup Sequence
When `ParameterTransferService` processes a Combined Sleeve:

1.  **Identify**: The service receives a Revit Element (`FamilyInstance`) which is identified as a Combined Sleeve.
2.  **Retrieve Structure**:
    *   Use `CombinedInstanceId` (Revit ID) to query the `CombinedSleeveConstituents` table.
    *   This returns a list of constituents, each defined by:
        *   `ConstituentType` (Individual or Cluster)
        *   `ClashZoneGuid` (for Individual constituents)
        *   `ClusterInstanceId` (for Cluster constituents)
3.  **Fetch Constituent Snapshots**:
    *   For **Individual Constituents**: Look up `SleeveSnapshots` using `ClashZoneGuid`.
        *   *Critical*: Do NOT rely on `SleeveInstanceId` for this step, as the constituent acts as a data reference via its stable definition (Guid).
    *   For **Cluster Constituents**: Look up `SleeveSnapshots` using `ClusterInstanceId`.
4.  **Aggregate**:
    *   Collect values for each target parameter (e.g., "System Name", "Service Type") from all found snapshots.
    *   Deduplicate values (Case-insensitive).
    *   Join distinct values with a separator (e.g., ", ").
5.  **Apply**:
    *   Write the aggregated string to the Combined Sleeve parameter in Revit.

## Detailed Implementation Steps

### Phase 1: Snapshot Indexing (SleeveSnapshotRepository)
*   [x] **Load Constituents**: Update `LoadSnapshotIndex` to populate `ByCombined` dictionary.
    *   `Dictionary<int, List<SleeveConstituentSnapshotReference>> ByCombined`
    *   Key: `CombinedInstanceId`
    *   Value: List of objects containing `{ Type, ClashZoneGuid, ClusterInstanceId }`.
*   [x] **Index by Guid**: Update `LoadSnapshotIndex` to populate `ByClashZoneGuid` dictionary.
    *   Ensures fast O(1) retrieval of snapshots by GUID.

### Phase 2: Parameter Aggregation (ParameterTransferService)
*   [x] **Update `TransferFromElementsWithSnapshot`**:
    *   Detect if the element is a Combined Sleeve (check `snapshotIndex.TryGetByCombined`).
    *   If yes, call `AggregateCombinedParameters`.
*   [x] **Implement `AggregateCombinedParameters`**:
    *   Input: List of `SleeveConstituentSnapshotReference`.
    *   Loop through constituents:
        *   If `Type == Individual` && `ClashZoneGuid` has value -> `snapshotIndex.TryGetByClashZoneGuid`.
        *   If `Type == Cluster` && `ClusterInstanceId` has value -> `snapshotIndex.TryGetByCluster`.
    *   Collect and Aggregate parameter values.
    *   Return a `Dictionary<string, string>` representing the "virtual" combined snapshot.

### Phase 3: Diagnostics & Fallbacks
*   [ ] **Diagnostic Logging**:
    *   Log how many constituents were found for each combined sleeve.
    *   Log which specific snapshots (GUIDs/IDs) were successfully retrieved.
    *   Log the final aggregated string for key parameters ("Size", "System Name").
*   [x] **Fallback Implementation**:
    *   If a constituent snapshot is missing, skip it gracefully but log a warning.
    *   Ensure at least one constituent contributes data, otherwise log "No Data Source".

## Verification Checklist
1.  **Place Combined Sleeve**: Create a combined sleeve from 2+ pipes/ducts.
2.  **Verify DB**: Confirm `CombinedSleeveConstituents` table has correct rows linking `CombinedInstanceId` to `ClashZoneGuid`s.
3.  **Run Transfer**: Execute Parameter Transfer.
4.  **Check Revit**: Verify the Combined Sleeve parameter (e.g., "Comments" or "System Name") contains the merged text (e.g., "Sanitary, Vent").
5.  **Check Logs**: `transfer_debug.log` should show:
    *   `[PARAM_TRANSFER] ✅ Matched Combined Sleeve {Id}...`
    *   `[AGGREGATE] Processing X constituents...`
    *   `[AGGREGATE] ✅ Found snapshot via ClashZoneGuid: ...`
