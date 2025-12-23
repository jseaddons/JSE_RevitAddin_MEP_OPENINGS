# IMPLEMENATION PLAN: Combined Sleeve Parameter Transfer & Integrity

## Overview
This plan outlines the steps to enable the "Transfer Parameters" feature for Combined Sleeves by aggregating data from their constituent sleeves (Individual and Cluster). It also enforces database integrity for Combined Sleeves using deterministic GUIDs to prevent duplicates.

## Phase 1: Database Integrity (Deterministic GUID)

**Objective:** Prevent duplicate rows in the `CombinedSleeves` table by using a deterministic identifier derived from the constituent elements.

### 1. Model Update
- [x] **CombinedSleeve.cs**: added `public string DeterministicGuid { get; set; }`.

### 2. Database Schema Update
- [ ] **SleeveDbContext.cs**: 
    - Update `CreateTables` to include `DeterministicGuid` column in `CombinedSleeves` table definition.
    - Add migration logic to add the column if it's missing in existing databases.
    - Create a `UNIQUE INDEX` on the `DeterministicGuid` column to enforce uniqueness at the database level.

### 3. Repository Update
- [ ] **CombinedSleeveRepository.cs**:
    - **GUID Generation**: Implement `GenerateDeterministicGuid(List<ClashZone> constituents)`.
        - Logic: Sort constituent ClashZone GUIDs alphabetically, join them with a separator, and hash the result (SHA256) to produce a consistent unique ID for that specific combination of sleeves.
    - **Save Logic**: Update `SaveCombinedSleeve`:
        - Generate the GUID before checking/saving.
        - Check if a row with this GUID already exists (`SELECT CombinedInstanceId FROM CombinedSleeves WHERE DeterministicGuid = @Guid`).
        - **If exists:** Update the existing row (or simply return the existing ID if no update is needed).
        - **If new:** Insert the new row.

## Phase 2: Parameter Transfer Logic

**Objective:** Allow `ParameterTransferService` to handle Combined Sleeves by aggregating parameters from the `SleeveSnapshots` of their constituents.

### 1. Service Update
- [ ] **ParameterTransferService.cs**:
    - Update `ExecuteTransferConfigurationInTransaction` or `TransferFromElementsWithSnapshot`.
    - **Detection**: Identify if a target sleeve is a Combined Sleeve (e.g., check `MEP_Category` param is "Multi-Service" or check internal DB).
    - **Constituent Lookup**:
        - Query `ClashZones` table where `CombinedClusterSleeveInstanceId` matches the Combined Sleeve's Instance ID.
    - **Snapshot Retrieval**:
        - For each constituent ClashZone found:
            - If it's an **Individual Sleeve**, use `SleeveInstanceId` to look up the snapshot from `SleeveSnapshots`.
            - If it's a **Cluster Sleeve**, use `ClusterInstanceId` to look up the snapshot.
    - **Aggregation**:
        - Collecting values: For a requested parameter (e.g., "System Name"), collect values from all constituent snapshots.
        - Processing: Distinct and Join (e.g., "Sanitary, Vent").
    - **Application**:
        - Apply the final aggregated value to the Combined Sleeve element in Revit.

## Phase 3: Verification Strategy

- [ ] **Test Deterministic GUID**: 
    - Attempt to place the same combined sleeve configuration twice.
    - Verify that the number of rows in `CombinedSleeves` does not increase.
- [ ] **Test Parameter Transfer**:
    - Place a Combined Sleeve.
    - Run "Transfer Parameters" (Standard/MEP).
    - detailed verification: Check that parameters like "MEP_System_Name" on the Combined Sleeve contain the comma-separated values from its constituents.
