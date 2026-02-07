-- Check sleeve corners in DB (for combined-sleeve constituents).
-- Run against your SQLite DB (e.g. Logs\R2023 or app Data path).
-- Replace 1506596 / 1506872 with your constituent SleeveInstanceId / ClusterInstanceId.

-- =============================================================================
-- 1. INDIVIDUAL SLEEVE CORNERS (ClashZones) — by SleeveInstanceId
-- =============================================================================
SELECT
  'Individual' AS SleeveType,
  ClashZoneId,
  ClashZoneGuid,
  SleeveInstanceId,
  StructuralType,
  SleeveCorner1X AS C1X, SleeveCorner1Y AS C1Y, SleeveCorner1Z AS C1Z,
  SleeveCorner2X AS C2X, SleeveCorner2Y AS C2Y, SleeveCorner2Z AS C2Z,
  SleeveCorner3X AS C3X, SleeveCorner3Y AS C3Y, SleeveCorner3Z AS C3Z,
  SleeveCorner4X AS C4X, SleeveCorner4Y AS C4Y, SleeveCorner4Z AS C4Z
FROM ClashZones
WHERE SleeveInstanceId = 1506596;

-- By ClashZoneGuid (e.g. for combined sleeve 1507026 individual constituent):
-- WHERE ClashZoneGuid = '419f5f8d-1196-7af0-7bf3-376aa0716ea6';

-- =============================================================================
-- 2. CLUSTER SLEEVE CORNERS (ClusterSleeves) — by ClusterInstanceId
-- =============================================================================
SELECT
  'Cluster' AS SleeveType,
  ClusterSleeveId,
  ClusterInstanceId,
  HostType,
  Corner1X AS C1X, Corner1Y AS C1Y, Corner1Z AS C1Z,
  Corner2X AS C2X, Corner2Y AS C2Y, Corner2Z AS C2Z,
  Corner3X AS C3X, Corner3Y AS C3Y, Corner3Z AS C3Z,
  Corner4X AS C4X, Corner4Y AS C4Y, Corner4Z AS C4Z
FROM ClusterSleeves
WHERE ClusterInstanceId = 1506872;

-- =============================================================================
-- 3. CLUSTER SLEEVE CORNERS (ClusterSleeves_v2) — if corners stored in v2
-- =============================================================================
SELECT
  'Cluster_v2' AS SleeveType,
  ClusterGUID,
  ClusterInstanceId,
  HostType,
  Corner1X AS C1X, Corner1Y AS C1Y, Corner1Z AS C1Z,
  Corner2X AS C2X, Corner2Y AS C2Y, Corner2Z AS C2Z,
  Corner3X AS C3X, Corner3Y AS C3Y, Corner3Z AS C3Z,
  Corner4X AS C4X, Corner4Y AS C4Y, Corner4Z AS C4Z
FROM ClusterSleeves_v2
WHERE ClusterInstanceId = 1506872;

-- =============================================================================
-- 4. QUICK CHECK: any corners set? (0 or NULL = not saved)
-- =============================================================================
-- Individual (SleeveInstanceId 1506596):
SELECT SleeveInstanceId,
  (SleeveCorner1X IS NOT NULL AND (SleeveCorner1X != 0 OR SleeveCorner1Y != 0 OR SleeveCorner1Z != 0)) AS C1_set,
  (SleeveCorner2X IS NOT NULL AND (SleeveCorner2X != 0 OR SleeveCorner2Y != 0 OR SleeveCorner2Z != 0)) AS C2_set,
  (SleeveCorner3X IS NOT NULL AND (SleeveCorner3X != 0 OR SleeveCorner3Y != 0 OR SleeveCorner3Z != 0)) AS C3_set,
  (SleeveCorner4X IS NOT NULL AND (SleeveCorner4X != 0 OR SleeveCorner4Y != 0 OR SleeveCorner4Z != 0)) AS C4_set
FROM ClashZones WHERE SleeveInstanceId = 1506596;

-- Cluster (ClusterInstanceId 1506872):
SELECT ClusterInstanceId,
  (Corner1X != 0 OR Corner1Y != 0 OR Corner1Z != 0) AS C1_set,
  (Corner2X != 0 OR Corner2Y != 0 OR Corner2Z != 0) AS C2_set,
  (Corner3X != 0 OR Corner3Y != 0 OR Corner3Z != 0) AS C3_set,
  (Corner4X != 0 OR Corner4Y != 0 OR Corner4Z != 0) AS C4_set
FROM ClusterSleeves WHERE ClusterInstanceId = 1506872;
