-- Wall sleeves with corners saved in ClashZones (baseline before/after commenting out Regenerate).
-- Run against your SQLite DB (e.g. in Logs\R2023 or project Data path).
-- Use: sqlite3 your_db.sqlite < query_wall_sleeve_corners.sql   or run in DB browser.

SELECT
  ClashZoneId,
  ClashZoneGuid,
  StructuralType,
  SleeveInstanceId,
  SleeveCorner1X AS C1X, SleeveCorner1Y AS C1Y, SleeveCorner1Z AS C1Z,
  SleeveCorner2X AS C2X, SleeveCorner2Y AS C2Y, SleeveCorner2Z AS C2Z,
  SleeveCorner3X AS C3X, SleeveCorner3Y AS C3Y, SleeveCorner3Z AS C3Z,
  SleeveCorner4X AS C4X, SleeveCorner4Y AS C4Y, SleeveCorner4Z AS C4Z
FROM ClashZones
WHERE (StructuralType = 'Wall' OR StructuralType = 'Walls')
  AND SleeveInstanceId > 0
  AND (SleeveCorner1X IS NOT NULL OR SleeveCorner2X IS NOT NULL)
ORDER BY ClashZoneId;

-- Count of wall sleeves with at least one corner set:
-- SELECT COUNT(*) FROM ClashZones
-- WHERE (StructuralType = 'Wall' OR StructuralType = 'Walls') AND SleeveInstanceId > 0
--   AND (SleeveCorner1X IS NOT NULL OR SleeveCorner2X IS NOT NULL);
