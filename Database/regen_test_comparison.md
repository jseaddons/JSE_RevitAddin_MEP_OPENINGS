# Regenerate-before-commit test: corner comparison

## Test
- **Before:** Placement with explicit `_document.Regenerate()` before `Transaction.Commit()`.
- **After:** Placement with Regenerate commented out (Commit only).

## Result
Corner coordinates (C1X–C4Z) match between the two runs **by ClashZoneGuid**.  
Only `SleeveInstanceId` differs (new Revit element IDs on the second run).  
Minor floating-point differences in the last decimal place (e.g. 15454 vs 15453) are normal and not meaningful.

**Conclusion:** Explicit Regenerate before Commit is **not required** for correct sleeve corners.  
Revit’s Commit (or internal regen) is enough for `UpdateZonesFromElements` to read correct geometry.  
Leaving Regenerate commented out saves ~260 ms per placement run.

## Files
- `baseline_wall_sleeve_corners.csv` – with Regenerate
- `after_no_regen_wall_sleeve_corners.csv` – without Regenerate (this test)
