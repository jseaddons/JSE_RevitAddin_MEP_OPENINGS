# Floor Sleeve Parameter Bug Analysis

## Problem
Individual floor sleeves in Revit model show:
- `Sleeve Instance ID = -1` (WRONG - should be the actual sleeve ID like 879790)
- `Cluster Sleeve Instance ID = <some_value>` (WRONG - should be -1 for individual sleeves)

Wall sleeves work correctly:
- `Sleeve Instance ID = 880327` (CORRECT)
- `Cluster Sleeve Instance ID = -1` (CORRECT)

## Evidence from Logs
From `placement_debug.log`:
- Floor sleeve 879790: `[SLEEVE_METADATA] Sleeve 879790: Filter='Ventilation_ducts.xml', InstanceID=879790, ClusterInstanceID=-1 ✓`
- This shows parameters are being SET correctly during placement

## Hypothesis
The parameters are being set correctly during individual sleeve placement, but something is overwriting them afterwards. Possible causes:

1. **Different parameter behavior in floor families** - The floor families (RectangularOpeningOnSlab, CircularOpeningOnSlab) might have parameters that behave differently than wall families
2. **Parameter overwrite after placement** - Some process after individual placement is updating the parameters
3. **Family instance type vs instance parameters** - The parameters might be type parameters instead of instance parameters on floor families
4. **Transaction rollback/regeneration** - Document regeneration might be resetting parameters to default values

## Next Steps
1. Check if floor family parameters are instance vs type parameters
2. Check if there's any logic that updates parameters after SetSleeveMetadata
3. Check if document regeneration is affecting floor sleeves differently than wall sleeves
4. Add diagnostic logging to track when parameters change value


