# How to Check Logs for Flag Reset Issue

## Log File Location
**`%AppData%\Roaming\JSE_MEP_Openings\Logs\Refresh_YYYY-MM-DD_HH-mm-ss.log`**

## What to Check in Logs

### 1. Flag Reset Confirmation
Look for these log entries:
```
[FLAG-MANAGER] ===== CHECKING FOR DELETED SLEEVES (after flag sync) =====
[FLAG-MANAGER] Reset X clash zones in category 'CategoryName' - sleeves deleted
[FLAG-MANAGER] ✅ Total X clash zones reset due to deleted sleeves - these zones are now unresolved
```

**Expected:** Should show that flags were reset after sleeve deletion.

### 2. Flag Sync After Reset
Look for:
```
[REFRESH-GLOBAL-SYNC-AFTER-MERGE] Synced flags for X clash zones in category 'CategoryName' AFTER merging new zones
```

**Expected:** Should sync flags from Global XML after reset.

### 3. Unresolved Count Check
Look for:
```
[CLASH_DEBUG] Using enabledFilter.ClashZoneStorage for unresolved count check: X zones
[CLASH_DEBUG] Flag check: X total, Y unresolved (after reset), Z cluster resolved, W individual resolved
```

**Expected:** After reset, `unresolved` count should be > 0 if sleeves were deleted.

### 4. "No New Zones" Message
Look for:
```
[CLASH_DEBUG] ✅ All X existing clash zones are resolved (Y cluster, Z individual) - no zones need processing
```

**Problem:** If this appears even after flags were reset, it means `unresolvedCount` is still 0.

## What the Fix Does

The fix now:
1. Uses `enabledFilter.ClashZoneStorage.ClashZones` (which contains updated flags after save)
2. Falls back to `existingClashZones.ClashZones` if enabledFilter is not available
3. Logs which source is being used for unresolved count check
4. Logs the actual flag counts after reset

## Next Steps

1. Run refresh after deleting a sleeve
2. Check the log file for the entries above
3. Look for:
   - Confirmation that flags were reset
   - The unresolved count should be > 0 after reset
   - If unresolved count is still 0, check if flags are being synced back from Global XML incorrectly

