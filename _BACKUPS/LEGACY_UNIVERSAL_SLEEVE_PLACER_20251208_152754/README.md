# Legacy UniversalSleevePlacerService Backup

**Date:** 2025-12-08 15:27:54

## Status
This service has been moved to backup as it is **LEGACY CODE**. The new implementation is `NewSleevePlacerService.cs`.

## Migration Notes
- **New Service:** `Services/NewSleevePlacerService.cs`
- **Replaced By:** NewSleevePlacerService (refactored with SOLID principles)
- **Reason:** Legacy code replaced with new architecture

## Files in this Backup
- `UniversalSleevePlacerService.cs` - Original legacy service
- `UniversalSleevePlacerService.cs.backup` - Previous backup
- `UniversalSleevePlacerService.cs.backup_before_logging_replace` - Backup before logging changes

## Active References (Need Migration)
⚠️ **IMPORTANT:** The following files still reference UniversalSleevePlacerService and **MUST be updated** to use `NewSleevePlacerService`:
- `Services/SleevePlacementCoordinator.cs` (line 234)
- `Services/Path3InvalidatedPlacementService.cs` (line 254)

These references will cause compilation errors until they are migrated to the new service.

## Do Not Use
⚠️ **DO NOT USE THIS SERVICE** - Use `NewSleevePlacerService` instead.

