# Known Bugs and Issues

## 1. Pipe Rectangular Sleeve Clustering/Replacement Not Working
- The clustering and replacement logic for pipe rectangular sleeves sometimes fails to cluster or replace sleeves as expected.
- Possible cause: The filter or clustering logic may not be correctly identifying all relevant sleeves, or suppression logic may be too aggressive.

## 2. Suppression Logic Based on Bounding Box
- Suppression of new sleeve placement is based on whether the intersection point is inside the bounding box of an existing cluster sleeve, not just proximity.
- This can cause sleeves to be suppressed even if they are not close enough in the center-to-center sense.

## 3. Pipe Command Processes Only Pipe Sleeves
- The `PipeOpeningsRectCommand` is correctly filtering for pipe rectangular sleeve families only.
- If cable tray sleeves are needed, they must be handled in a separate command.

## 4. Lack of Logging/Diagnostics in Some Commands
- Some commands (e.g., cluster replacement) lack sufficient debug logging, making it hard to diagnose why sleeves are not being clustered or replaced.
- More detailed logging is needed for bounding box checks, proximity, and suppression events.

## 5. Cable Tray Rectangular Sleeve Clustering/Replacement Not Implemented
- There is no dedicated command for clustering and replacing cable tray rectangular sleeves (e.g., `CableTrayOpeningOnWall`, `CableTrayOpeningOnSlab`).
- As a result, cable tray sleeves are not being clustered or replaced as expected.
- Solution: Implement a new command (e.g., `RectangularSleeveClusterCommand`) to process only cable tray rectangular sleeve families, using similar logic as the pipe command.
