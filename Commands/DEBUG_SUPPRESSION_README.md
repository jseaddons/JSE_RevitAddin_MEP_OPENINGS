# DEBUG README: Suppression Logic Enhancement for StructuralSleevePlacementCommand.cs

## Objective

Enhance the suppression logic for cable tray, duct, and pipe sleeve placement to ensure sleeves are only suppressed if an existing sleeve is found at the intended placement area, using the following rules:

- **Individual sleeves** (families ending with `OnWall` or `OnSlab`):
  - Suppression requires BOTH checks:
    1. **Center-to-center check:** If an existing individual sleeve is found within 10mm of the intended placement point, placement is suppressed.
    2. **Cluster sleeve overlap check:** If the intended placement point falls within the bounding box (with 10mm tolerance) of any existing cluster (rectangular) sleeve, placement is suppressed.
  - Both checks must be performed for every individual sleeve placement.
- **Cluster (rectangular) sleeves** (families ending with `Rect`):
  - Suppression only checks for the presence of a prior cluster sleeve at the placement point (i.e., if the intended placement point is inside the bounding box (with 10mm tolerance) of any existing cluster sleeve, placement is suppressed).
  - No center-to-center check is performed for cluster sleeves.

## Family Naming Conventions
- Cluster sleeves: family names ending with `Rect` (e.g., `JSE_PipeSleeveOnWallRect`)
- Individual sleeves: family names ending with `OnWall` or `OnSlab` (e.g., `JSE_PipeSleeveOnWall`, `JSE_PipeSleeveOnSlab`)

## Scope
- **Files:**
  - `StructuralSleevePlacementCommand.cs`
  - `PipeOpeningsRectCommand.cs`
  - `RectangularSleeveClusterCommand.cs`
- **Elements:** Cable tray, duct, and pipe sleeves only.
- **Suppression logic:**
  - For each new sleeve placement, determine the family type:
    - If the family name ends with `OnWall` or `OnSlab` (individual sleeve):
      - Perform BOTH suppression checks:
        1. Center-to-center (10mm) against all existing individual sleeves.
        2. Bounding box (10mm tolerance) against all existing cluster sleeves.
      - If either check finds a match, suppress placement.
    - If the family name ends with `Rect` (cluster sleeve):
      - Only perform bounding box (10mm tolerance) check against existing cluster sleeves.
      - If a match is found, suppress placement.
  - If no suppression condition is met, allow placement.

## Debugging Steps
1. **First Run:**
   - No sleeves exist in the model.
   - All valid sleeves should be placed (no suppression occurs).
2. **Second Run (and after):**
   - For each intended placement:
     - If placing an individual sleeve, check for both:
       1. Existing individual sleeve within 10mm (center-to-center).
       2. Existing cluster sleeve whose bounding box (with 10mm tolerance) contains the point.
     - If placing a cluster sleeve, check only for existing cluster sleeve bounding box overlap (10mm tolerance).
   - If any relevant check finds a match, log suppression and skip placement.
   - If not, place the sleeve as normal.

## Key Points
- Suppression logic is determined by the sleeve family name:
  - **OnWall/OnSlab** (individual): BOTH center-to-center (10mm) and cluster bounding box (10mm tolerance) checks are performed.
  - **Rect** (cluster): ONLY cluster bounding box (10mm tolerance) check is performed.
- Applies to all MEP types: cable tray, duct, and pipe.
- The debug log should clearly indicate when suppression occurs, which method(s) were used, and why.

## Example Log Messages
- `SUPPRESSED: Existing individual sleeve found at {point} (within 10mm, center-to-center), skipping placement.`
- `SUPPRESSED: Existing cluster sleeve bounding box contains {point} (within 10mm), skipping placement.`
- `PLACED: No existing sleeve found at {point}, placing new sleeve.`

## Implementation Notes
- Suppression method is selected based on the family name of the sleeve being placed.
- For individual sleeves (`OnWall`/`OnSlab`):
  - Perform both center-to-center (10mm) and cluster bounding box (10mm tolerance) checks.
- For cluster sleeves (`Rect`):
  - Only perform cluster bounding box (10mm tolerance) check.
- The logic should be easy to adjust if the allowance (10mm) needs to be changed.
- All debug and info logs should be clear and focused on suppression events, and indicate which method(s) were used.

---

**This README is for debugging and development purposes only.**

---

# Existing Code Flow in StructuralSleevePlacementCommand.cs

1. **Initialization:**
   - Sets up logging and collects structural elements (floors, framing) from the active and linked documents.
   - Activates required family symbols for sleeve placement.
   - Selects a non-template 3D view for intersection checks.

2. **MEP Element Collection:**
   - Collects all MEP elements (pipes, ducts, cable trays) from the document and linked files.

3. **Intersection Detection:**
   - For each MEP element, determines if it should be processed based on orientation and available structural elements.
   - Checks for geometric intersection with structural elements using solid geometry and ray casting.
   - If an intersection is found, calculates the sleeve placement point.

4. **Suppression Logic:**
   - **Individual sleeves (families ending with `OnWall` or `OnSlab`):**
     - Suppression uses BOTH:
       1. Center-to-center check (10mm) against all existing individual sleeves.
       2. Cluster sleeve bounding box check (10mm tolerance) against all existing cluster sleeves.
     - If either check finds a match, placement is suppressed.
   - **Cluster sleeves (families ending with `Rect`):**
     - Suppression uses ONLY the cluster sleeve bounding box check (10mm tolerance) against all existing cluster sleeves.
     - If a match is found, placement is suppressed.
   - The logic automatically selects the correct suppression method(s) based on the family name of the sleeve being placed.
   - All debug and info logs clearly indicate which suppression method(s) were used and why suppression occurred.

5. **Sleeve Placement:**
   - If not suppressed, places the sleeve at the calculated point.
   - Logs placement and suppression events.

6. **Summary and Logging:**
   - Logs summary of placed and suppressed sleeves, and details of all intersecting elements.

---

# Action Required

- Ensure the suppression logic matches the above rules:
  - For individual sleeves (family name ends with `OnWall` or `OnSlab`): perform BOTH center-to-center and cluster bounding box suppression checks.
  - For cluster sleeves (family name ends with `Rect`): perform ONLY the cluster bounding box suppression check.
- Ensure logs reflect which suppression method(s) were used and why.

---

# Debug Note: Cluster Sleeve Placement and Individual Sleeve Suppression

- Placement of cluster sleeves is working as intended in `Commands/RectangularSleeveClusterCommand.cs`.
- **Current state:**
  - All cluster sleeves are being placed, except for floor pipe cluster sleeves (these are not being placed).
  - All individual sleeves are being placed (no suppression is occurring).
- However, duplicate suppression of individual sleeve placement on top of existing cluster sleeves is **not** happening as expected.
- This means individual sleeves can still be placed inside the bounding box of an existing cluster sleeve, which should be suppressed.
- **Action:**
  - Debug and fix the logic so that individual sleeve placement is properly suppressed when the placement point falls within the bounding box (with 10mm tolerance) of any existing cluster sleeve.
  - Investigate and resolve why floor pipe cluster sleeves are not being placed.
