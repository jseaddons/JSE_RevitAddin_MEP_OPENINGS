# Structural Sleeve Placement Modification

This document describes the planned modifications for structural sleeve placement logic in the Revit add-in, specifically for cluster sleeve placement and family selection.




## Key Changes

1. **Family Selection for Cluster Sleeves**
   - For walls and structural framings:


   - For floors:


   - The host element (wall, framing, or floor) detected by the individual sleeve command will be used for cluster sleeve placement.

   - For each cluster, the host characteristics (orientation, type) of the sleeves being replaced will be used to select the correct family and placement logic.
   - For walls/framing:
     - If the host is an X-axis wall/framing, use `ClusterOpeningOnWallX`.
     - If the host is a Y-axis wall/framing, use `ClusterOpeningOnWallY`.
   - For floors, always use `ClusterOpeningOnSlab`.


3. **Parameter Assignment**

   - Width will always be set in the X direction, and Height in the Y direction, avoiding the need to swap parameters based on orientation.

- Update the cluster placement logic in `RectangularSleeveClusterCommand.cs` to:
  - Detect host element and orientation for each cluster.

  - Select the appropriate family (`ClusterOpeningOnWallX`, `ClusterOpeningOnWallY`, `ClusterOpeningOnSlab`).


## Duplicate Suppression Logic

- The existing logic for duplicate suppression (preventing placement of overlapping or duplicate cluster sleeves) will remain unchanged.
- Only the family names used for suppression checks will be updated to the new cluster family names (`ClusterOpeningOnWallX`, `ClusterOpeningOnWallY`, `ClusterOpeningOnSlab`).
- No other changes will be made to the suppression algorithm or its behavior.

---

*Document created July 21, 2025 for planned code modifications.*

Implementation Plan
Refactor Cluster Placement Logic

Update RectangularSleeveClusterCommand.cs to group sleeves by host element and orientation (wall/framing/floor, X/Y axis).
For each cluster, detect the host type and orientation using the same logic as the individual sleeve command (reuse or refactor host detection code if needed).
Select the appropriate family for placement:
For X-axis walls/framing: ClusterOpeningOnWallX
For Y-axis walls/framing: ClusterOpeningOnWallY
For floors: ClusterOpeningOnSlab
Replace all usages of families ending with WallRect and SlabRect with the new cluster family names.
Bounding Box and Parameter Assignment

Use the ClusterBoundingBoxServices.cs class to calculate the bounding box for each cluster.
Assign Width to the X direction and Height to the Y direction, based on the bounding box, for all cluster placements.
Remove any logic that swaps width/height based on orientation.
Host Assignment for Cluster Sleeves

When placing a cluster sleeve, assign the detected host element (wall, framing, or floor) to the new family instance, matching the host of the sleeves being replaced.
Ensure the cluster sleeve is placed with the correct orientation and reference level.
Duplicate Suppression

Update suppression checks to use the new cluster family names, but keep the suppression logic unchanged otherwise.
Testing and Validation

Test the updated logic for all host types (walls, framings, floors) and orientations (X/Y axis).
Validate that cluster sleeves are placed with correct parameters and hosts, and that duplicate suppression works as expected.
Helper Classes for Host and Orientation Detection
A new helper class will be created for detecting host wall or framing orientation:
For walls: Orientation will be determined using the wall's normal vector (e.g., X-axis or Y-axis alignment).
For structural framings: Orientation will be determined using the family type and geometry direction.
The helper class will provide methods to:
Identify the host type (wall, framing, or floor) for each sleeve.
Determine the orientation (X or Y axis) for walls and framings.
Return the appropriate family name (ClusterOpeningOnWallX, ClusterOpeningOnWallY, or ClusterOpeningOnSlab) for cluster placement.
This logic will be used in both individual and cluster sleeve placement to ensure consistency.
You can copy and paste this into your file, or let me know if you want me to try again to insert it automatically.

Certainly! Here is the complete logical flow of the code, including all helper/service file logic:

Complete Logical Flow of RectangularSleeveClusterCommandV2.cs and Helpers
Initialization and Logging

The command starts by initializing a timestamped log file and logging assembly/build info.
Sets up the Revit document context.
Collect Individual Sleeves

Uses a FilteredElementCollector to find all placed sleeve family instances in the model whose family names match patterns like "OpeningOnWall" or "OpeningOnSlab".
These are considered individual sleeves to be clustered.
Group by Host and Orientation

For each sleeve, determines its host element (wall, framing, or floor).
Uses HostOrientationHelper.GetHostTypeAndOrientation(host) to:
Identify the host type (wall, framing, or floor).
For walls: Determines orientation (X or Y axis) using the wall's normal vector.
For framings: Determines orientation using the family type and geometry direction.
Groups sleeves by host type, orientation, and host ID.
Cluster Formation (Edge-to-Edge ≤ 100mm)

For each group, forms clusters of sleeves where every sleeve in a cluster is within 100mm edge-to-edge of at least one other sleeve in the same cluster.
Uses a breadth-first search (BFS) approach:
For each unprocessed sleeve, starts a new cluster.
Adds all sleeves within 100mm edge-to-edge to the cluster, marking them as processed.
Edge-to-edge distance is calculated using sleeve center points and their Width/Height parameters.
Begin Transaction

Starts a Revit transaction to batch all changes.
Process Each Cluster

For each cluster in each group:
Skips clusters with only one sleeve (no clustering needed).
Determines the correct cluster family name to use based on host type and orientation:
For X-axis walls/framing: ClusterOpeningOnWallX
For Y-axis walls/framing: ClusterOpeningOnWallY
For floors: ClusterOpeningOnSlab
Finds the corresponding cluster family symbol in the project.
Uses ClusterBoundingBoxServices.GetClusterBoundingBox(cluster) to:
Calculate the bounding box dimensions (Width=X, Height=Y, Depth) and midpoint for the cluster.
Uses HostLevelHelper.GetHostReferenceLevel(doc, cluster[0]) to:
Get the reference level for placement.
Checks for duplicate cluster openings of the same family within 100mm of the intended placement.
If no duplicate exists:
Places a new cluster sleeve at the calculated midpoint.
Sets its Width (X direction), Height (Y direction), and Depth parameters.
Deletes all original sleeves in the cluster.
Commit Transaction and Log Summary

Commits the transaction and logs the number of cluster openings placed and sleeves deleted.
Helper/Service File Logical Flows
HostOrientationHelper.cs
GetHostTypeAndOrientation(Element host):
If host is a Wall, returns "Wall" and determines orientation (X or Y) using the wall's normal vector.
If host is a FamilyInstance with StructuralType.Beam, returns "Framing" and determines orientation using the geometry direction of the framing.
If host is a Floor, returns "Floor".
Otherwise, returns "Unknown".
ClusterBoundingBoxServices.cs
GetClusterBoundingBox(List<FamilyInstance> cluster):
For each sleeve in the cluster, gets its location and Width/Height/Depth parameters.
Calculates the min/max extents in X, Y, and Z directions.
Returns the bounding box dimensions (Width, Height, Depth) and the midpoint.
HostLevelHelper.cs
GetHostReferenceLevel(Document doc, FamilyInstance sleeve):
Determines the reference level for the sleeve's host element.
Used to ensure the new cluster sleeve is placed at the correct level.
This covers every logical step and helper/service class used in the clustering and placement process.