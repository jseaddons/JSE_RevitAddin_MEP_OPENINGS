# Extending MEP Openings for Linked MEP Models

## 1. Overview

This document outlines the challenges and proposed solutions for extending the JSE_RevitAddin_MEP_OPENINGS tool. The goal is to enhance the existing functionality to support MEP (Mechanical, Electrical, and Plumbing) elements that reside in linked Revit models, in addition to the current capability of processing MEP elements in the active host model.

---

## 2. Current Functionality

- The tool identifies MEP elements (like ducts, pipes, cable trays) within the **active Revit document**.
- It then checks for intersections between these elements and structural or architectural elements within **linked Revit models**.
- Openings (sleeves) are created in the active model at the points of intersection.

---

## 3. Proposed Enhancement: Automated Linked Model Processing

The program will be extended to **automatically** process MEP elements from **all linked models** without requiring user selection. The tool will seamlessly combine elements from the active document and all loaded links into a single collection for intersection analysis.

**Automated Workflow:**
1.  **Consolidate MEP Elements:**
    -   The tool first gathers all relevant MEP elements (ducts, pipes, cable trays) from the **active host document**.
    -   It then iterates through every `RevitLinkInstance` in the project. For each link, it opens the linked document and collects all MEP elements.
    -   This creates a comprehensive in-memory collection of all MEP elements, regardless of their origin (active vs. linked).
2.  **Consolidate Structural/Architectural Elements:**
    -   Similarly, the tool gathers all relevant structural and architectural elements (walls, floors, beams) from the active document and all linked models.
3.  **Run Intersection Checks:**
    -   The tool checks for intersections between the consolidated collection of MEP elements and the consolidated collection of structural/architectural elements. All geometric comparisons will be performed in the host model's coordinate system.
4.  **Create Openings:**
    -   Openings are created in the **active host model** at the intersection locations, with the correct level association as described in the challenges below.

---

## 4. Key Implementation Challenges & Solutions

### Challenge 1: Accessing and Transforming Elements from Linked Models

- **The Challenge:** Elements within a linked document exist in their own coordinate system. To accurately detect intersections between elements from two different links (e.g., a linked MEP model and a linked structural model), their geometry must be compared in a common coordinate system. The host model's coordinate system is the logical choice for this.

- **The Solution:**
    1.  **Get the Linked Document:** Access the linked document via the `RevitLinkInstance.GetLinkDocument()` method.
    2.  **Get the Transform:** Obtain the transformation matrix that positions the linked model within the host model's space using `RevitLinkInstance.GetTotalTransform()`. This transform accounts for the link's origin, rotation, and elevation.
    3.  **Apply the Transform:** Before performing any geometric calculations or intersection checks, apply this transform to the geometry of each element retrieved from the linked model.

    **Example Snippet:**
    ```csharp
    // Get the link instance and its transform
    RevitLinkInstance linkInstance = doc.GetElement(linkId) as RevitLinkInstance;
    Transform transform = linkInstance.GetTotalTransform();

    // Get an element from the linked document
    Document linkDoc = linkInstance.GetLinkDocument();
    Element linkedElement = linkDoc.GetElement(elementId);

    // Get the element's geometry and apply the transform
    Options options = new Options();
    GeometryElement geoElem = linkedElement.get_Geometry(options);
    GeometryElement transformedGeo = geoElem.GetTransformed(transform);

    // Now, use 'transformedGeo' for intersection checks in the host model's coordinate space.
    ```

    **A Note on Shared Coordinates:**
    While using Shared Coordinates is a fundamental BIM best practice for aligning models, the Revit API still requires the use of the `RevitLinkInstance.GetTotalTransform()` method. This transform correctly accounts for the link's position and orientation relative to the host model, regardless of whether Shared Coordinates, Origin-to-Origin, or another method was used. Relying on this programmatic transform makes the tool more robust and ensures it works correctly in all scenarios, without requiring a specific linking methodology from the user.

### Challenge 2: Setting the Correct "Schedule Level" for Openings

- **The Challenge:** When an opening is created for an MEP element from a linked model, its "Schedule Level" parameter must be set correctly. The opening family instance exists in the host model, so it needs to be associated with a level from the host model. As you identified, the most logical approach is to find the nearest level *below* the MEP element's elevation.

- **The Solution:**
    1.  **Get Element Elevation:** After transforming the MEP element's geometry into the host coordinate system, determine its Z-axis elevation. For a linear element like a duct or pipe, the centerline's Z-coordinate is suitable.
    2.  **Collect Host Levels:** Get all `Level` elements from the active host document.
    3.  **Find Nearest Level Below:** Iterate through all host levels and compare their elevations to the MEP element's Z-coordinate. The correct level will be the one with the highest elevation that is still less than or equal to the element's elevation.
    4.  **Set the Parameter:** Once the target level is identified, get its `ElementId` and assign it to the "Schedule Level" parameter of the newly created opening family instance.

    **Example Logic:**
    ```csharp
    // 'mepElementElevationZ' is the calculated elevation of the linked MEP element in the host's coordinates.
    double mepElementElevationZ = transformedLocation.Z;

    Level targetLevel = null;
    double maxElevationBelow = -Double.MaxValue;

    // Get all levels from the HOST document
    FilteredElementCollector levelCollector = new FilteredElementCollector(doc).OfClass(typeof(Level));

    foreach (Level level in levelCollector)
    {
        double levelElevation = level.ProjectElevation;
        if (levelElevation <= mepElementElevationZ && levelElevation > maxElevationBelow)
        {
            maxElevationBelow = levelElevation;
            targetLevel = level;
        }
    }

    // If a valid level was found, set the parameter on the new opening
    if (targetLevel != null)
    {
        Parameter scheduleLevelParam = newOpening.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
        scheduleLevelParam.Set(targetLevel.Id);
    }
    ```

### Challenge 3: Performance

- **The Challenge:** Processing multiple linked documents, each potentially containing thousands of elements, can be computationally expensive and slow. A naive approach of checking every MEP element against every structural/architectural element will not be performant.

- **The Solution:**
    1.  **Pre-filter with Bounding Boxes:** Before performing precise geometric intersection checks, use a `BoundingBoxIntersectsFilter`. This filter is much faster and can quickly eliminate elements that are not in the same general vicinity.
    2.  **Get the Bounding Box:** Get the bounding box of the MEP element from the linked file.
    3.  **Transform the Bounding Box:** Apply the link's `TotalTransform` to the min and max points of the bounding box to correctly position it in the host model's coordinate system.
    4.  **Apply the Filter:** Use this transformed bounding box to filter for potentially intersecting elements in the structural/architectural linked models (whose elements must also be conceptually transformed for the comparison).

---
## 5. Using a Section Box for Element Selection

### Challenge 4: Filtering with a Section Box

- **The Challenge:** If the user activates a section box in a 3D view, the expectation is that the tool will only process elements visible within that box. This applies to elements from both the host model and all linked models. The primary challenge is that the section box exists in the host model's coordinate system, while the linked elements exist in their own.

- **The Mitigation/Solution:**
    1.  **Check for Active Section Box:** First, determine if the active view is a 3D view and if a section box is enabled.
        ```csharp
        View3D view3D = uidoc.ActiveView as View3D;
        if (view3D == null || !view3D.IsSectionBoxActive)
        {
            // No section box, or not a 3D view, so process all elements.
            return; 
        }
        ```
    2.  **Get the Section Box Bounding Box:** Retrieve the bounding box of the section box. This `BoundingBoxXYZ` is defined in the host model's coordinate system.
        ```csharp
        BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
        ```
    3.  **Create an `Outline` and `BoundingBoxIntersectsFilter`:** Use the section box's coordinates to create an `Outline` object, which can then be used to initialize a `BoundingBoxIntersectsFilter`. This filter will be the primary tool for identifying relevant elements.
        ```csharp
        Outline outline = new Outline(sectionBox.Min, sectionBox.Max);
        BoundingBoxIntersectsFilter sectionBoxFilter = new BoundingBoxIntersectsFilter(outline);
        ```
    4.  **Filter Elements in Linked Models:** This is the most critical step. You cannot directly apply the host-coordinate filter to a linked document. Instead, you must transform the *filter itself* into the coordinate system of the link.
        - Get the link instance and its total transform.
        - **Invert the transform.** The inverse transform will map coordinates from the host model back into the linked model's coordinate system.
        - Apply this inverse transform to the `BoundingBoxIntersectsFilter`.
        - Use this newly transformed filter to collect elements from the linked document.

        ```csharp
        // Inside a loop for each MEP link instance
        RevitLinkInstance linkInstance = ...;
        Transform linkTransform = linkInstance.GetTotalTransform();
        
        // Create a filter based on the INVERSE transform of the link
        BoundingBoxIntersectsFilter invertedFilter = new BoundingBoxIntersectsFilter(outline, linkTransform.Inverse);

        // Now collect elements from the LINKED document using the inverted filter
        Document linkDoc = linkInstance.GetLinkDocument();
        FilteredElementCollector collector = new FilteredElementCollector(linkDoc).WherePasses(invertedFilter);
        // ...further filter this collector for ducts, pipes, etc.
        ```
    5.  **Filter Structural/Architectural Linked Models:** Repeat the same process as in step 4 for the structural and architectural links to ensure you are only checking for intersections against elements that are also within the user's section box view.
    6.  **User Interface:** Add a checkbox or option in the user interface, such as "Use active view's section box if available," to make this feature optional. This provides clarity and control to the user.