# MEP Openings - Pseudocode Implementation Plan

This document provides a detailed pseudocode implementation plan for the enhanced MEP Openings tool. It covers the automated discovery of elements in linked models, handling of coordinate transformations, filtering with a section box, and correct level assignment.

---

### Data Structures

First, we define a simple structure to hold an element and its corresponding transformation, which will simplify passing data between functions.

```
// A helper class to bundle an element with the transform needed to bring it into the host model's coordinate system.
class TransformedElement {
    Element Element;
    Transform Transform; // For linked elements, this is the link's total transform. For host elements, it's the Identity transform.
}
```

---

### Main Execution Flow

This is the primary function that orchestrates the entire process.

```
function Execute(commandData) {
    // 1. SETUP
    uiDoc = commandData.Application.ActiveUIDocument;
    doc = uiDoc.Document;
    activeView = uiDoc.ActiveView;

    // 2. HANDLE SECTION BOX FILTERING
    // This filter will be used to select elements. It might be null if no section box is active.
    BoundingBoxIntersectsFilter sectionBoxFilter = GetSectionBoxFilter(activeView);

    // 3. GATHER ALL ELEMENTS
    // Get all MEP and Structural elements from the host and all linked models, already filtered by the section box if applicable.
    List<TransformedElement> allMepElements = GetAllMepElements(doc, sectionBoxFilter);
    List<TransformedElement> allStructuralElements = GetAllStructuralElements(doc, sectionBoxFilter);

    // 4. PROCESS INTERSECTIONS
    // Use a transaction to make all changes to the Revit model.
    using (Transaction tx = new Transaction(doc, "Place MEP Openings")) {
        tx.Start();

        // Loop through every MEP element against every structural element.
        foreach (TransformedElement mepElement in allMepElements) {
            foreach (TransformedElement structuralElement in allStructuralElements) {

                // 5. CHECK FOR INTERSECTION
                // This function handles the geometric transformation and returns the intersection solid.
                Solid intersectionSolid = CheckForIntersection(mepElement, structuralElement);

                if (intersectionSolid != null && intersectionSolid.Volume > 0) {
                    // 6. CREATE OPENING
                    // If an intersection exists, create the opening family instance.
                    CreateOpening(doc, intersectionSolid, mepElement);
                }
            }
        }

        tx.Commit();
    }
}
```

---

### Helper Functions

These functions break down the main logic into manageable pieces.

```csharp
//----------------------------------------------------------------------------------
// Gathers all MEP elements (Ducts, Pipes, Cable Trays) from the host and linked models.
//----------------------------------------------------------------------------------
function GetAllMepElements(hostDoc, sectionBoxFilter) {
    List<TransformedElement> elements = new List<TransformedElement>();
    Transform identityTransform = Transform.Identity;

    // A. Get elements from the HOST document
    // Apply the section box filter directly if it exists.
    FilteredElementCollector hostCollector = new FilteredElementCollector(hostDoc);
    // ... add category filters for Ducts, Pipes, CableTrays ...
    if (sectionBoxFilter != null) {
        hostCollector.WherePasses(sectionBoxFilter);
    }
    foreach (Element elem in hostCollector) {
        elements.Add(new TransformedElement(elem, identityTransform));
    }

    // B. Get elements from LINKED documents
    // Get all link instances in the host document.
    FilteredElementCollector linkCollector = new FilteredElementCollector(hostDoc).OfClass(typeof(RevitLinkInstance));

    foreach (RevitLinkInstance linkInstance in linkCollector) {
        Document linkDoc = linkInstance.GetLinkDocument();
        if (linkDoc == null) continue; // Skip unloaded links

        Transform linkTransform = linkInstance.GetTotalTransform();
        BoundingBoxIntersectsFilter linkFilter = sectionBoxFilter;

        // If there's a section box, we must transform the filter into the link's coordinate system.
        if (sectionBoxFilter != null) {
            linkFilter = new BoundingBoxIntersectsFilter(sectionBoxFilter.GetOutline(), linkTransform.Inverse);
        }

        FilteredElementCollector linkedElementCollector = new FilteredElementCollector(linkDoc);
        // ... add category filters for Ducts, Pipes, CableTrays ...
        if (linkFilter != null) {
            linkedElementCollector.WherePasses(linkFilter);
        }

        foreach (Element elem in linkedElementCollector) {
            elements.Add(new TransformedElement(elem, linkTransform));
        }
    }

    return elements;
}

//----------------------------------------------------------------------------------
// Gathers all Structural elements (Walls, Floors, Beams). This function is nearly
// identical to GetAllMepElements but targets structural categories.
//----------------------------------------------------------------------------------
function GetAllStructuralElements(hostDoc, sectionBoxFilter) {
    // ... Logic is the same as GetAllMepElements, but with different category filters
    // (Walls, Floors, Structural Framing, etc.).
    return structuralElements;
}

//----------------------------------------------------------------------------------
// Checks for intersection between a single MEP and a single structural element.
//----------------------------------------------------------------------------------
function CheckForIntersection(mep, structural) {
    // Get geometry and apply the necessary transform to bring it into the host coordinate system.
    GeometryElement mepGeom = mep.Element.get_Geometry(new Options()).GetTransformed(mep.Transform);
    GeometryElement structGeom = structural.Element.get_Geometry(new Options()).GetTransformed(structural.Transform);

    // Find the first solid in each geometry element for the check.
    Solid mepSolid = mepGeom.OfType<Solid>().FirstOrDefault(s => s.Volume > 0);
    Solid structSolid = structGeom.OfType<Solid>().FirstOrDefault(s => s.Volume > 0);

    if (mepSolid == null || structSolid == null) {
        return null;
    }

    // Perform the boolean operation to find the intersection volume.
    Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(mepSolid, structSolid, BooleanOperationsType.Intersect);

    return intersection;
}

//----------------------------------------------------------------------------------
// Creates the opening family instance in the host model.
//----------------------------------------------------------------------------------
function CreateOpening(hostDoc, intersectionSolid, mepElement) {
    // 1. Find insertion point
    XYZ insertionPoint = intersectionSolid.ComputeCentroid();

    // 2. Find the correct level for the opening
    Level level = FindNearestLevelBelow(hostDoc, insertionPoint.Z);
    if (level == null) {
        // Handle error: no suitable level found below the element.
        return;
    }

    // 3. Load and activate the opening family symbol if needed
    FamilySymbol openingSymbol = GetOpeningFamilySymbol(hostDoc);
    if (!openingSymbol.IsActive) {
        openingSymbol.Activate();
    }

    // 4. Create the family instance
    FamilyInstance opening = hostDoc.Create.NewFamilyInstance(insertionPoint, openingSymbol, level, StructuralType.NonStructural);

    // 5. Set parameters (size, etc.) on the new opening based on the intersection solid's bounding box and the MEP element's properties.
    // ... e.g., opening.LookupParameter("Width").Set(intersectionSolid.BoundingBox.Max.X - intersectionSolid.BoundingBox.Min.X);
    // ... e.g., opening.LookupParameter("Height").Set(intersectionSolid.BoundingBox.Max.Y - intersectionSolid.BoundingBox.Min.Y);
}

//----------------------------------------------------------------------------------
// Finds the nearest level in the host document that is below a given elevation.
//----------------------------------------------------------------------------------
function FindNearestLevelBelow(hostDoc, elementElevation) {
    Level bestLevel = null;
    double maxElevationBelow = -Double.MaxValue;

    FilteredElementCollector levelCollector = new FilteredElementCollector(hostDoc).OfClass(typeof(Level));

    foreach (Level level in levelCollector) {
        if (level.ProjectElevation <= elementElevation && level.ProjectElevation > maxElevationBelow) {
            maxElevationBelow = level.ProjectElevation;
            bestLevel = level;
        }
    }
    return bestLevel;
}

//----------------------------------------------------------------------------------
// Gets the BoundingBoxIntersectsFilter if a section box is active in a 3D view.
//----------------------------------------------------------------------------------
function GetSectionBoxFilter(activeView) {
    View3D view3D = activeView as View3D;
    if (view3D != null && view3D.IsSectionBoxActive) {
        BoundingBoxXYZ sectionBox = view3D.GetSectionBox();
        Outline outline = new Outline(sectionBox.Min, sectionBox.Max);
        return new BoundingBoxIntersectsFilter(outline);
    }
    return null; // Return null if not a 3D view or no section box is active.
}

```
