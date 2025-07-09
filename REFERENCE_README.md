## MEP/Structural Coordination: Cropped View and Linked Model Processing

### MEP Element Processing
- Only MEP elements (pipes, ducts, cable trays, etc.) visible in the current active view are processed.
- This includes MEP elements from:
  - The active model (host document)
  - Any visible linked MEP models (linked Revit files that are visible in the view and not hidden)
- Visibility is determined by the current view’s crop box, view filters, and link visibility.

### Intersection Checking
- For each visible MEP element (from active or visible linked MEP models), intersection is checked against all relevant elements from linked structural and architectural models.
- Structural/architectural elements are always in linked models.
- Only elements from linked models that are visible in the current view (i.e., the link instance is not hidden or turned off in the view) are considered.

### Implementation Helper
- Use `ViewVisibilityHelper.GetAllVisibleMEPElementsIncludingLinks(doc, view)` to collect all visible MEP elements in the current view, from both the active and visible linked MEP models.
- For intersection, iterate over all visible structural/architectural links in the view, and process their elements for intersection with the visible MEP elements.

This approach ensures that only what the user can see (and intends to coordinate) is processed, maximizing performance and reliability.
