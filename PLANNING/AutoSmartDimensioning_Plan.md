# Auto / Smart Dimensioning for Builders Work Openings
## JSE MEP Opening Addin — Detailed Implementation Plan

---

## 1. WHAT THIS FEATURE IS

**Auto Dimensioning** adds Revit `Dimension` annotation elements automatically to builders work
openings (sleeves) after they are placed, showing:

- Horizontal / vertical distance from the nearest **Grid line**
- Distance from the **host wall face** or **slab face**
- Distance from **Reference Planes** (structural datum lines)
- **Opening size** (width x height, or diameter) as annotated dimension strings
- **Elevation / Bottom of Opening (BOO)** dimension from finished floor level (FFL)

This is **separate** from the existing `DimensionStage.cs` which computes the clearance-based
SIZE of the sleeve in the pipeline. This new feature creates *Revit dimension objects visible
on drawings*.

---

## 2. SCOPE — PHASE 1 (Builders Work / Sleeves only)

| Category | Phase 1 | Future |
|---|---|---|
| Circular Pipe Sleeves | YES | - |
| Rectangular Duct Sleeves | YES | - |
| Cable Tray Sleeves | YES | - |
| Cluster Openings | YES | - |
| Fire Damper Frames | - | Phase 2 |
| Generic Structural Openings | - | Phase 2 |
| All Other Categories | - | Phase 3 |

---

## 3. RESEARCH FINDINGS — INDUSTRY BEST PRACTICES

### 3.1 How Leading Tools Do It (from research)

| Tool | Strategy |
|---|---|
| **AGACAD Cut Opening** | Creates dimensions in-view automatically post-placement; dimensions to nearest structural grid + wall face |
| **ConVoid** | One-click wall dimensioning with MEP elements; supports linked model references |
| **Naviate Auto Dimension** | Configurable side placement + offsets; supports off-grid openings |
| **Auto Dimension Pack** | Batch dimensions per view; configurable reference targets (grids, walls, levels) |
| **THBIM Add-ins** | X/Y coordinate-based dimensions; aligned clean string output |

### 3.2 Revit API Core Mechanism

```csharp
// Step 1: Build reference array of targets
ReferenceArray refs = new ReferenceArray();
refs.Append(openingCenterRef);       // opening reference plane
refs.Append(gridRef);                // nearest grid
refs.Append(wallFaceRef);            // host wall face

// Step 2: Define dimension line direction
Line dimLine = Line.CreateUnbound(openingCenter, XYZ.BasisX); // horizontal

// Step 3: Create the dimension annotation
Dimension dim = doc.Create.NewDimension(view, dimLine, refs);
```

### 3.3 Reference Types Available

- `Grid.Curve.GetEndPointReference()` — grid line references
- `Wall.GetReferences(FaceReferenceType)` — wall face center / finish
- `ReferencePlane.GetReference()` — reference plane
- `FamilyInstance` geometry — opening center reference planes (built into family)
- `Level.GetPlaneReference()` — for elevation dimensions

---

## 4. ARCHITECTURE DESIGN

### 4.1 New Folder Structure

```
Services/
  Dimensioning/
    IAutoDimensioningService.cs          (interface)
    AutoDimensioningService.cs           (main orchestrator)
    DimensionReferenceResolver.cs        (resolves nearest grid/wall refs)
    DimensionLineCalculator.cs           (computes line position and direction)
    OpeningSizeDimensionWriter.cs        (writes W x H or diameter dims)
    OpeningLocationDimensionWriter.cs    (writes X/Y from grid/wall dims)
    ElevationDimensionWriter.cs          (writes BOO from level dim)
    DimensionViewSelector.cs             (picks correct view for annotation)
    DimensionTypeResolver.cs             (resolves DimensionType from name)
  Interfaces/
    IDimensionReferenceResolver.cs
    IDimensionWriter.cs
Models/
  Dimensioning/
    DimensionRequest.cs                  (input: sleeve + view + options)
    DimensionResult.cs                   (output: ElementIds of created dims)
    DimensionTarget.cs                   (enum: Grid, WallFace, ReferencePlane, Level)
    DimensionOptions.cs                  (user configuration for the feature)
Commands/
  AutoDimensionCommand.cs               (IExternalCommand entry point)
Views/
  AutoDimensionSettingsDialog.xaml      (WPF UI for settings)
  AutoDimensionSettingsDialog.xaml.cs
ViewModels/
  AutoDimensionSettingsViewModel.cs
```

### 4.2 Key Interfaces

```csharp
// IAutoDimensioningService.cs
public interface IAutoDimensioningService
{
    DimensionResult DimensionSleeve(Document doc, FamilyInstance sleeve,
        View view, DimensionOptions options);

    IReadOnlyList<DimensionResult> DimensionAllSleevesInView(Document doc,
        View view, DimensionOptions options);

    IReadOnlyList<DimensionResult> DimensionSelectedSleeves(Document doc,
        IReadOnlyList<ElementId> sleeveIds, View view, DimensionOptions options);
}

// IDimensionReferenceResolver.cs
public interface IDimensionReferenceResolver
{
    Reference ResolveNearestGrid(Document doc, XYZ point, XYZ direction);
    Reference ResolveHostWallFace(Document doc, FamilyInstance sleeve);
    Reference ResolveNearestReferencePlane(Document doc, XYZ point);
    Reference ResolveLevelReference(Document doc, Level level);
}

// IDimensionWriter.cs
public interface IDimensionWriter
{
    ElementId WriteDimension(Document doc, View view,
        ReferenceArray refs, Line dimensionLine, string dimensionTypeName);
}
```

### 4.3 DimensionOptions Model

```csharp
public class DimensionOptions
{
    // Reference targets
    public bool DimensionToGrid { get; set; } = true;
    public bool DimensionToWallFace { get; set; } = true;
    public bool DimensionToReferencePlane { get; set; } = false;

    // What to annotate
    public bool AnnotateSize { get; set; } = true;          // W x H / diameter
    public bool AnnotateLocation { get; set; } = true;      // X/Y from reference
    public bool AnnotateElevation { get; set; } = true;     // BOO from FFL

    // Offset from element for the dimension line (mm, converted to feet internally)
    public double DimensionLineOffsetMm { get; set; } = 500.0;

    // Name of DimensionType to use (null = project default)
    public string DimensionTypeName { get; set; } = null;

    // Category filter (Phase 1: sleeves only; future: all categories)
    public IReadOnlyList<string> CategoryFilter { get; set; }
        = new[] { "Pipe Sleeves", "Rectangular Sleeves", "Cable Tray Sleeves" };

    // Grid search radius (mm)
    public double GridSearchRadiusMm { get; set; } = 15000.0;
}
```

---

## 5. IMPLEMENTATION PHASES

### Phase 1A — Core Infrastructure (no UI)
**Goal:** Service layer that can dimension a single sleeve in a given view.

**Files to create:**
1. `Models/Dimensioning/DimensionOptions.cs`
2. `Models/Dimensioning/DimensionRequest.cs`
3. `Models/Dimensioning/DimensionResult.cs`
4. `Models/Dimensioning/DimensionTarget.cs`
5. `Services/Dimensioning/IDimensionReferenceResolver.cs` (interface)
6. `Services/Dimensioning/IDimensionWriter.cs` (interface)
7. `Services/Dimensioning/IAutoDimensioningService.cs` (interface)
8. `Services/Dimensioning/DimensionTypeResolver.cs`
9. `Services/Dimensioning/DimensionReferenceResolver.cs`
10. `Services/Dimensioning/DimensionLineCalculator.cs`
11. `Services/Dimensioning/OpeningSizeDimensionWriter.cs`
12. `Services/Dimensioning/OpeningLocationDimensionWriter.cs`
13. `Services/Dimensioning/ElevationDimensionWriter.cs`
14. `Services/Dimensioning/AutoDimensioningService.cs`

**Phase 1A Deliverable:** Unit-testable services, no Revit transaction dependency (transaction managed by caller).

---

### Phase 1B — View Selection + Command
**Goal:** User can trigger auto-dimensioning from the ribbon.

**Files to create:**
15. `Services/Dimensioning/DimensionViewSelector.cs`
16. `Commands/AutoDimensionCommand.cs`

**Logic in DimensionViewSelector:**
- Find Section/Elevation views that cut through the sleeve bounding box
- Prefer views at scale 1:50 or 1:20 (configurable)
- Fall back to active view if no section found
- Support "dimension in all relevant views" mode

---

### Phase 1C — WPF Settings Dialog
**Goal:** User can configure dimension options before running.

**Files to create:**
17. `Views/AutoDimensionSettingsDialog.xaml`
18. `Views/AutoDimensionSettingsDialog.xaml.cs`
19. `ViewModels/AutoDimensionSettingsViewModel.cs`

**UI Controls:**
- CheckBoxes for: Dimension to Grid, Dimension to Wall Face, Annotate Size, Annotate Elevation
- TextBox: Dimension line offset (mm)
- ComboBox: Dimension type selection (populated from project)
- RadioButtons: Current view / All section views / Selected elements only
- Preview count: "X sleeves will be dimensioned"

---

### Phase 1D — Integration with Placement Pipeline (optional post-placement)
**Goal:** Auto-dimension sleeves immediately after bulk placement.

**Changes to existing files:**
- `Services/Placement/SleevePlacementOrchestrator.cs` — add optional post-pipeline step
- `Models/GlobalApplicationSettings.cs` — add `AutoDimensionAfterPlacement` flag

**Integration pattern:**
```csharp
// In SleevePlacementOrchestrator.Execute() AFTER batch flush:
if (_dimensioningService != null && _config.AutoDimensionAfterPlacement)
{
    var placedIds = currentContext.PlacedInstances.Select(fi => fi.Id).ToList();
    var activeView = doc.ActiveView;
    _dimensioningService.DimensionSelectedSleeves(doc, placedIds, activeView, _dimOptions);
}
```

---

### Phase 2 — Extended Category Support
**Goal:** Apply same logic to Fire Dampers, structural openings.

**Changes:**
- `DimensionOptions.CategoryFilter` expanded
- New `IDimensionReferenceResolver` implementations per category if reference extraction differs
- Category-specific clearance offsets for dimension line position

---

### Phase 3 — All Categories
**Goal:** User selects any Revit category; tool auto-dimensions all instances in view.

**Changes:**
- `DimensionViewSelector` supports all category element bounding box intersection
- Generic reference extraction using `GeometryElement` iteration
- `AutoDimensionSettingsDialog` expanded with category picker

---

## 6. DETAILED SERVICE LOGIC

### 6.1 DimensionReferenceResolver — Nearest Grid

```
Algorithm:
1. Collect all Grid elements in document (FilteredElementCollector)
2. For each grid, compute distance from sleeve XY center to grid curve
3. Filter grids within GridSearchRadiusMm
4. Separate into: grids roughly parallel to X-axis (horizontal) vs Y-axis (vertical)
5. Return nearest horizontal grid + nearest vertical grid
6. Extract reference using grid.Curve.GetEndPointReference() or grid face reference
```

**Key consideration:** Grids from linked structural models require transform of the link
instance. Use `RevitLinkInstance.GetTotalTransform()` to transform grid curve to host coords.

### 6.2 DimensionReferenceResolver — Host Wall Face

```
Algorithm:
1. Get sleeve's host element via FamilyInstance.Host
2. If host is Wall: wall.GetReferences(FaceReferenceType.CenterReference) for centerline
   OR wall.GetReferences(FaceReferenceType.SideFaceReference) for face
3. If wall is in linked model: use RevitLinkInstance reference wrapping
4. Return both faces; caller picks the near face using dot product with view direction
```

### 6.3 DimensionLineCalculator — Position

```
Algorithm:
1. Get sleeve bounding box in the active view's coordinate system
2. Offset the dimension line AWAY from the sleeve by DimensionLineOffsetMm
3. For horizontal dimensions (X): line is horizontal, offset vertically above/below
4. For vertical dimensions (Y): line is vertical, offset horizontally left/right
5. For size dimensions: line passes through center of sleeve face
6. Avoid overlap: check existing dimension elements in view, shift if clash detected
```

### 6.4 OpeningSizeDimensionWriter

```
For circular: single value dimension showing diameter
  - refs[0] = left edge reference plane of sleeve family
  - refs[1] = right edge reference plane of sleeve family
  - dimension line through center

For rectangular: two dimensions W and H
  - W: refs = [left face, right face], horizontal line
  - H: refs = [top face, bottom face], vertical line
  - Both lines offset from sleeve center
```

### 6.5 ElevationDimensionWriter (BOO from FFL)

```
Algorithm:
1. Get sleeve center Z (in feet, Revit internal units)
2. Get host floor/slab level elevation
3. Compute BOO = (sleeve Z - clearance/2) - level elevation
4. In a section/elevation view: create vertical dimension from level datum to bottom of opening
5. Use Level.GetPlaneReference() for level reference
```

---

## 7. BEST PRACTICES ALIGNMENT

| Principle | Implementation |
|---|---|
| **Single Responsibility** | Each writer class handles one dimension type |
| **Open/Closed** | Add new categories via new IDimensionWriter implementations; no core changes |
| **Dependency Inversion** | All services depend on interfaces; concrete classes injected |
| **Fail-safe** | Each dimension creation wrapped in try/catch; failures logged, never abort |
| **Transaction safety** | Caller (Command) owns the transaction; services are transaction-agnostic |
| **Undo-ability** | All dimensions created within a single named Transaction for one-step undo |
| **Linked model support** | Reference resolver handles both host and linked model references |
| **View scale awareness** | Dimension line offset scaled by view scale factor |
| **No hardcoded strings** | DimensionType names from `DimensionOptions`; fallback to project default |
| **Extensibility** | `CategoryFilter` list controls scope; Phase 3 just expands this list |

---

## 8. EXISTING CODE HOOKS — WHERE TO PLUG IN

| Existing File | Integration Point |
|---|---|
| [SleevePlacementOrchestrator.cs](Services/Placement/SleevePlacementOrchestrator.cs) | Optional post-pipeline auto-dim step after batch flush |
| [DimensionStage.cs](Services/Placement/Stages/DimensionStage.cs) | Currently computes SIZE only; add flag to also trigger annotation writer |
| [PlacementContext.cs](Services/Placement/PlacementContext.cs) | Add `DimensionResults` property to track annotation element IDs |
| [GlobalApplicationSettings.cs](Models/GlobalApplicationSettings.cs) | Add `AutoDimensionAfterPlacement`, `DefaultDimensionOptions` |
| [ExternalEventHandlers.cs](Commands/ExternalEventHandlers.cs) | Register AutoDimensionCommand external event |
| Main ribbon/app registration | Add "Auto Dimension" button to JSE panel |

---

## 9. TRANSACTION PATTERN (Critical)

Revit API requires all model modifications inside a Transaction.
All dimension creation must follow this pattern:

```csharp
using (var tx = new Transaction(doc, "JSE - Auto Dimension Openings"))
{
    tx.Start();
    try
    {
        var results = _autoDimService.DimensionAllSleevesInView(doc, view, options);
        tx.Commit();
        return results;
    }
    catch
    {
        tx.RollBack();
        throw;
    }
}
```

If running post-placement (inside orchestrator), use `SubTransaction` or ensure the caller's
outer transaction is still open.

---

## 10. ERROR HANDLING STRATEGY

```
DimensionResult.cs:
  - Success: bool
  - CreatedElementIds: List<ElementId>
  - SkippedCount: int (sleeves with no valid references found)
  - Errors: List<string>
  - WarningCount: int (e.g. no grid found within radius)
```

**User feedback:** Summary dialog after command:
- "X openings dimensioned successfully"
- "Y openings skipped (no grid reference within Zmm)"
- "Open log for details"

---

## 11. LINKED MODEL REFERENCE WRAPPING

This is the #1 complexity in Revit API dimensioning. Pattern:

```csharp
// For elements in linked model:
var linkInstance = GetLinkInstance(doc, linkedDocId);
var transform = linkInstance.GetTotalTransform();

// Wrap the reference:
var stableRef = linkInstance.UniqueId + ":"
              + linkedElement.UniqueId + ":"
              + geometryReference.ConvertToStableRepresentation(linkedDoc);

var wrappedRef = Reference.ParseFromStableRepresentation(doc, stableRef);
```

This is required when host wall is in a linked architectural or structural model
(very common in BW workflows).

---

## 12. UNIT TEST TARGETS

| Test Class | What to Test |
|---|---|
| `DimensionReferenceResolverTests` | Grid distance calculation, wall face resolution |
| `DimensionLineCalculatorTests` | Offset logic, direction calculation |
| `OpeningSizeDimensionWriterTests` | Ref array construction for circular/rectangular |
| `AutoDimensioningServiceTests` | End-to-end with mock doc (no Revit API required) |
| `DimensionViewSelectorTests` | View selection logic |

Use Moq or NSubstitute to mock `Document`, `View`, `FamilyInstance`.

---

## 13. RIBBON BUTTON

Location: JSE MEP Panel > "Annotation" group (new group)
Icon: Dimension symbol (ruler with arrows)
Name: "Auto Dimension BW"
Tooltip: "Automatically add dimension annotations to all builders work openings in the active view"

Future: "Auto Dimension All" for Phase 3 (all categories)

---

## 14. PHASED DELIVERY CHECKLIST

### Phase 1A - Core Services
- [ ] Create Models/Dimensioning/*.cs (4 files)
- [ ] Create IDimensionReferenceResolver, IDimensionWriter, IAutoDimensioningService
- [ ] Implement DimensionTypeResolver
- [ ] Implement DimensionReferenceResolver (grid + wall face, host model only)
- [ ] Implement DimensionLineCalculator
- [ ] Implement OpeningSizeDimensionWriter
- [ ] Implement OpeningLocationDimensionWriter
- [ ] Implement ElevationDimensionWriter
- [ ] Implement AutoDimensioningService (orchestrates writers)

### Phase 1B - Command
- [ ] Implement DimensionViewSelector
- [ ] Create AutoDimensionCommand.cs (IExternalCommand)
- [ ] Register in addin manifest / ribbon

### Phase 1C - Settings UI
- [ ] Create AutoDimensionSettingsDialog XAML
- [ ] Create AutoDimensionSettingsViewModel
- [ ] Connect to command (show before running)

### Phase 1D - Pipeline Integration
- [ ] Add AutoDimensionAfterPlacement flag to GlobalApplicationSettings
- [ ] Inject IAutoDimensioningService into SleevePlacementOrchestrator (optional)
- [ ] Add DimensionResults to PlacementContext

### Phase 2 - Linked Model Support
- [ ] Extend DimensionReferenceResolver for linked model references
- [ ] Test with arch + structural linked models

### Phase 3 - All Categories
- [ ] Expand CategoryFilter support
- [ ] Add category picker to UI

---

## 15. RECOMMENDED FILE NAMING CONVENTIONS

Follow existing project conventions:
- Services: `*Service.cs`, `*ServiceTests.cs`
- Interfaces: `I*.cs` in `Services/Interfaces/` or inline folder
- Models: flat name in `Models/Dimensioning/`
- Commands: `*Command.cs`
- Views: `*Dialog.xaml` + `*Dialog.xaml.cs`
- ViewModels: `*ViewModel.cs`

---

## Sources Referenced

- [AGACAD Cut Opening Webinar](https://agacad.com/blog/cut-opening-webinar-20210527)
- [Auto Dimension for Openings - AGACAD](https://agacad.com/tutorials/automated-dimensions-for-openings)
- [Auto Dimension Pack - Archithetics](https://archithetics.net/auto-dimension-pack-for-revit/)
- [Naviate Auto Dimension - Revit Structure Blog](https://revitstructureblog.wordpress.com/2024/04/19/naviate-structure-whats-new-auto-dimension/)
- [ConVoid Openings](https://www.conclass.tech/convoid)
- [RevitApiDocs - NewDimension](https://www.revitapidocs.com/2023/47b3977d-da93-e1a4-8bfa-f23a29e5c4c1.htm)
- [The Building Coder - Auto Dimension Filled Region](https://jeremytammik.github.io/tbc/a/1765_autodim_filled_region.html)
- [The Building Coder - Dimension Reference Hints](https://jeremytammik.github.io/tbc/a/1316_dim_ref_hints.htm)
- [jeremytammik/the_building_coder_samples](https://github.com/jeremytammik/the_building_coder_samples)
- [jeremytammik/ArcDimensionIssue](https://github.com/jeremytammik/ArcDimensionIssue)
- [The Building Coder - Create Dimension Between Two Lines](https://thebuildingcoder.typepad.com/blog/2012/09/create-dimension-between-two-lines.html)
- [The Building Coder - Retrieve Dimensioning References](https://thebuildingcoder.typepad.com/blog/2015/05/how-to-retrieve-dimensioning-references.html)
