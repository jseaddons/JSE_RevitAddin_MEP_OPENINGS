# BCF Multi-Project Collaboration — Implementation Plan
## JSE MEP Openings Addin

**Document Version:** 1.0
**Date:** 2026-02-23
**Status:** Planning / Active Development
**Reference:** `BCF_OPENINGS_EXPORT_IMPLEMENTATION_PLAN.md` (root), conVoid BCF pattern (CONVOID_NOTES.md `BCF/` folder)
**Audience:** Development team, technical leads

---

## 1. Why BCF Instead of APS

APS (Autodesk Platform Services) provides cloud sync and multi-project visibility but costs **$20–50 per project per month** ongoing. For a small MEP team running 10–20 active projects, that is $200–1000/month with no ceiling.

**BCF (BIM Collaboration Format)** achieves the same multi-project collaboration goal at **zero cost**:

| Capability | APS | BCF |
|---|---|---|
| Share clash/sleeve status across engineers | Cloud sync (paid) | BCF file exchange (free) |
| View issues in 3D | APS Viewer (paid) | Solibri / BIM Track / Navisworks (free/existing) |
| Per-project status registry | Custom DB + ACC APIs | One BCF file per project |
| Structural team reviews opening requests | ACC Issues API | Import BCF into any IFC tool |
| Approval tracking | Custom ACC workflow | BCF topic status: Open → In Progress → Closed |
| Cost | $20–50/project/month | $0 |
| Internet required | Yes (cloud) | No (file-based) |
| Supports unlimited projects | Paid per use | Yes, unlimited |

BCF is an **ISO 19650-aligned open standard** supported by every major BIM tool: Solibri, BIM Track, Navisworks, Revit (native import), ArchiCAD, Tekla, and all coordination platforms. File exchange replaces cloud sync.

---

## 2. What BCF Provides for Multi-Project Workflows

### 2.1 The Multi-Project Problem (Current State)

```
[Engineer A - Machine 1]           [Engineer B - Machine 2]
  Revit + JSE Addin                  Revit + JSE Addin
  SQLite DB (local)                  SQLite DB (local)
  clash_zones.db                     clash_zones.db  ← SEPARATE, unsynced
  sleeve_snapshots.db                sleeve_snapshots.db ← SEPARATE

[Structural Engineer]              [Project Manager]
  No visibility into              No status dashboard
  opening requirements            No progress tracking
```

**Problems:**
- Structural team cannot see what openings are needed without screen sharing
- Project manager has no cross-project view of sleeve placement progress
- Two MEP engineers on the same project have no coordination mechanism
- Opening approval status exists only in the addin database — not sharable

### 2.2 BCF Solution Architecture

```
[Engineer A] → runs JSE Addin
     │
     ├── Export BCF (one file per project)
     │     └── Project_X_Sleeves_2026-02-23.bcf
     │           ├── Topic per ClashZone → "Sleeve required - Duct through Floor at Level 3"
     │           ├── BCF status: Open / In Progress / Closed / Unresolved
     │           ├── Comments: detection timestamp, placement result, failure reason
     │           ├── Viewpoint: camera anchored to clash zone XYZ in model space
     │           └── Component: MEP element + host element highlighted
     │
     ├── Share via email / SharePoint / USB
     │
     ▼
[Structural Engineer] → opens in Solibri / BIM Track / Navisworks
     │    Reviews each sleeve request, marks as Approved/Rejected, adds comments
     │    Returns modified BCF file
     ▼
[Engineer A] → imports BCF back into JSE Addin
     │    Updates ClashZone.ApprovalStatus in SQLite from BCF topic status
     │    Prioritises placement of Approved sleeves
     │    Re-flags Rejected zones for review
     ▼
[Project Manager] → opens BCF in any BIM tool
     Dashboard: Total openings / Placed / Pending / Rejected per project
```

### 2.3 BCF File Content Per Project

Each exported BCF file contains:
- **1 BCF Topic per ClashZone** (regardless of sleeve placement status)
- Topics grouped by status: Placed (Closed), Pending (Open), Failed (Unresolved), Cluster (In Progress)
- **Viewpoints** with camera positioned at each clash zone, MEP + host highlighted
- **Comments** with full operation history (detected at, placed at, cluster ID, failure reason)
- **Extensions** with JSE custom data: sleeve dimensions, MEP category, host type, cluster membership

---

## 3. Current Addin Data → BCF Mapping

### 3.1 ClashZone Table → BCF Topic

| ClashZone Field | BCF Field | Value / Mapping |
|---|---|---|
| `ZoneId` (GUID) | `Topic.Guid` | Direct — stable across exports |
| `MepElementId` + `HostElementId` | `Topic.Title` | `"Sleeve required — {MepCategory} through {HostType} at {LevelName}"` |
| `IsResolved = true` | `Topic.TopicStatus` | `Closed` |
| `IsResolved = false, PlacementAttempted = false` | `Topic.TopicStatus` | `Open` |
| `IsResolved = false, PlacementAttempted = true` | `Topic.TopicStatus` | `Unresolved` |
| Cluster member | `Topic.TopicStatus` | `In Progress` |
| `ClashZone.Level` | `Topic.Labels[]` | Level name as label |
| `MepCategory` (Duct/Pipe/Conduit/Tray) | `Topic.Labels[]` | Category as label |
| `HostType` (Wall/Floor/Framing) | `Topic.Labels[]` | Host type as label |
| `ClashZone.CreatedAt` | `Topic.CreationDate` | ISO 8601 |
| `SleeveSnapshot.PlacedAt` | `Topic.ModifiedDate` | Last action timestamp |

### 3.2 Sleeve Dimensions → BCF Comment

Each ClashZone topic gets a structured comment with the computed sleeve data:

```
[JSE MEP Openings] Sleeve Computed
Width: 350mm | Height: 250mm | Depth: 300mm
MEP: HVAC Supply Duct (id: 1234567)
Host: Concrete Wall 300mm (id: 7654321)
Cluster: cluster-abc-001 (8 members)
Placed at: 2026-02-23 14:32:11
Result: SUCCESS
```

### 3.3 Failure Reasons → BCF Comment

Failed placement zones include a diagnostic comment:

```
[JSE MEP Openings] Placement FAILED
Reason: Clearance insufficient — 45mm available, 50mm minimum required
Risk classification: High-risk (rotated cluster)
Recommended action: Manual review required
```

### 3.4 BCF Status Lifecycle

```
Detection → BCF: Open
     │
     ├── Placement attempt succeeded → BCF: Closed
     ├── Placement attempt failed    → BCF: Unresolved
     ├── In cluster, awaiting batch  → BCF: In Progress
     └── Structural approved         → BCF: (label added: "Approved")

Re-import structural BCF:
     BCF Closed + "Approved" label → ClashZone.StructuralApproval = true
     BCF Unresolved + comment      → ClashZone.RejectionReason = comment text
```

---

## 4. Implementation Architecture

### 4.1 New Service and Model Files

```
Services/
├── BcfExportService.cs              ← Core: reads SQLite, writes BCF archive
├── BcfImportService.cs              ← Core: reads BCF, updates ClashZone.ApprovalStatus
├── BcfTopicBuilder.cs               ← Maps ClashZone → BcfTopic + Comments
├── BcfViewpointBuilder.cs           ← Computes camera XYZ from ClashZone centroid
└── BcfFileManager.cs                ← ZIP/unzip .bcfzip, folder structure, manifest

Models/Bcf/
├── BcfProject.cs                    ← Top-level container (project.bcfp)
├── BcfTopic.cs                      ← markup.bcf per topic
├── BcfViewpoint.cs                  ← viewpoint.bcfv
├── BcfComment.cs                    ← Comment inside markup.bcf
└── BcfExportOptions.cs              ← User-configurable export settings

Data/
└── ClashZoneRepository.cs           ← Extend: GetAllForBcfExport(), UpdateApprovalFromBcf()

Views/
├── BcfExportDialog.xaml             ← Export settings dialog (existing plan)
└── BcfImportDialog.xaml             ← NEW: import structural BCF back
```

### 4.2 BCF File Structure (BCF 2.1)

```
Project_X_Sleeves_2026-02-23.bcfzip
├── bcf.version                      ← {"Major": 2, "Minor": 1}
├── project.bcfp                     ← project GUID + name
└── {TopicGuid}/                     ← one folder per ClashZone
      ├── markup.bcf                 ← topic title, status, labels, comments
      ├── viewpoint.bcfv             ← camera position, component GUIDs
      └── snapshot.png               ← optional; skip to keep file small
```

**BCF 2.1 is the target version** (widest tool support, used by Solibri, BIM Track, Revit native).

### 4.3 Core Export Logic

```csharp
// Services/BcfExportService.cs
public class BcfExportService
{
    private readonly ISleeveDbContext _db;
    private readonly BcfTopicBuilder _topicBuilder;
    private readonly BcfViewpointBuilder _viewpointBuilder;
    private readonly BcfFileManager _fileManager;

    /// <summary>
    /// Export all clash zones for the current project to a BCF 2.1 file.
    /// Called from the addin UI — runs on a background thread, reports progress.
    /// </summary>
    public async Task<BcfExportResult> ExportProjectAsync(
        string projectName,
        string outputPath,
        BcfExportOptions options,
        IProgress<BcfExportProgress> progress,
        CancellationToken ct)
    {
        // 1. Load all clash zones + sleeve snapshots from SQLite
        var zones = await _db.ClashZones
            .Include(z => z.SleeveSnapshot)
            .ToListAsync(ct);

        // 2. Create BCF project container
        var bcfProject = new BcfProject
        {
            ProjectId = _db.ProjectId,      // stable GUID per project DB
            ProjectName = projectName
        };

        // 3. Generate one topic per zone
        int done = 0;
        foreach (var zone in zones)
        {
            ct.ThrowIfCancellationRequested();

            var topic    = _topicBuilder.Build(zone, options);
            var viewport = _viewpointBuilder.Build(zone);   // XYZ from zone centroid
            bcfProject.Topics.Add((topic, viewport));

            progress?.Report(new BcfExportProgress(++done, zones.Count, zone.ZoneId));
        }

        // 4. Write .bcfzip archive to outputPath
        await _fileManager.WriteArchiveAsync(bcfProject, outputPath, ct);

        return new BcfExportResult
        {
            TopicCount   = zones.Count,
            ClosedCount  = zones.Count(z => z.IsResolved),
            OpenCount    = zones.Count(z => !z.IsResolved && !z.PlacementAttempted),
            FailedCount  = zones.Count(z => !z.IsResolved && z.PlacementAttempted),
            OutputPath   = outputPath
        };
    }
}
```

### 4.4 Viewpoint Builder — Camera from ClashZone Centroid

```csharp
// Services/BcfViewpointBuilder.cs
public class BcfViewpointBuilder
{
    private const double ViewDistanceMeters = 3.0;

    public BcfViewpoint Build(ClashZone zone)
    {
        // Zone stores centroid in Revit internal units (feet)
        // BCF requires metres
        double x = zone.CentroidX * 0.3048;
        double y = zone.CentroidY * 0.3048;
        double z = zone.CentroidZ * 0.3048;

        // Default camera: offset back and up, looking at centroid
        return new BcfViewpoint
        {
            Guid = Guid.NewGuid().ToString(),
            PerspectiveCamera = new PerspectiveCamera
            {
                CameraViewPoint  = new Point3D(x - ViewDistanceMeters,
                                               y - ViewDistanceMeters,
                                               z + ViewDistanceMeters),
                CameraDirection  = Normalise(new Vector3D(1, 1, -1)),
                CameraUpVector   = new Vector3D(0, 0, 1),
                FieldOfView      = 60.0
            },
            // Highlight MEP and host elements by IFC GUID or Revit ID
            Components = new BcfComponents
            {
                Visibility = new ComponentVisibility { DefaultVisibility = true },
                Selection  = new[]
                {
                    new Component { OriginatingSystem = "JSE MEP Openings",
                                    AuthoringToolId  = zone.MepElementId.ToString() },
                    new Component { OriginatingSystem = "JSE MEP Openings",
                                    AuthoringToolId  = zone.HostElementId.ToString() }
                }
            }
        };
    }

    private Vector3D Normalise(Vector3D v)
    {
        double len = Math.Sqrt(v.X*v.X + v.Y*v.Y + v.Z*v.Z);
        return new Vector3D(v.X/len, v.Y/len, v.Z/len);
    }
}
```

### 4.5 Import Service — Structural Feedback

```csharp
// Services/BcfImportService.cs
public class BcfImportService
{
    private readonly ISleeveDbContext _db;

    /// <summary>
    /// Read a BCF file returned by structural engineer.
    /// Update ClashZone approval status in SQLite based on topic status + labels.
    /// </summary>
    public async Task<BcfImportResult> ImportStructuralFeedbackAsync(
        string bcfPath,
        CancellationToken ct)
    {
        var archive = await BcfFileManager.ReadArchiveAsync(bcfPath, ct);
        int approved = 0, rejected = 0, skipped = 0;

        foreach (var topic in archive.Topics)
        {
            // Match by Topic.Guid == ClashZone.ZoneId
            var zone = await _db.ClashZones
                .FirstOrDefaultAsync(z => z.ZoneId.ToString() == topic.Guid, ct);

            if (zone == null) { skipped++; continue; }

            // Structural "Closed" + label "Approved" → structural approval granted
            if (topic.TopicStatus == "Closed" && topic.Labels.Contains("Approved"))
            {
                zone.StructuralApproval  = true;
                zone.ApprovalComment     = GetLatestComment(topic);
                approved++;
            }
            // "Unresolved" or comment "Rejected" → flag for manual review
            else if (topic.TopicStatus == "Unresolved" ||
                     topic.Labels.Contains("Rejected"))
            {
                zone.StructuralRejected  = true;
                zone.RejectionReason     = GetLatestComment(topic);
                rejected++;
            }
        }

        await _db.SaveChangesAsync(ct);

        return new BcfImportResult
        {
            TotalTopics = archive.Topics.Count,
            Approved = approved,
            Rejected = rejected,
            Skipped  = skipped        // not found in this project's DB
        };
    }

    private string GetLatestComment(BcfTopic topic) =>
        topic.Comments.OrderByDescending(c => c.Date).FirstOrDefault()?.Comment ?? string.Empty;
}
```

---

## 5. Multi-Project Workflow — Practical Usage

### 5.1 Engineer Workflow (No APS, No Cloud)

```
WEEK 1 — MEP Engineer
━━━━━━━━━━━━━━━━━━━━
1. Run Clash Detection on Project X → 150 clash zones in SQLite
2. Run Sleeve Placement → 120 placed, 30 failed/pending
3. File → Export BCF → saves "ProjectX_Sleeves_2026-02-23.bcfzip"
4. Email BCF to Structural Engineer (or drop in SharePoint folder)

WEEK 1 — Structural Engineer
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
5. Open BCF in Solibri / BIM Track (free tools)
6. Review each sleeve location in 3D (viewpoints included)
7. Mark approved topics: status Closed + label "Approved"
8. Mark rejected topics: status Unresolved + comment "Too close to rebar zone"
9. Returns BCF: "ProjectX_Sleeves_REVIEWED_2026-02-26.bcfzip"

WEEK 2 — MEP Engineer
━━━━━━━━━━━━━━━━━━━━
10. File → Import BCF → 118 Approved, 2 Rejected
11. Addin auto-prioritises Approved zones for batch re-run
12. Re-export updated BCF → "ProjectX_Sleeves_Final_2026-03-01.bcfzip"
    (all 118 Closed, 2 Unresolved for manual fix)
```

### 5.2 Multi-Project Summary (Project Manager View)

Since each project produces a BCF file, the project manager can open all BCF files in a BCF platform (BIM Track, Trimble Connect free tier, etc.) and see a dashboard:

```
Project A: 234 topics — 230 Closed / 4 Open
Project B: 89 topics  — 45 Closed / 44 Open  ← needs attention
Project C: 412 topics — 400 Closed / 12 Unresolved
```

This requires **zero cloud subscription** — just a shared folder of BCF files.

### 5.3 Sharing Channels

| Channel | Setup | Cost | Suitable For |
|---|---|---|---|
| Email attachment | None | Free | Small projects (< 5MB BCF) |
| SharePoint / OneDrive | Existing Office 365 | Free (existing) | Team sharing |
| Network share | Company server | Free | Office LAN |
| BIM Track (free tier) | 2 active projects | Free | Structured review |
| Trimble Connect (free) | Unlimited projects | Free with limits | BIM platform |
| GitHub / GitLab | Code-repo style | Free | Version-controlled BCF |

---

## 6. BCF Topic Builder — Complete Mapping

### 6.1 Topic Title Patterns

```
Sleeve required — Duct through Concrete Floor at Level 3
Sleeve required — Supply Pipe through RC Wall (200mm) at Basement
Sleeve required — Cable Tray through Composite Slab at Roof Plant
CLUSTER: 8-member duct cluster through Floor at Level 2 (cluster-abc-001)
FAILED: Duct through RC Wall — clearance insufficient at Level 1
```

### 6.2 Label Taxonomy

```
[MEP Category]   Duct | Pipe | Conduit | CableTray | Flexible
[Host Type]      Floor | Wall | Framing | Slab
[Status]         Placed | Pending | Failed | Cluster
[Priority]       Approved | Rejected | Review   ← added by structural engineer
[Level]          Level 1 | Level 2 | Roof | Basement ...
```

### 6.3 Comment Thread Structure

```
Comment 1 (automated, export time):
  Author: JSE MEP Openings
  Date:   2026-02-23T14:32:11
  Text:   "Clash detected. MEP: Duct 600x300 (id:1234567).
            Host: Concrete Wall 300mm (id:7654321).
            Sleeve computed: Width=650mm Height=350mm Depth=310mm."

Comment 2 (automated, if placed):
  Author: JSE MEP Openings
  Date:   2026-02-23T15:01:44
  Text:   "Sleeve placed successfully. Family: JSE_MEP_RectSleeve.
            Cluster: cluster-abc-001 (8 members). Batch #3."

Comment 3 (from structural, after review):
  Author: J.Smith (structural)
  Date:   2026-02-26T09:15:00
  Text:   "Approved. Coordinate with reinforcement drawing S-104."
```

---

## 7. UI Integration

### 7.1 Export Button in EmergencyMainDialog

Add a "BCF Export" button to the existing results panel in `EmergencyMainDialog.xaml.cs`:

```csharp
private async void OnBcfExportClick(object sender, EventArgs e)
{
    var dialog = new BcfExportDialog();
    if (dialog.ShowDialog() != true) return;

    var opts = dialog.Options;
    var outputPath = ShowSaveFileDialog("BCF Files|*.bcfzip",
                                        $"{_projectName}_Sleeves_{DateTime.Now:yyyy-MM-dd}.bcfzip");
    if (outputPath == null) return;

    var progress = new Progress<BcfExportProgress>(p =>
        UpdateStatusBar($"Exporting BCF: {p.Done}/{p.Total} topics..."));

    try
    {
        var result = await _bcfExportService.ExportProjectAsync(
            _projectName, outputPath, opts, progress, CancellationToken.None);

        MessageBox.Show(
            $"BCF export complete.\n\n" +
            $"Topics: {result.TopicCount}\n" +
            $"  Placed (Closed): {result.ClosedCount}\n" +
            $"  Pending (Open):  {result.OpenCount}\n" +
            $"  Failed:          {result.FailedCount}\n\n" +
            $"File: {outputPath}",
            "BCF Export", MessageBoxButton.OK, MessageBoxImage.Information);
    }
    catch (Exception ex)
    {
        MessageBox.Show($"BCF export failed: {ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
```

### 7.2 Import Button — Structural Feedback

```csharp
private async void OnBcfImportClick(object sender, EventArgs e)
{
    var filePath = ShowOpenFileDialog("BCF Files|*.bcfzip;*.bcf");
    if (filePath == null) return;

    try
    {
        var result = await _bcfImportService.ImportStructuralFeedbackAsync(
            filePath, CancellationToken.None);

        MessageBox.Show(
            $"BCF import complete.\n\n" +
            $"Topics read:  {result.TotalTopics}\n" +
            $"  Approved:   {result.Approved}\n" +
            $"  Rejected:   {result.Rejected}\n" +
            $"  Skipped:    {result.Skipped} (not in this project)\n\n" +
            $"SQLite updated. Re-run Placement to prioritise approved zones.",
            "BCF Import", MessageBoxButton.OK, MessageBoxImage.Information);

        RefreshClashZoneGrid();   // update the UI grid
    }
    catch (Exception ex)
    {
        MessageBox.Show($"BCF import failed: {ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
```

### 7.3 Required NuGet Package

Use **xBim Toolkit** — open source, zero cost, .NET Framework 4.8 compatible:

```xml
<!-- JSE_MEP_OPENINGS.csproj -->
<PackageReference Include="Xbim.BCF" Version="5.0.*" />
```

xBim.BCF handles BCF 2.0 and 2.1 read/write, ZIP packaging, XML serialisation — no custom XML writer needed. This is what conVoid's `BCF/` folder pattern is built on.

---

## 8. Database Changes Required

### 8.1 ClashZone Table — New Columns

```sql
ALTER TABLE ClashZones ADD COLUMN StructuralApproval  INTEGER NOT NULL DEFAULT 0;
ALTER TABLE ClashZones ADD COLUMN StructuralRejected   INTEGER NOT NULL DEFAULT 0;
ALTER TABLE ClashZones ADD COLUMN ApprovalComment      TEXT;
ALTER TABLE ClashZones ADD COLUMN RejectionReason      TEXT;
ALTER TABLE ClashZones ADD COLUMN LastBcfExportAt              TEXT;   -- ISO 8601 timestamp
ALTER TABLE ClashZones ADD COLUMN HasUnacknowledgedBcfUpdate   INTEGER NOT NULL DEFAULT 0; -- red dot
```

### 8.2 C# Entity Update

```csharp
// Data/ClashZone.cs (add to existing entity)
public bool   StructuralApproval  { get; set; } = false;
public bool   StructuralRejected  { get; set; } = false;
public string ApprovalComment     { get; set; }
public string RejectionReason     { get; set; }
public DateTime? LastBcfExportAt  { get; set; }
```

### 8.3 Migration

Add to `SleeveDbContext.EnsureSchema()` (or equivalent migration method):

```csharp
private void AddBcfColumnsIfMissing(SQLiteConnection conn)
{
    var existingCols = GetColumns(conn, "ClashZones");
    if (!existingCols.Contains("StructuralApproval"))
        conn.Execute("ALTER TABLE ClashZones ADD COLUMN StructuralApproval INTEGER NOT NULL DEFAULT 0");
    if (!existingCols.Contains("StructuralRejected"))
        conn.Execute("ALTER TABLE ClashZones ADD COLUMN StructuralRejected INTEGER NOT NULL DEFAULT 0");
    if (!existingCols.Contains("ApprovalComment"))
        conn.Execute("ALTER TABLE ClashZones ADD COLUMN ApprovalComment TEXT");
    if (!existingCols.Contains("RejectionReason"))
        conn.Execute("ALTER TABLE ClashZones ADD COLUMN RejectionReason TEXT");
    if (!existingCols.Contains("LastBcfExportAt"))
        conn.Execute("ALTER TABLE ClashZones ADD COLUMN LastBcfExportAt TEXT");
}
```

---

## 9. Revit View Color Coding — conVoid Pattern

After importing structural BCF feedback, placed sleeves are colorised in the active Revit view using `OverrideGraphicSettings`, **matching the conVoid StatusManager color scheme exactly**.

### 9.1 Color Scheme (conVoid-aligned)

| State | Revit View Color | Hex | Meaning |
|---|---|---|---|
| **Approved** by structural | Green | `#00AA00` (RGB 0,170,0) | Structural sign-off received, safe to proceed |
| **Placed, awaiting review** | Cyan/Blue | `#0088CC` (RGB 0,136,204) | Sleeve placed, BCF exported, no feedback yet |
| **Pending** (not yet placed) | Yellow | `#FFCC00` (RGB 255,204,0) | Clash detected, placement not yet run |
| **Rejected** by structural | Red | `#CC0000` (RGB 204,0,0) | Structural objection — needs manual resolution |
| **Failed placement** | Orange | `#FF6600` (RGB 255,102,0) | Placement attempted, Revit error — needs review |
| **No status** / unprocessed | Gray | `#888888` (RGB 136,136,136) | Detected, not yet processed |

This directly mirrors conVoid's approach: their `StatusManager` applies the same green/yellow/red/orange palette to void openings in the Revit view through view filter overrides.

### 9.2 Revit Color Override Service

```csharp
// Services/BcfRevitColorOverrideService.cs
public class BcfRevitColorOverrideService
{
    /// <summary>
    /// Apply conVoid-aligned color overrides to all placed sleeve elements
    /// based on their StructuralApproval / StructuralRejected flags in SQLite.
    /// Call this after BcfImportService completes.
    /// Must run inside a Revit Transaction on the Revit API thread.
    /// </summary>
    public void ApplyColorOverrides(Document doc, View activeView,
                                    IEnumerable<ClashZone> zones)
    {
        using var tx = new Transaction(doc, "BCF: Apply approval colours");
        tx.Start();

        foreach (var zone in zones)
        {
            if (zone.PlacedSleeveFamilyInstanceId == null) continue;

            var elemId = new ElementId(zone.PlacedSleeveFamilyInstanceId.Value);
            var ogs    = new OverrideGraphicSettings();

            Color revitColor = GetRevitColor(zone);
            ogs.SetProjectionLineColor(revitColor);
            ogs.SetSurfaceForegroundPatternColor(revitColor);
            ogs.SetSurfaceTransparency(20); // slight transparency so structure shows through

            activeView.SetElementOverrides(elemId, ogs);
        }

        tx.Commit();
    }

    private static Color GetRevitColor(ClashZone zone)
    {
        // Priority: Rejected > Approved > Failed > Placed/Awaiting > Pending > Gray
        if (zone.StructuralRejected)   return new Color(204, 0,   0);    // Red
        if (zone.StructuralApproval)   return new Color(0,   170, 0);    // Green
        if (!zone.IsResolved &&
             zone.PlacementAttempted)  return new Color(255, 102, 0);    // Orange - failed
        if (zone.IsResolved)           return new Color(0,   136, 204);  // Cyan - placed, awaiting
        if (!zone.PlacementAttempted)  return new Color(255, 204, 0);    // Yellow - pending
        return                                new Color(136, 136, 136);  // Gray - unprocessed
    }
}
```

### 9.3 Revit View Filter (Alternative — persistent across sessions)

For persistent coloring that survives session restarts (matching how conVoid implements it via shared parameters + view filters):

```
Revit View Filter name: "JSE_BCF_Approved"
  → Filter by shared parameter: JSE_StructuralApproval == 1
  → Override: surface color Green #00AA00

Revit View Filter name: "JSE_BCF_Rejected"
  → Filter by shared parameter: JSE_StructuralApproval == 0 AND JSE_StructuralRejected == 1
  → Override: surface color Red #CC0000

Revit View Filter name: "JSE_BCF_Pending"
  → Filter by shared parameter: JSE_Placed == 0
  → Override: surface color Yellow #FFCC00
```

These shared parameters (`JSE_StructuralApproval`, `JSE_StructuralRejected`, `JSE_Placed`) are written to the sleeve family instance by `ParameterTransferService` after BCF import — the same mechanism already used by `ParameterTransferService.cs` for other sleeve parameters.

### 9.4 BCF Viewer Color (External Tools)

BCF 2.1 topic status maps to standard viewer colors automatically in all tools:

| BCF TopicStatus | Solibri / BIM Track color | Our mapping |
|---|---|---|
| `Open` | Red/Pink | Pending (not placed) |
| `In Progress` | Yellow | Cluster in progress |
| `Closed` | Green | Placed / Approved |
| `Unresolved` | Red (darker) | Failed / Rejected |

No custom configuration needed — BCF viewers handle this by spec.

---

### Phase 1 — Core Export (Week 1)
- [ ] Add `Xbim.BCF` NuGet reference
- [ ] Create `Models/Bcf/` classes (BcfProject, BcfTopic, BcfViewpoint, BcfComment)
- [ ] Implement `BcfFileManager` (write BCF 2.1 ZIP archive)
- [ ] Implement `BcfTopicBuilder` (ClashZone → markup.bcf)
- [ ] Implement `BcfViewpointBuilder` (centroid XYZ → camera)
- [ ] Implement `BcfExportService.ExportProjectAsync()`
- [ ] Add BCF columns to ClashZone (migration in EnsureSchema)
- [ ] Wire "Export BCF" button in EmergencyMainDialog

### Phase 2 — Import / Structural Feedback (Week 2)
- [ ] Implement `BcfImportService.ImportStructuralFeedbackAsync()`
- [ ] Match BCF topic GUIDs to ClashZone.ZoneId
- [ ] Parse structural labels (Approved / Rejected)
- [ ] Update SQLite ClashZone.StructuralApproval / StructuralRejected
- [ ] Wire "Import BCF" button in EmergencyMainDialog
- [ ] Update clash zone grid to show approval status column
- [ ] Implement `BcfRevitColorOverrideService` — apply conVoid color scheme after import
  - Green = Approved, Red = Rejected, Yellow = Pending, Orange = Failed, Cyan = Placed/Awaiting
- [ ] Write `JSE_StructuralApproval` / `JSE_StructuralRejected` shared params to sleeve instances (via `ParameterTransferService`) for persistent view filter coloring

### Phase 3 — Placement Integration (Week 3)
- [ ] `BulkPlacementService`: prioritise `StructuralApproval = true` zones first
- [ ] `BulkPlacementService`: skip `StructuralRejected = true` zones (log reason)
- [ ] Add filter in EmergencyMainDialog: "Show only Approved" / "Show only Rejected"
- [ ] Auto-update BCF topic status after placement re-run (re-export diff only)

### Phase 4 — Cluster BCF Topics (Week 4)
- [ ] Group cluster members under a single parent BCF topic
- [ ] Cluster topic viewpoint: camera covers all cluster members (bounding box)
- [ ] Cluster comment: list all member element IDs + cluster algorithm used
- [ ] Handle cluster-within-cluster hierarchy in BCF

### Phase 5 — Testing and Validation (Week 5)
- [ ] Open exported BCF in: Solibri, BIM Track, Navisworks, Revit native BCF import
- [ ] Verify topic GUIDs survive round-trip (export → review → import)
- [ ] Performance test: 500 zone export target < 10 seconds
- [ ] Validate BCF 2.1 schema compliance with BuildingSMART validator

---

## 10. BCF Status UI — conVoid Pattern

### 10.0 What conVoid Manager Actually Is (Source: support.conclass.tech)

ConVoid Manager is a **dedicated modeless manager window** — NOT a tab inside the placement dialog. It is a separate window opened from the Revit ribbon, providing a full coordination hub. Key confirmed behaviours from the official documentation:

```
conVoid Manager — Confirmed UI Structure
────────────────────────────────────────
┌─ Toolbar ──────────────────────────────────────────────────────────────┐
│  [Filter checkboxes: Drawn Status ☐ ☐ ☐  |  Approval Status ☐ ☐ ☐]   │
│  [Export ▾]  [BCF Exchange ▾]  [Sync]  [Column visibility ▾]          │
└────────────────────────────────────────────────────────────────────────┘
┌─ Main Table (all openings) ────────────────────┬── Comments Panel ──────┐
│ • Revit ID  (🔴 red dot = new approval update) │  (activates when a     │
│ • Host Element                                 │   row is selected)     │
│ • Level                                        │                        │
│ • Drawn Status    [colour-coded badge]         │  [Comment text box]    │
│ • Approval Status [colour-coded badge]         │  [Snapshot button]     │
│   — one column per discipline                  │  [Set Approval button] │
│     e.g. Structural | Architect | Fire         │  [Edit metadata]       │
│ • Date Modified                                │  [Delete issue]        │
│                                                │                        │
│  Row click → 3D view navigates to opening      │                        │
└────────────────────────────────────────────────┴────────────────────────┘

Key facts confirmed:
  ✅ TWO separate statuses: Drawn Status (placed?) + Approval Status (approved?)
  ✅ Multi-discipline: separate approval column per discipline (BCF labels)
  ✅ Red notification dot on Revit-ID cell when new BCF import updates approval
  ✅ Comments Panel is a SIDE PANEL (not a tab), activates on row selection
  ✅ BCF Exchange menu: Export selected openings as BCF, Import BCF for approval
  ✅ Export options: Excel, CSV, BCF
  ✅ BCF labels distinguish discipline approvals (Structural / Architect / Fire)
  ✅ BCF auto-creates one issue per opening with Snapshot + Viewpoint + unique title
  ✅ Multi-approval: structural AND architect can both approve/reject independently
```

### 10.1 Our Implementation — BCF Manager Window

Replicate the conVoid Manager as a **dedicated BCF Manager dialog** (`BcfManagerDialog`) launched from the existing `EmergencyMainDialog` toolbar button. The existing placement workflow is untouched — BCF lives in its own window.

### 10.2 BcfManagerDialog — XAML Layout

Separate modeless window (`Views/BcfManagerDialog.xaml`), launched by a "BCF Manager" button in `EmergencyMainDialog`. Matches conVoid's two-panel layout: table on the left, Comments Panel on the right.

```xml
<!-- Views/BcfManagerDialog.xaml -->
<Window x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.BcfManagerDialog"
        Title="JSE — BCF Manager" Width="1100" Height="650"
        WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- Toolbar -->
            <RowDefinition Height="*"/>     <!-- Table + Comments Panel -->
        </Grid.RowDefinitions>

        <!-- ① Toolbar — filter checkboxes + action menus (conVoid pattern) -->
        <Border Grid.Row="0" Background="#2E3440" Padding="10,6">
            <DockPanel>
                <!-- Left: status filter chips -->
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                    <TextBlock Text="Drawn:" Foreground="#CCC" VerticalAlignment="Center" Margin="0,0,6,0"/>
                    <CheckBox Content="Placed"    IsChecked="{Binding ShowPlaced}"
                              Foreground="White" Margin="0,0,8,0"/>
                    <CheckBox Content="Pending"   IsChecked="{Binding ShowPending}"
                              Foreground="White" Margin="0,0,8,0"/>
                    <CheckBox Content="Failed"    IsChecked="{Binding ShowFailed}"
                              Foreground="White" Margin="0,0,20,0"/>

                    <TextBlock Text="Approval:" Foreground="#CCC" VerticalAlignment="Center" Margin="0,0,6,0"/>
                    <CheckBox Content="Approved"  IsChecked="{Binding ShowApproved}"
                              Foreground="White" Margin="0,0,8,0"/>
                    <CheckBox Content="Rejected"  IsChecked="{Binding ShowRejected}"
                              Foreground="White" Margin="0,0,8,0"/>
                    <CheckBox Content="Awaiting"  IsChecked="{Binding ShowAwaiting}"
                              Foreground="White" Margin="0,0,20,0"/>
                </StackPanel>

                <!-- Right: action buttons -->
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Right" HorizontalAlignment="Right">
                    <Button Content="Export BCF"     Click="OnExportBcfClick"
                            Margin="0,0,6,0" Width="100" Height="26"/>
                    <Button Content="Import BCF"     Click="OnImportBcfClick"
                            Margin="0,0,6,0" Width="100" Height="26"/>
                    <Button Content="Apply Colours"  Click="OnApplyColoursClick"
                            Margin="0,0,6,0" Width="110" Height="26"/>
                    <Button Content="Export Excel"   Click="OnExportExcelClick"
                            Margin="0,0,6,0" Width="100" Height="26"/>
                    <Button Content="Reset Approval" Click="OnResetApprovalClick"
                            Foreground="Salmon"      Width="110" Height="26"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- ② Main area: Table LEFT + Comments Panel RIGHT (conVoid split) -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>             <!-- Table -->
                <ColumnDefinition Width="5"/>             <!-- Splitter -->
                <ColumnDefinition Width="300"
                    MinWidth="0"
                    x:Name="CommentsPanelColumn"/>        <!-- Comments Panel -->
            </Grid.ColumnDefinitions>

            <!-- Table -->
            <DataGrid Grid.Column="0"
                      ItemsSource="{Binding FilteredZones}"
                      SelectedItem="{Binding SelectedZone}"
                      AutoGenerateColumns="False"
                      IsReadOnly="True"
                      SelectionMode="Single"
                      SelectionChanged="OnZoneSelectionChanged"
                      MouseDoubleClick="OnZoneDoubleClick">
                <DataGrid.Columns>

                    <!-- Revit-ID with red notification dot (conVoid pattern) -->
                    <DataGridTemplateColumn Header="Zone ID" Width="130">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <StackPanel Orientation="Horizontal">
                                    <!-- Red dot: new BCF approval not yet acknowledged -->
                                    <Ellipse Width="8" Height="8" Fill="Red"
                                             Margin="0,0,4,0" VerticalAlignment="Center"
                                             Visibility="{Binding HasNewApproval,
                                                Converter={StaticResource BoolToVisConverter}}"/>
                                    <TextBlock Text="{Binding ShortZoneId}"
                                               VerticalAlignment="Center"/>
                                </StackPanel>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>

                    <DataGridTextColumn Header="MEP Type" Binding="{Binding MepCategory}" Width="85"/>
                    <DataGridTextColumn Header="Host"     Binding="{Binding HostType}"    Width="75"/>
                    <DataGridTextColumn Header="Level"    Binding="{Binding LevelName}"   Width="75"/>
                    <DataGridTextColumn Header="Dims mm"  Binding="{Binding DimSummary}"  Width="90"/>

                    <!-- Drawn Status — colour-coded badge (conVoid status 1) -->
                    <DataGridTemplateColumn Header="Drawn Status" Width="100">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <Border Background="{Binding DrawnStatusColor}" CornerRadius="3"
                                        Padding="4,2" HorizontalAlignment="Center">
                                    <TextBlock Text="{Binding DrawnStatusLabel}"
                                               Foreground="White" FontSize="10" FontWeight="Bold"/>
                                </Border>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>

                    <!-- Approval Status — colour-coded badge (conVoid status 2) -->
                    <DataGridTemplateColumn Header="Structural" Width="90">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <Border Background="{Binding StructuralApprovalColor}" CornerRadius="3"
                                        Padding="4,2" HorizontalAlignment="Center">
                                    <TextBlock Text="{Binding StructuralApprovalLabel}"
                                               Foreground="White" FontSize="10" FontWeight="Bold"/>
                                </Border>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>

                    <!-- Future discipline columns added dynamically -->
                    <!-- e.g. Architect, Fire — populated from BCF label names in imported file -->

                    <DataGridTextColumn Header="Modified"
                                        Binding="{Binding LastModified, StringFormat='dd/MM HH:mm'}"
                                        Width="90"/>
                </DataGrid.Columns>
            </DataGrid>

            <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" Background="#4C566A"/>

            <!-- ③ Comments Panel (RIGHT side, activates on row selection — conVoid pattern) -->
            <Border Grid.Column="2" BorderBrush="#DEE2E6" BorderThickness="1,0,0,0"
                    Visibility="{Binding IsZoneSelected,
                        Converter={StaticResource BoolToVisConverter}}">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>   <!-- Zone summary -->
                        <RowDefinition Height="*"/>      <!-- Comment history -->
                        <RowDefinition Height="Auto"/>   <!-- Add comment -->
                        <RowDefinition Height="Auto"/>   <!-- Action buttons -->
                    </Grid.RowDefinitions>

                    <!-- Zone summary header -->
                    <StackPanel Grid.Row="0" Margin="0,0,0,8">
                        <TextBlock Text="{Binding SelectedZone.Title}"
                                   FontWeight="Bold" TextWrapping="Wrap" FontSize="12"/>
                        <TextBlock Text="{Binding SelectedZone.LevelName}"
                                   Foreground="#888" FontSize="11"/>
                    </StackPanel>

                    <!-- Comment history (all timestamps + authors) -->
                    <ListBox Grid.Row="1" ItemsSource="{Binding SelectedZoneHistory}"
                             BorderThickness="0" FontSize="11">
                        <ListBox.ItemTemplate>
                            <DataTemplate>
                                <StackPanel Margin="0,4,0,4">
                                    <StackPanel Orientation="Horizontal">
                                        <TextBlock Text="{Binding Date, StringFormat='yyyy-MM-dd HH:mm'}"
                                                   Foreground="#888" FontSize="10"/>
                                        <TextBlock Text="  "/>
                                        <TextBlock Text="{Binding Author}"
                                                   FontWeight="Bold" FontSize="10"/>
                                    </StackPanel>
                                    <TextBlock Text="{Binding Comment}"
                                               TextWrapping="Wrap" Margin="0,2,0,0"/>
                                </StackPanel>
                            </DataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>

                    <!-- Add new comment -->
                    <TextBox Grid.Row="2" Text="{Binding NewComment, UpdateSourceTrigger=PropertyChanged}"
                             Height="60" TextWrapping="Wrap" AcceptsReturn="True"
                             Margin="0,6,0,6"
                             Tag="Add a comment..."/>

                    <!-- Panel action buttons -->
                    <StackPanel Grid.Row="3" Orientation="Vertical">
                        <Button Content="Add Comment"
                                Click="OnAddCommentClick" Height="26" Margin="0,0,0,4"/>
                        <Button Content="Set Approved"
                                Click="OnSetApprovedClick" Height="26"
                                Background="#00AA00" Foreground="White" Margin="0,0,0,4"/>
                        <Button Content="Set Rejected"
                                Click="OnSetRejectedClick" Height="26"
                                Background="#CC0000" Foreground="White" Margin="0,0,0,4"/>
                        <Button Content="Navigate to Zone"
                                Click="OnNavigateClick" Height="26"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>
    </Grid>
</Window>
```

### 10.3 BcfStatusChip — Reusable Colour Counter

Small colour-coded circle + count, the same pill/chip pattern conVoid uses in its status bar:

```xml
<!-- Views/Controls/BcfStatusChip.xaml -->
<UserControl x:Class="JSE_RevitAddin_MEP_OPENINGS.Views.Controls.BcfStatusChip">
    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
        <Ellipse Width="12" Height="12"
                 Fill="{Binding Color, RelativeSource={RelativeSource AncestorType=UserControl}}"/>
        <TextBlock Text="{Binding LabelText, RelativeSource={RelativeSource AncestorType=UserControl}}"
                   FontWeight="Bold" FontSize="13" Margin="4,0,0,0" VerticalAlignment="Center"/>
    </StackPanel>
</UserControl>
```

### 10.4 ViewModel Properties Required

Add to the existing `EmergencyMainDialog` code-behind (or a dedicated `BcfStatusViewModel`):

```csharp
// Bound to the BCF tab summary bar
public int BcfApprovedCount { get; private set; }
public int BcfRejectedCount { get; private set; }
public int BcfPendingCount  { get; private set; }
public int BcfFailedCount   { get; private set; }
public int BcfAwaitingCount { get; private set; }
public string LastBcfExportText { get; private set; } = "Not yet exported";

// Bound to the DataGrid
public ObservableCollection<BcfZoneRow> BcfZones { get; } = new();
public BcfZoneRow SelectedBcfZone { get; set; }

// Bound to the history ListBox
public ObservableCollection<BcfCommentRow> SelectedZoneHistory { get; } = new();

/// <summary>
/// Reload BCF counters and grid from SQLite after import or placement.
/// Call after: BcfImportService completes, BulkPlacementService completes.
/// </summary>
private void RefreshBcfTab()
{
    var zones = _db.ClashZones.ToList();

    BcfApprovedCount = zones.Count(z => z.StructuralApproval);
    BcfRejectedCount = zones.Count(z => z.StructuralRejected);
    BcfPendingCount  = zones.Count(z => !z.IsResolved && !z.PlacementAttempted
                                         && !z.StructuralApproval && !z.StructuralRejected);
    BcfFailedCount   = zones.Count(z => !z.IsResolved && z.PlacementAttempted);
    BcfAwaitingCount = zones.Count(z => z.IsResolved
                                         && !z.StructuralApproval && !z.StructuralRejected);

    var lastExport = zones.Where(z => z.LastBcfExportAt.HasValue)
                          .Max(z => z.LastBcfExportAt);
    LastBcfExportText = lastExport.HasValue
        ? $"Last export: {lastExport.Value:yyyy-MM-dd HH:mm}"
        : "Not yet exported";

    BcfZones.Clear();
    foreach (var z in zones)
        BcfZones.Add(new BcfZoneRow(z));   // maps zone → display row

    OnPropertyChanged(nameof(BcfApprovedCount));
    OnPropertyChanged(nameof(BcfRejectedCount));
    OnPropertyChanged(nameof(BcfPendingCount));
    OnPropertyChanged(nameof(BcfFailedCount));
    OnPropertyChanged(nameof(BcfAwaitingCount));
    OnPropertyChanged(nameof(LastBcfExportText));
}
```

### 10.5 BcfZoneRow — Display Model (Two-Status, conVoid Pattern)

```csharp
// Models/Bcf/BcfZoneRow.cs
public class BcfZoneRow
{
    public BcfZoneRow(ClashZone z)
    {
        ZoneId       = z.ZoneId;
        ShortZoneId  = z.ZoneId.ToString()[..8];   // first 8 chars for display
        MepCategory  = z.MepCategory;
        HostType     = z.HostType;
        LevelName    = z.LevelName;
        DimSummary   = z.SleeveSnapshot != null
            ? $"{z.SleeveSnapshot.Width}x{z.SleeveSnapshot.Height}"
            : "—";
        LastModified = z.LastBcfExportAt ?? z.CreatedAt;
        Title        = $"{z.MepCategory} through {z.HostType} at {z.LevelName}";

        // Red dot: new BCF import brought a new approval/rejection not yet acknowledged
        HasNewApproval = z.HasUnacknowledgedBcfUpdate;

        // STATUS 1 — Drawn Status (was the sleeve placed?) — conVoid "Drawn Status" column
        if (!z.PlacementAttempted)
        {
            DrawnStatusColor = "#FFCC00"; DrawnStatusLabel = "Pending";
        }
        else if (z.IsResolved)
        {
            DrawnStatusColor = "#0088CC"; DrawnStatusLabel = "Placed";
        }
        else
        {
            DrawnStatusColor = "#FF6600"; DrawnStatusLabel = "Failed";
        }

        // STATUS 2 — Approval Status (structural feedback) — conVoid "Approval Status" column
        // In conVoid, each discipline gets its own column; we start with Structural only
        if (z.StructuralRejected)
        {
            StructuralApprovalColor = "#CC0000"; StructuralApprovalLabel = "Rejected";
        }
        else if (z.StructuralApproval)
        {
            StructuralApprovalColor = "#00AA00"; StructuralApprovalLabel = "Approved";
        }
        else
        {
            StructuralApprovalColor = "#888888"; StructuralApprovalLabel = "—";
        }

        LatestComment = z.StructuralRejected ? z.RejectionReason
                      : z.StructuralApproval ? z.ApprovalComment
                      : string.Empty;
    }

    public Guid     ZoneId                  { get; }
    public string   ShortZoneId             { get; }
    public string   Title                   { get; }
    public string   MepCategory             { get; }
    public string   HostType                { get; }
    public string   LevelName               { get; }
    public string   DimSummary              { get; }
    public DateTime LastModified            { get; }
    public string   LatestComment           { get; }
    public bool     HasNewApproval          { get; }   // drives red dot visibility

    // Drawn Status (STATUS 1) — conVoid "Drawn Status" column
    public string   DrawnStatusColor        { get; }
    public string   DrawnStatusLabel        { get; }

    // Structural Approval (STATUS 2) — conVoid per-discipline approval column
    public string   StructuralApprovalColor { get; }
    public string   StructuralApprovalLabel { get; }

    // Future: ArchitectApprovalColor / FireApprovalColor when multi-discipline is needed
}
```

The `HasUnacknowledgedBcfUpdate` flag (new column in ClashZone) is set to `true` when BCF import brings in a new Approved/Rejected status, and cleared to `false` when the user clicks the zone row in the manager — matching conVoid's red notification dot behaviour.

### 10.6 Double-Click → Jump to Revit Viewpoint

When the user double-clicks a row, navigate the active 3D view camera to the clash zone centroid — same behaviour as conVoid's BCF viewpoint navigation:

```csharp
private void OnBcfZoneDoubleClick(object sender, MouseButtonEventArgs e)
{
    if (SelectedBcfZone == null) return;

    // Raise an ExternalEvent to navigate on the Revit API thread
    _navigateToBcfZoneHandler.ZoneId = SelectedBcfZone.ZoneId;
    _navigateToBcfZoneEvent.Raise();
}

// ExternalEventHandler:
public class NavigateToBcfZoneHandler : IExternalEventHandler
{
    public Guid ZoneId { get; set; }

    public void Execute(UIApplication uiApp)
    {
        var zone = /* load from DB by ZoneId */;
        var view3d = /* get or create a 3D view */;

        // Convert Revit feet → UI units
        var pt = new XYZ(zone.CentroidX, zone.CentroidY, zone.CentroidZ);
        uiApp.ActiveUIDocument.ShowElements(new[] { new ElementId(zone.MepElementId) });

        // Orient camera
        var orientation = new ViewOrientation3D(
            eye:    pt + new XYZ(-3, -3, 3),   // back and up
            up:     XYZ.BasisZ,
            forward: (pt - (pt + new XYZ(-3,-3,3))).Normalize());

        view3d.SetOrientation(orientation);
        uiApp.ActiveUIDocument.ActiveView = view3d;
    }

    public string GetName() => "NavigateToBcfZone";
}
```

### 10.7 Integration Points

| Event | Action |
|---|---|
| "BCF Manager" button in EmergencyMainDialog | Open `BcfManagerDialog` modelessly |
| BcfManagerDialog opens | Load all zones from SQLite → populate table |
| Toolbar filter checkbox toggled | Filter `FilteredZones` collection in-memory (no DB query) |
| Row selected | Activate Comments Panel (right side); load `SelectedZoneHistory`; clear `HasUnacknowledgedBcfUpdate` flag |
| Row double-click | Raise `NavigateToBcfZoneEvent` → camera jumps to zone in Revit 3D view |
| "Set Approved" / "Set Rejected" button in panel | Update `StructuralApproval`/`StructuralRejected` in SQLite; add comment; refresh row badge |
| "Export BCF" toolbar button | Run `BcfExportService`; set `LastBcfExportAt` on exported zones; refresh table |
| "Import BCF" toolbar button | Run `BcfImportService`; set `HasUnacknowledgedBcfUpdate = true` on changed zones; refresh table (red dots appear) |
| "Apply Colours" toolbar button | Raise ExternalEvent → `BcfRevitColorOverrideService.ApplyColorOverrides()` on Revit thread |
| "Reset Approval" toolbar button | Confirm dialog → clear all approval flags in DB → refresh table |
| "Export Excel" toolbar button | Write `FilteredZones` to `.xlsx` (EPPlus or ClosedXML) |
| Placement run completes (in EmergencyMainDialog) | If `BcfManagerDialog` is open, call `Refresh()` on it to update Drawn Status badges |

---

## 11. Files to Create / Modify

### New Files

| File | Purpose |
|---|---|
| `Services/BcfExportService.cs` | Core export orchestration |
| `Services/BcfImportService.cs` | Structural feedback import |
| `Services/BcfRevitColorOverrideService.cs` | conVoid-aligned Revit view color overrides post-import |
| `Services/BcfTopicBuilder.cs` | ClashZone → BCF topic XML |
| `Services/BcfViewpointBuilder.cs` | Camera position from centroid |
| `Services/BcfFileManager.cs` | ZIP archive read/write |
| `Models/Bcf/BcfProject.cs` | Top-level BCF container |
| `Models/Bcf/BcfTopic.cs` | Topic model (markup.bcf) |
| `Models/Bcf/BcfViewpoint.cs` | Viewpoint model |
| `Models/Bcf/BcfComment.cs` | Comment model |
| `Models/Bcf/BcfExportOptions.cs` | User export settings |
| `Models/Bcf/BcfExportResult.cs` | Export result summary |
| `Models/Bcf/BcfImportResult.cs` | Import result summary |
| `Views/BcfManagerDialog.xaml` | Dedicated BCF Manager window (conVoid Manager pattern) |
| `Views/BcfManagerDialog.xaml.cs` | Code-behind: filter logic, table refresh, panel activation, red dot clear |
| `Models/Bcf/BcfZoneRow.cs` | Display model — two-status (DrawnStatus + ApprovalStatus), red dot flag |
| `Models/Bcf/BcfCommentRow.cs` | Display model for status history ListBox in Comments Panel |
| `ExternalEvents/NavigateToBcfZoneHandler.cs` | Jump Revit 3D camera to selected zone on double-click |

### Modified Files

| File | Change |
|---|---|
| `Data/ClashZone.cs` | Add BCF approval columns |
| `Data/SleeveDbContext.cs` | Add BCF column migration in `EnsureSchema()` |
| `Data/ClashZoneRepository.cs` | Add `GetAllForBcfExport()`, `UpdateApprovalFromBcf()` |
| `Views/EmergencyMainDialog.xaml.cs` | Add "BCF Manager" toolbar button that opens `BcfManagerDialog` modelessly |
| `Services/BulkPlacementService.cs` | Prioritise `StructuralApproval = true` zones |
| `JSE_MEP_OPENINGS.csproj` | Add `Xbim.BCF` NuGet reference |

---

## 12. Comparison: BCF vs APS — Decision Summary

| Factor | BCF (chosen) | APS (rejected) |
|---|---|---|
| **Cost** | $0 | $20–50/project/month |
| **Internet required** | No | Yes |
| **Works offline / site** | Yes | No |
| **Structural team access** | Any BCF viewer (free) | APS Viewer or ACC license |
| **Multi-project** | Unlimited, file-per-project | Paid per project |
| **Standard** | ISO, open, permanent | Proprietary Autodesk |
| **Tool support** | All BIM tools | Autodesk ecosystem only |
| **Version history** | File timestamps / Git | ACC versioning (paid) |
| **Implementation effort** | Medium (xBim toolkit) | High (OAuth, cloud, DA4R) |
| **Maintenance** | None (file format stable) | API deprecations, billing |

**Conclusion:** BCF delivers the same multi-project collaboration capability — structural approval tracking, 3D viewpoints, status history — at zero ongoing cost, using an open standard that every stakeholder on a construction project already has tooling for.

---

## 13. References

- `BCF_OPENINGS_EXPORT_IMPLEMENTATION_PLAN.md` — Original BCF export design (root of repo)
- `CONVOID_NOTES.md` — conVoid's `BCF/` folder: import/export topics, viewpoints, cameras (reference implementation)
- `architecture/APS_MULTI_PROJECT_IMPLEMENTATION.md` — APS alternative (rejected: cost)
- [BuildingSMART BCF 2.1 Standard](https://www.buildingsmart.org/standards/bsi-standards/bim-collaboration-format/)
- [xBim Toolkit — BCF for .NET](https://github.com/xBimTeam/XbimBCF)
- [conVoid BCF demo — YouTube](https://www.youtube.com/watch?v=NroxalO4Wwk)
- [BCF Schema Validator](https://bcf-validator.buildingsmart.org/)
