# Autodesk Platform Services (APS) — Multi-Project Implementation Plan
## JSE MEP Openings Addin

**Document Version:** 1.0
**Date:** 2026-02-23
**Status:** Planning / Research
**Audience:** Development team, technical leads

---

## 1. Executive Summary

This document describes how Autodesk Platform Services (APS) can extend the JSE MEP Openings addin from a **single-project, single-machine Revit addin** into a **multi-project, multi-user cloud-capable platform**.

The current addin works well for one engineer on one machine. The goal is to allow a **team of engineers across multiple live projects** to share clash data, track sleeve placement status, and optionally run detection headlessly in the cloud — without duplicating work or losing sync.

---

## 2. Current Architecture (Baseline)

```
[Engineer's Machine]
    Revit 2023–2026
        └── JSE MEP Openings Addin
              ├── Clash Detection (MepIntersectionService)
              ├── Sleeve Placement (BulkPlacementService)
              ├── Clustering (BatchClusterPlacementService)
              └── SQLite DB (SleeveDbContext)
                    └── C:\Users\...\AppData\JSE_MEP_Openings\
                          ├── clash_zones.db       ← per project
                          ├── sleeve_snapshots.db
                          └── Logs\
```

**Key constraints of the current model:**

| Constraint | Impact |
|---|---|
| SQLite is local per machine | Two engineers on the same project have separate, unsynced DBs |
| Revit must be open and interactive | No batch/overnight processing |
| No project registry | No way to see "all projects" status |
| Logs are local only | Support requires screen sharing to diagnose issues |
| Families/conditions are local XML | Different machines may have different config versions |

---

## 3. What APS Provides

APS (formerly Autodesk Forge) is a set of REST APIs running on Autodesk's cloud. The services relevant to this addin are:

| APS Service | Purpose | Relevance |
|---|---|---|
| **Design Automation for Revit (DA4R)** | Run Revit workloads headlessly in the cloud | Clash detection + sleeve placement without opening Revit |
| **Data Management API** | Cloud file/folder storage in ACC / BIM 360 hubs | Store and sync the SQLite DB per project |
| **Model Derivative API** | Translate RVT → SVF2 for web viewing | View placed sleeves without Revit |
| **APS Viewer (v7)** | Browser-based 3D viewer | Overlay clash zones on the model |
| **ACC APIs** | Issues, RFIs, project members | Link clash zones to formal ACC Issues |
| **Webhooks** | Event notifications (model uploaded, issue created) | Trigger auto-refresh when structural/MEP model updates |
| **Authentication (2-legged / 3-legged OAuth)** | Secure API access | Per-user and service-to-service calls |

---

## 4. Target Architecture (Multi-Project)

```
┌─────────────────────────────────────────────────────────────────┐
│                     APS CLOUD LAYER                             │
│                                                                 │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────────┐  │
│  │ Data Mgmt    │  │ Design Auto  │  │   APS Viewer +       │  │
│  │ API          │  │ for Revit    │  │   Model Derivative   │  │
│  │ (DB sync)    │  │ (DA4R)       │  │   (Web Dashboard)    │  │
│  └──────┬───────┘  └──────┬───────┘  └──────────┬───────────┘  │
│         │                 │                      │              │
│  ┌──────▼─────────────────▼──────────────────────▼───────────┐  │
│  │              JSE MEP Openings Backend Service              │  │
│  │         (ASP.NET Core / Azure Functions / AWS Lambda)      │  │
│  │    - Project registry                                      │  │
│  │    - Job queue (clash detection, placement runs)           │  │
│  │    - Aggregated reporting                                  │  │
│  └──────────────────────────────────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘
          ▲                          ▲
          │ DB sync (on save)        │ Trigger DA4R run
          │                          │
┌─────────┴──────────┐   ┌──────────┴──────────┐
│  Engineer A         │   │  Engineer B          │
│  Revit + Addin     │   │  Revit + Addin       │
│  Project: Tower A  │   │  Project: Tower B    │
│  Local SQLite DB   │   │  Local SQLite DB     │
│  (synced to APS)   │   │  (synced to APS)     │
└────────────────────┘   └─────────────────────┘
```

---

## 5. Multi-Project Data Model

### 5.1 Project Identity

Each project needs a unique identifier that links the local SQLite DB to the APS cloud record.

```
Project Registry (cloud)
├── ProjectId        : GUID (generated on first use)
├── ProjectName      : "Tower A - Block 1"
├── RevitModelUrn    : APS Model URN (from ACC / BIM 360)
├── LastRefresh      : DateTime
├── LastPlacement    : DateTime
├── TotalClashZones  : int
├── ResolvedZones    : int
├── ActiveEngineers  : List<UserId>
└── DbFileUrn        : APS object URN for the latest clash_zones.db
```

### 5.2 SQLite DB Sync Strategy

The current `SleeveDbContext` (SQLite) remains the **working database** on the local machine. APS Data Management acts as the **sync/backup layer**.

```
On Refresh Complete:
    → Upload clash_zones.db to APS bucket
    → Tag with ProjectId + Timestamp + ModelVersion

On Addin Startup (project open):
    → Check APS for newer DB version
    → If newer: download and merge (or replace)
    → If same: use local

Conflict resolution rule:
    → IsResolvedFlag=1 always wins (sleeve placed = fact)
    → Clash zones with newer timestamp win for non-resolved
```

### 5.3 Bucket Structure in APS Data Management

```
APS Bucket: jse-mep-openings-{company-id}
    ├── projects/
    │   ├── {projectId}/
    │   │   ├── clash_zones_{timestamp}.db
    │   │   ├── sleeve_snapshots_{timestamp}.db
    │   │   ├── conditions/
    │   │   │   ├── Ducts.xml
    │   │   │   ├── Pipes.xml
    │   │   │   └── CableTrays.xml
    │   │   └── families/
    │   │       ├── RectangularOpeningOnWall.rfa
    │   │       └── CircularOpeningOnFloor.rfa
    │   └── {projectId-2}/
    │       └── ...
    └── reports/
        └── {projectId}/
            └── placement_performance_{date}.log
```

---

## 6. Design Automation for Revit (DA4R)

### 6.1 What It Enables

DA4R runs a Revit AppBundle (your addin DLL + dependencies) against a cloud-hosted RVT file **without any UI, without opening Revit locally**. This means:

- Overnight batch clash detection across 5 projects → results ready in the morning
- Trigger a placement run from a web dashboard or Teams bot
- Consistent, reproducible results (same addin version, same conditions XML)

### 6.2 Entry Point Changes Required

Currently your addin uses `IExternalCommand.Execute()` which requires Revit to be open. For DA4R, you need `IExternalDBApplication`:

```csharp
// New DA4R entry point (wraps existing orchestrator)
public class JseMepOpeningsAutomation : IExternalDBApplication
{
    public ExternalDBApplicationResult OnStartup(ControlledApplication application)
    {
        application.ApplicationInitialized += OnApplicationInitialized;
        return ExternalDBApplicationResult.Succeeded;
    }

    private void OnApplicationInitialized(object sender, ApplicationInitializedEventArgs e)
    {
        // Read job parameters from input JSON (passed by DA4R)
        var jobParams = ReadJobParameters(); // filterName, operationMode, projectId

        var app = sender as Application;
        var doc = app.OpenDocumentFile(jobParams.ModelPath);

        // Reuse EXISTING orchestrator — no logic changes needed
        var orchestrator = new OpeningCommandOrchestrator(doc, ...);
        var result = orchestrator.Execute(jobParams.FilterName, jobParams.Mode);

        // Write results to output JSON (returned by DA4R)
        WriteJobResult(result);

        doc.Close(false);
    }
}
```

**Key point:** `OpeningCommandOrchestrator`, `BulkPlacementService`, `BatchClusterPlacementService` require **zero changes** for DA4R. Only the entry point changes.

### 6.3 Input / Output Bundle

```
DA4R Activity Input:
{
    "modelFile":        "urn:adsk.objects:...rvt",    // ACC model URN
    "dbFile":           "urn:adsk.objects:...db",     // clash_zones.db from APS
    "conditionsFile":   "urn:adsk.objects:...zip",    // conditions XMLs
    "jobParams": {
        "filterName":   "HVAC-L04",
        "operationMode": "RefreshAndPlace",            // or "RefreshOnly"
        "projectId":    "a1b2c3d4-..."
    }
}

DA4R Activity Output:
{
    "modifiedModel":    "output.rvt",                 // RVT with sleeves placed
    "updatedDb":        "clash_zones_updated.db",     // updated SQLite
    "report":           "placement_performance.log",  // existing perf log
    "summary": {
        "placed":       405,
        "clusters":     55,
        "failed":       0,
        "durationMs":   12590
    }
}
```

### 6.4 DA4R Constraints & Mitigations

| Constraint | Impact on This Addin | Mitigation |
|---|---|---|
| No interactive UI | All dialog prompts must be removed | Already handled — `DeploymentConfiguration.DeploymentMode = true` suppresses UI |
| No local file paths | SQLite path must be relative or passed as input | Change `SharedDbContextProvider` to accept path parameter |
| 1-hour timeout | Large models with many clashes | Batch into zones using section box slicing |
| Revit version must match | DA4R supports R2022–R2026 | Already multi-version with `#if REVIT2023` etc. |
| No Revit UI (no views) | Some operations need an active view | `doc.ActiveView` fallback to first 3D view |
| Family loading | RFA files must be bundled | Include in AppBundle zip |

### 6.5 Required Code Changes for DA4R

**File: `Services/SharedDbContextProvider.cs`**
```csharp
// Current: hardcoded AppData path
// Change: accept optional override path (from DA4R input)
public static string GetDbPath(string overridePath = null)
{
    if (!string.IsNullOrEmpty(overridePath))
        return overridePath;
    // existing AppData logic...
}
```

**File: `Services/FamilyLoadingService.cs`**
```csharp
// Current: searches local filesystem for RFA
// Change: also check working directory (DA4R bundle extraction path)
var bundlePath = Path.Combine(
    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
    "Families", familyName + ".rfa");
```

**File: `Services/DeploymentConfiguration.cs`**
```csharp
// Add DA4R detection
public static bool IsRunningInDesignAutomation =>
    Environment.GetEnvironmentVariable("ADSK_DA_ENVIRONMENT") == "1";

// DA4R implies deployment mode
public static bool DeploymentMode
{
    get
    {
        if (IsRunningInDesignAutomation) return true;
        if (OptimizationFlags.UseDiagnosticMode) return false;
        return _deploymentMode;
    }
}
```

---

## 7. APS Viewer — Web Dashboard

### 7.1 What the Dashboard Shows

A browser-based interface using APS Viewer (v7) that shows:

- The translated RVT model in 3D
- Clash zones as coloured bubble overlays
  - Red: unresolved
  - Green: sleeve placed
  - Blue: clustered
- Sleeve dimensions as tooltips
- Project-level KPIs: total clashes / resolved / pending

### 7.2 Technology Stack

```
Frontend:  React + APS Viewer SDK (v7)
Backend:   ASP.NET Core Web API (or Azure Functions)
Auth:      APS 3-legged OAuth (per user) + 2-legged (service)
Storage:   APS Data Management (DB files) + your existing SQLite schema
```

### 7.3 Clash Zone Overlay Implementation

```javascript
// Viewer extension: JseMepOpeningsExtension.js
class JseMepClashOverlay extends Autodesk.Viewing.Extension {
    onModelLoaded(viewer, model) {
        // Load clash zones from backend API
        fetch(`/api/projects/${projectId}/clashzones`)
            .then(r => r.json())
            .then(zones => this.renderOverlays(viewer, zones));
    }

    renderOverlays(viewer, zones) {
        zones.forEach(zone => {
            const color = zone.isResolved
                ? new THREE.Color(0x00ff00)   // green
                : new THREE.Color(0xff4444);   // red

            // Place sprite at clash zone world coordinates
            viewer.impl.createOverlay('clashZones', new THREE.Mesh(
                new THREE.SphereGeometry(0.1),
                new THREE.MeshBasicMaterial({ color })
            ));

            // Position in Revit internal units (converted to viewer units)
            overlay.position.set(
                zone.intersectionPointX * 0.3048,
                zone.intersectionPointY * 0.3048,
                zone.intersectionPointZ * 0.3048
            );
        });
    }
}
```

### 7.4 Backend API Endpoints Required

```
GET  /api/projects                          → list all projects
GET  /api/projects/{id}/clashzones          → all clash zones with status
GET  /api/projects/{id}/clashzones/{guid}   → single zone detail
POST /api/projects/{id}/refresh             → trigger DA4R refresh job
POST /api/projects/{id}/place               → trigger DA4R placement job
GET  /api/projects/{id}/jobs/{jobId}        → poll job status
GET  /api/projects/{id}/report              → latest performance report
GET  /api/auth/callback                     → APS OAuth callback
```

---

## 8. ACC Issues Integration

### 8.1 Link Clash Zones to ACC Issues

Each unresolved clash zone can be raised as a formal ACC Issue, giving it a workflow (Assigned → In Review → Closed):

```
ClashZone (SQLite)          ACC Issue (APS)
├── Id (GUID)          ←→   ├── id
├── ZoneDisplayName    ←→   ├── title
├── MepElementId       ←→   ├── linkedDocumentUrn (model element ref)
├── IntersectionPoint  ←→   ├── pushpinAttributes.location
├── IsResolved         ←→   ├── status ("open" / "closed")
└── PlacementStatus    ←→   └── assignedTo (engineer name)
```

### 8.2 Sync Rules

```
On ClashZone created (Refresh):
    → Create ACC Issue if not already linked
    → Store issue ID in ClashZone.AccIssueId (new column)

On Sleeve placed (BulkPlacementService):
    → PATCH ACC Issue status → "resolved"
    → Comment: "Sleeve placed: {FamilyName}, Size: {W}x{H}mm"

On Sleeve deleted (flag reset):
    → PATCH ACC Issue status → "open"
```

---

## 9. Implementation Phases

### Phase 1 — DB Sync (Low Risk, High Value)
**Duration:** 2–3 weeks
**Goal:** Shared SQLite across machines, no logic changes

| Task | Files Affected | Effort |
|---|---|---|
| Add `ProjectId` GUID to DB schema | `Data/SleeveDbContext.cs` | Small |
| Upload DB to APS bucket on refresh complete | New `ApsDbSyncService.cs` | Medium |
| Download latest DB on addin startup | `Application.cs` startup | Medium |
| Simple merge strategy (resolved wins) | New `DbMergeService.cs` | Medium |
| APS OAuth 2-legged setup | New `ApsAuthService.cs` | Small |

**Result:** Two engineers working on the same project see each other's placed sleeves after next refresh.

---

### Phase 2 — Project Dashboard (Medium Risk)
**Duration:** 4–6 weeks
**Goal:** Web interface showing all projects' status

| Task | Files Affected | Effort |
|---|---|---|
| ASP.NET Core backend API | New project | Medium |
| APS Model Derivative translation | New `ModelTranslationService.cs` | Medium |
| APS Viewer with clash zone overlay | New web project (React) | Large |
| Project registry in APS Data Management | New `ProjectRegistryService.cs` | Small |
| REST endpoints for clash zone data | New `ClashZoneController.cs` | Medium |

**Result:** Project manager opens browser, sees all projects, clicks into any model, sees sleeve placement progress in 3D.

---

### Phase 3 — Design Automation (Higher Risk)
**Duration:** 6–8 weeks
**Goal:** Headless clash detection + placement triggered from web

| Task | Files Affected | Effort |
|---|---|---|
| `IExternalDBApplication` entry point | New `JseMepOpeningsAutomation.cs` | Small |
| DB path parameterisation | `SharedDbContextProvider.cs` | Small |
| Family bundle packaging | `FamilyLoadingService.cs`, build scripts | Medium |
| DA4R activity + workitem setup | New `DesignAutomationService.cs` | Large |
| Job queue and polling | Backend API | Large |
| Multi-Revit version bundles | Build pipeline | Medium |
| Testing DA4R locally (DA4R emulator) | Dev tooling | Medium |

**Result:** Engineer clicks "Run Refresh" in web dashboard → DA4R picks up latest RVT from ACC → runs detection + placement → saves results back → model and DB updated in ACC.

---

### Phase 4 — ACC Issues (Optional)
**Duration:** 3–4 weeks
**Goal:** Formal issue tracking per clash zone

| Task | Files Affected | Effort |
|---|---|---|
| Add `AccIssueId` column to ClashZones | `Data/SleeveDbContext.cs` | Small |
| Create issues on clash zone creation | `ClashZoneService.cs` | Medium |
| Update issues on sleeve placement | `BulkPlacementService.cs` | Small |
| OAuth 3-legged (per-user auth) | `ApsAuthService.cs` | Medium |

---

## 10. Cost Estimate (APS)

APS pricing is usage-based. For typical project volumes:

| Service | Unit | Estimated Usage | Monthly Cost (USD) |
|---|---|---|---|
| Data Management storage | per GB | ~5 GB (DB files + families) | ~$1/month |
| Model Derivative | per translation | 2–5 translations/project/month | ~$5–15/project |
| Design Automation | per compute hour | 2–4 hrs/project/month | ~$15–30/project |
| Viewer | per token | Free for translated models | $0 |
| **Total per project** | | | **~$20–50/month** |

For 10 active projects: **~$200–500/month**.

> Note: Autodesk Construction Cloud (ACC) subscribers get included API calls. If the client already has ACC, Data Management and Issues API costs may be covered.

---

## 11. Risk Register

| Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|
| DA4R timeout on large models (>1hr) | Medium | High | Section box batching; split by floor |
| SQLite merge conflicts | Medium | Medium | Resolved-wins rule; timestamp-based |
| APS API rate limits | Low | Medium | Retry logic with exponential backoff |
| Revit version mismatch in DA4R | Low | High | Build one AppBundle per Revit version |
| Addin behaviour differs in DA4R vs local | Medium | High | DA4R emulator for local testing |
| ACC OAuth token expiry during long run | Low | Medium | Token refresh in `ApsAuthService` |
| Network unavailable (offline use) | Medium | Medium | Fall back to local-only mode gracefully |

---

## 12. Files to Create / Modify

### New Files

```
Services/
├── ApsAuthService.cs              ← OAuth 2-legged + 3-legged token management
├── ApsDbSyncService.cs            ← Upload/download SQLite DB to APS bucket
├── ApsProjectRegistryService.cs   ← Multi-project registry in APS Data Management
├── DbMergeService.cs              ← Merge clash zones from two DB versions
└── DesignAutomationService.cs     ← DA4R workitem submission + polling

Automation/
└── JseMepOpeningsAutomation.cs    ← IExternalDBApplication DA4R entry point
```

### Modified Files

```
Application.cs                     ← Check for DB updates on startup
Services/SharedDbContextProvider.cs ← Accept override DB path (for DA4R)
Services/DeploymentConfiguration.cs ← Add IsRunningInDesignAutomation flag
Services/FamilyLoadingService.cs   ← Check bundle directory for RFAs
Data/SleeveDbContext.cs            ← Add ProjectId + AccIssueId columns
```

### No Changes Required

The entire placement and clustering pipeline:
- `BulkPlacementService.cs`
- `BatchClusterPlacementService.cs`
- `OpeningCommandOrchestrator.cs`
- `MepIntersectionService.cs`
- All repositories and services under `Services/Clustering/`

This is the key benefit of the existing SOLID architecture — the core logic is already decoupled from the UI and machine-specific paths.

---

## 13. Recommended First Step

**Start with Phase 1 (DB Sync) only.**

It requires no cloud infrastructure, no DA4R setup, and no web frontend. Just:
1. Register an APS application (free, 5 minutes)
2. Create a bucket per company
3. Upload the SQLite DB to APS after each successful refresh
4. Download on startup if a newer version exists

This immediately solves the multi-user problem (two engineers, one project) with minimal risk and maximum value.

---

## 14. APS Application Registration

To get started:
1. Go to [aps.autodesk.com](https://aps.autodesk.com) → Create Account
2. Create a new application → note `CLIENT_ID` and `CLIENT_SECRET`
3. Store credentials securely (not in source code) — use `UserConfiguration` or environment variables
4. For Phase 1 (2-legged): request scope `data:read data:write bucket:create bucket:read`
5. For Phase 3 (DA4R): also request `code:all`

---

*Document created: 2026-02-23*
*Next review: After Phase 1 proof of concept*
