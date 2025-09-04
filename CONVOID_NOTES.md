### conclassConVoid – structure and key notes

#### Build/runtime
- **Target framework**: .NET Framework v4.8 (classic csproj).
- **Output**: `conclassConVoid.dll` (Library). **Root namespace/assembly**: `conclassConVoid`.
- **UI tech**: WinForms (primary). WPF assemblies referenced but UI is WinForms-based.
- **Revit refs**: `RevitAPI`, `RevitAPIUI`, `AdWindows` (must match installed Revit version; icons suggest 2024).
- **Other refs**: `Newtonsoft.Json`, `Cryptolens.Licensing` (licensing), standard .NET assemblies.
- **Manifest**: no `.addin` file in repo. Deployment requires an external Revit add-in manifest pointing to the built DLL under `%AppData%/Autodesk/Revit/Addins/20XX`.

#### Entry points / Revit integration
- `Ribbon.cs` implements `IExternalApplication` with `OnStartup`/`OnShutdown`. Ribbon construction uses reflection helpers (obfuscation wrappers). Split buttons and panels appear to be added here.
- Multiple `IExternalCommand` implementations exist, many in obfuscated subfolders (class names look random). Named commands also live under `Commands/`.
- `ButtonAvailability.cs` implements `IExternalCommandAvailability` to control UI enablement.
- `ExternalEvents/` contains numerous `IExternalEventHandler` classes to bridge modeless WinForms UI actions to the Revit API (e.g., set parameters, create openings, dimensions, profiles, filters, selection boxes, status sync, etc.). Includes `DisableOnIdling`.

#### Major folders (business logic)
- `Commands/`
  - Feature groups: `ConfigManager/`, `ConvertVoid/`, `VoidManager/` (with `DisciplineManager`, `ParameterManager`, `StatusManager`, `VoidManagerCommands`, `VoidManagerContent`), `SmartTag/`, `Dimensions/`, `PlanCheck/`, `LevelElevation/`.
  - Utility logic: `GeometryLibary.cs`, `UnitManager.cs`, `IfcGuid.cs`, `MathUtil.cs`, `FilterElements.cs`, `ActiveFeatures.cs`.
- `ExternalEvents/`
  - Many modeless handlers: create openings/sections/dimensions, set parameters/levels/filters/oversize/minimum size, approval sync, clash zones, selection boxes, sheet fields, etc.
- `Interfaces/`
  - WinForms UI: `ConVoidInterface`, `ConVoidManager`, `ConVoidSettings`, `ConVoidManagerSettings`, `ConVoidManagerSync`, `ConVoidDimension`, `ConVoidDrillingZone`, `ConVoidPlanCheck`, `SmartTagMain`, `ConclassDiscipline`, `ConVoidHelp`, progress bars, schedule windows, parameter transfer, activation/license windows (`ConVoidActivate`, `ConVoidConnectLicenseServer`).
  - `WindowHandle_*` helpers for modal/modeless ownership.
- `DataSchema/`
  - Extensible Storage schema definitions and updates: `CreateConfigSchema`, `CreateSettingSchema`, `CreateManagerSchema`, `CreateDimensionSchema`, `CreateSmartTagSchema`, `UpdateDataSchema`.
- `BCF/`
  - BIM Collaboration Format support: import/export, topics, viewpoints, cameras, imaging utils.
- `Utils/`
  - `DocumentChangedManager`, `MonitorOnDocumentChange`, `PlaneEqualityComparer`, `Utils.cs`.
- `Resources/`
  - Ribbon icons 16x16/32x32 (many `*_2024_*` variants): ConvertVoid, ConVoidManager, SmartTag, DimensionVoid, PlanCheck, Drilling, Join, Level, etc.
- `Properties/`
  - `AssemblyInfo.cs`, `Resources`, `Settings`, `GlobalAppLanguage`.
- Top-level notable files: `Ribbon.cs`, `ButtonAvailability.cs`, `Geometry.cs`, `ConvexHull.cs`, `ConvexHullThreadUsage.cs`, `app.manifest`.

#### Obfuscation characteristics
- Thousands of files with randomized names and `-Module-*.cs` present.
- Frequent `MethodImplOptions.NoInlining`, reflection wrappers, and helper types with obfuscated identifiers used to invoke Revit UI/API methods.
- Control-flow/string obfuscation likely. Avoid renaming/moving these files; focus changes in clearly named folders (`Commands/`, `ExternalEvents/`, `Interfaces/`, etc.).

#### Licensing / activation
- `Cryptolens.Licensing` is referenced.
- UI forms for activation and license server connection exist (`ConVoidActivate`, `ConVoidConnectLicenseServer`). Expect license checks during startup/UI actions.

#### Functional themes (from structure)
- Void creation/management (approval status, parameters, oversize/min size, elevation).
- Conversion of elements to voids.
- Smart tagging and dimensioning of voids.
- Plan checks and level elevation tools.
- View filters and coloring, selection/section boxes, sheet field population.
- BCF import/export for issue exchange.

#### Void Manager quick map (focus for MEP openings)
- **Files**: `Commands/VoidManager/`
  - `VoidManagerCommands.cs`: entry/orchestration of manager flows, hooks from ribbon/UI.
  - `VoidManagerContent.cs`: core execution for finding hosts, computing/placing openings, batching ops.
  - `ParameterManager.cs`: applies/syncs element parameters on created openings.
  - `StatusManager.cs`: approval/manually placed/imported statuses and synchronization.
  - `DisciplineManager.cs`: toggles per-discipline logic (MEP groups, categories, filters).
  - `CreateManageSchema.cs`: Extensible Storage schema for manager settings/state.
- **Related**:
  - `ExternalEvents/*Set*` and `*Create*` handlers: modeless bridge used by the UI; optional if invoking direct code paths.
  - `DataSchema/*`: global/config/dimension schemas referenced by the manager.
  - `Commands/GeometryLibary.cs`, `Geometry.cs`: host/solid math utilities.
- **Typical data flow** (UI path): Ribbon → WinForms (Interfaces) → ExternalEvent → `VoidManagerContent` → Revit API transactions → `ParameterManager`/`StatusManager` → commit.
- **Headless/non-UI path**: Call `VoidManagerContent` (and managers) directly inside a controlled `TransactionGroup`/`Transaction`; avoid ExternalEvents if already on Revit's API thread (modal command).
- **Linked models**: Host detection and placement must resolve `RevitLinkInstance` transforms for floors/walls/framing in links (not just active doc). Keep transforms correct when computing opening location/orientation. This is mandatory.
- **Key parameters** (from schemas/manager): oversize, minimum size, elevation adjustments, approval status, discipline, profile shape.
- **Performance**: Pre-compute selection/filtering/geometry outside transactions; keep Revit API calls inside minimal-scoped transactions; avoid multi-threading API calls.
- **Logging**: Prefer appending to `placement*.log` when placing/adjusting openings (timestamps, element ids, host/link context).
- **Integration with your existing non-UI creator**: Reuse `VoidManagerContent` for computation/placement; bypass WinForms by invoking manager methods directly and setting config/state via schema managers. ExternalEvents only needed for modeless contexts.

#### Build/deploy quick notes
- Build with VS targeting .NET Framework 4.8, AnyCPU. Ensure `RevitAPI*` references resolve to your installed Revit (likely 2024 given resources).
- Create a `.addin` manifest (for the target Revit year) pointing to the built `conclassConVoid.dll`.
- If loading modeless WinForms, external events in `ExternalEvents/` are used; they must be created on startup (often held as static instances).

#### UI migration plan (adapt non-UI opening creation to modeless UI)
- **Approach**: Keep your existing non-UI opening creation logic; invoke it from a modeless WinForms flow via an `ExternalEventHandler` to stay API-safe.
- **Handler**: Reuse `ExternalEvents/CreateOpeningsExternalEvent.cs` or add a dedicated wrapper (e.g., `ExternalEvents/RunVoidManagerExternalEvent.cs`) that calls into `Commands/VoidManager/VoidManagerContent` with current settings.
- **Singletons**: Follow repo convention: static `HandlerInstance` + `HandlerEvent` on the handler class. Create/register these on add-in startup (see `Ribbon.OnStartup`).
- **UI**: Add/extend a form in `Interfaces/ConVoidManager*.cs` with your controls (oversize/min size, filters, elevation, approval, discipline, profile). On Apply/Run, push settings to schema/state and `HandlerEvent.Raise()`.
- **State/config**: Persist and read via `DataSchema/CreateSettingSchema.cs` (and related). Avoid passing large state through static fields; store/retrieve on run.
- **Core call**: In the handler `Execute`, acquire `UIDocument/Document`, resolve selection/targets, then call `VoidManagerContent` and downstream managers (`ParameterManager`, `StatusManager`).
- **Transactions**: Open a `Transaction`/`TransactionGroup` inside the handler. Keep geometry prep outside transactions when possible; API mutations inside.
- **Linked hosts**: Resolve `RevitLinkInstance` transform(s) before computing/placing openings; write element/log context to `placement*.log`.
- **Ribbon wiring**: Ensure a button on the Ribbon opens the modeless manager UI (form is shown modelessly and raises the external event on actions).
- **Modal fallback**: If you need a one-click command without UI, keep an `IExternalCommand` path that calls the same core logic without the ExternalEvent layer.

#### Domain-specific constraint
- Structural hosts (floors, walls, framing) may be in linked files. Placement/intersection/orientation logic should handle linked file proxies, not only native elements.

#### Gaps / to verify later
- No `.addin` file present – confirm external deployment process and addin GUIDs.
- Exact Revit version target (icons imply 2024) – confirm against installed references.
- Logging locations and formats (team convention mentions `placement*.log`).

#### Quick top-level map
```
conclassConVoid/
  Commands/            ExternalEvents/        Interfaces/        DataSchema/
  BCF/                 Utils/                 Resources/         Properties/
  Ribbon.cs            ButtonAvailability.cs  Geometry.cs        ConvexHull*.cs
  app.manifest         conclassConVoid.csproj  [many obfuscated *.cs]
```


