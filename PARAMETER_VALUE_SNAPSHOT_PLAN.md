### Element Parameter Value Snapshot Plan (Refresh → Store → Later Map)

#### Objective
- Minimize Refresh cost by snapshotting a curated, high-value parameter set only (not "all" parameters), then mapping later without re-querying Revit/links. If a user requests an unmapped parameter, fetch it on-demand and remember it for the next run.

#### What to Store (per clash zone)
- MepParameterValues: Dictionary<string,string>  // whitelisted parameter name → value at refresh time
- HostParameterValues: Dictionary<string,string>
- SourceDocKey / HostDocKey: string              // e.g., Document.PathName (distinguishes active vs linked docs)
- Optional: CollectedAt timestamp, DocHash (for staleness checks)

Notes:
- Store best-effort string values only. No geometry or binaries.
- Backward compatible: fields may be null/absent in older files.

#### Where It Lives (no new files)
- On the existing `ClashZone` entries inside `ClashZoneStorage` already serialized to XML.
- Main filter XML and per-category filter XMLs will automatically include the new fields when saved (no new persistence paths).

#### Integration Point
- File: `Services/RefreshService.cs`
- Method: `ExecuteRefresh(...)`
- Placement: Immediately after intersections (`currentIntersections`) are computed, and before saving to XML.

Flow:
1) Build a per-refresh parameter cache to avoid duplicate reads:
   - Key: `(docKey, elementId)` where `docKey = element.Document.PathName ?? "Unknown"`
   - Value: `Dictionary<string,string>` of whitelisted parameters
2) For each intersection tuple `(mepElement, structuralElement, bbox, point)`:
   - Resolve each element’s `Document` (linked or active) from `element.Document`.
   - Extract values via `ReadWhitelistedParams(Element, Whitelist)` (see below).
   - Create/update the `ClashZone`, assign `MepParameterValues`, `HostParameterValues`, `SourceDocKey`, `HostDocKey`.
3) Proceed with existing save to XML (main filter and category-specific files).

#### Harvesting Logic (whitelist + fallback)
- Parameter names come from a curated whitelist per role (MEP/Host) and element type, plus any previously learned keys from the same XML. Actual values are read from each `Element` in `currentIntersections`.
- `ReadWhitelistedParams(Element e, HashSet<string> whitelist) -> Dictionary<string,string>`:
  - For each requested `key` in `whitelist`, find parameter by name on the element; if not found, try the type (Symbol) as needed.
  - Convert using `AsString()` → `AsValueString()` → typed fallback (int/double/ElementId) with culture-invariant formatting.
  - Skip null/empty; return compact dictionary.

Whitelist definition (extensible):
- MEP common (applies to Duct/Pipe/CableTray/Accessory where present):
  - Size keys: `"Size"`, `"Diameter"`, `"Nominal Diameter"`, `"Width"`, `"Height"`
  - Level keys: `"Reference Level"`, `"Level"`, `"Schedule Level"`, `"Reference Level Elevation"` (if present)
  - System keys: `"System Type"`, `"System Classification"`, `"Service Type"`, `"System Abbreviation"`
- Host common (Wall/Floor/Structural):
  - `"Fire Rating"`
  - Room context: `"Room Name"`, `"Room Number"` (best-effort via nearest room at clash point)
  - Nearby grids: computed list of nearest grid names around clash point (e.g., up to 2 X and 2 Y)

Learning + fallback:
- If a user maps a parameter not present in the snapshot:
  - Fetch on-demand from the owning document/linked file for the specific element at mapping time.
  - Append the parameter name into `LearnedParameterKeys` in the XML.
  - On the next Refresh, merge `LearnedParameterKeys` into the whitelist so it gets snapshotted proactively.

Linked vs Active Document Handling:
- The `Element` already belongs to its owning `Document` (active or linked). Use `element.Document` directly.
- Use `docKey = element.Document.PathName ?? "Unknown"` to distinguish cross-document element ids.

#### Performance Safeguards
- Only process elements present in `currentIntersections` (not whole models).
- Per-refresh parameter cache by `(docKey, elementId)` to reuse extracted bags across multiple intersections involving the same element.
- Store strings only; optionally truncate extremely long values (e.g., >1KB) to keep XML small.
- Whitelist reduces per-element reads by ~80–95% vs. full scan; typical added Refresh time drops from ~1 min to a few seconds for thousands of elements.
- Keep measured timings (element count, keys persisted, milliseconds) in debug log per Refresh.

#### Persistence Flow (already present)
- Main filter save: `_filterManagementService.SaveFilterToXmlFile(targetFilter, mainFilePath)` will include updated `ClashZoneStorage` with parameter bags.
- Category-specific save: category filters copy clash zone lists; parameter bags are carried along into each category XML.
- New optional top-level fields in the same XML (backward compatible):
  - `ParameterKeyWhitelist` (stored once per file)
  - `LearnedParameterKeys` (stored once per file, merged into whitelist on next Refresh)

#### Later Mapping (Parameter Service)
- On opening the Parameter Service:
  - Load chosen filter/category XML.
  - For each mapping row, locate the `ClashZone` by stored element ids (existing identifiers) and read `MepParameterValues` or `HostParameterValues`.
  - If a requested key is missing in the bag, perform on-demand fetch for that element/param, then write; also add the key into `LearnedParameterKeys` so it’s included next run.

#### Minimal Pseudo-code (inside ExecuteRefresh)
```csharp
// After currentIntersections computed
var cache = new Dictionary<(string docKey, int id), Dictionary<string,string>>();
var learnedKeys = new HashSet<string>(LoadLearnedKeysFromXmlIfAny(), StringComparer.OrdinalIgnoreCase);
var whitelist = BuildWhitelist(currentIntersections)
    .Union(learnedKeys, StringComparer.OrdinalIgnoreCase)
    .ToHashSet(StringComparer.OrdinalIgnoreCase);

Dictionary<string,string> ReadWhitelistedParams(Element e, HashSet<string> keys)
{
    var bag = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
    foreach (var k in keys)
    {
        var p = e.LookupParameter(k) ?? (e is FamilyInstance fi ? fi.Symbol?.LookupParameter(k) : null);
        if (p == null) continue;

        string? value = p.AsString();
        if (string.IsNullOrEmpty(value)) value = p.AsValueString();
        if (string.IsNullOrEmpty(value) && p.StorageType == StorageType.Integer) value = p.AsInteger().ToString();
        if (string.IsNullOrEmpty(value) && p.StorageType == StorageType.Double) value = p.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (string.IsNullOrEmpty(value) && p.StorageType == StorageType.ElementId) value = p.AsElementId()?.IntegerValue.ToString();
        if (string.IsNullOrWhiteSpace(value)) continue;
        if (!bag.ContainsKey(k)) bag[k] = value;
    }
    return bag;
}

foreach (var (mep, host, bbox, pt) in currentIntersections)
{
    var mepKey  = (mep.Document.PathName ?? "Unknown",  mep.Id.IntegerValue);
    var hostKey = (host.Document.PathName ?? "Unknown", host.Id.IntegerValue);

    if (!cache.TryGetValue(mepKey, out var mepBag))  { mepBag  = ReadWhitelistedParams(mep, whitelist);  cache[mepKey]  = mepBag; }
    if (!cache.TryGetValue(hostKey, out var hostBag)) { hostBag = ReadWhitelistedParams(host, whitelist); cache[hostKey] = hostBag; }

    var cz = CreateOrUpdateClashZone(mep, host, bbox, pt, _document, clearanceSettings);
    cz.SourceDocKey        = mepKey.Item1;
    cz.HostDocKey          = hostKey.Item1;
    cz.MepParameterValues  = mepBag;
    cz.HostParameterValues = hostBag;
}

SaveLearnedKeysToXml(learnedKeys);
// Then save as usual (main + category XMLs)
```

#### Acceptance Checklist
- After Refresh:
  - Main filter XML contains `ClashZoneStorage` whose clash zones include `MepParameterValues`, `HostParameterValues`, `SourceDocKey`, `HostDocKey`.
  - Category-specific XMLs contain the same parameter bags, filtered by category as you already do.
- Parameter Service later reads values from XML and writes to opening parameters without Revit queries; if a missing key is requested, it fetches on-demand and records the key for inclusion in the next Refresh.


