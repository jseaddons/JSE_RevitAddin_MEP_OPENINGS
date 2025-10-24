# Clustering System Context Reference

## PRIMARY REFERENCE DOCUMENT
**ALWAYS READ FIRST**: `Docs/CLUSTERING_COMPLETE_GUIDE.md`

## CRITICAL ARCHITECTURE
```
Phase 1: Refresh Button
├── ClashZoneService → Detects intersections
├── UniversalSleevePlacerService → Places individual sleeves
├── SleeveCoordinateService → Creates _CLUSTER.xml files
└── Result: _CLUSTER.xml files ready for clustering

Phase 2: OK Button  
├── UniversalClusterService → Loads from _CLUSTER.xml files
├── Cache Building → Uses MepElementId as keys
├── Proximity Clustering → Groups by host type, system type, orientation
└── Result: Clustered sleeves replace individual sleeves
```

## CRITICAL SUCCESS FACTORS
1. ✅ Use MepElementId as cache key consistently
2. ✅ Read from _CLUSTER.xml files, not regular XML
3. ✅ Use MepElementOrientationDirection for orientation detection (NOT HostOrientation)
4. ✅ Apply correct coordinate system based on host type and orientation
5. ✅ Fix pipe strategy logic flow to prevent fallthrough to "NO STRATEGY MATCHED"
6. ✅ Fix SleeveCoordinateService category mapping to return "Pipes" not "Pipe Accessories"
7. ✅ Fix SleeveCoordinateService orientation detection to use clash zone data

## KEY FILES
- `Services/UniversalClusterService.cs` - Main clustering logic
- `Services/SleeveCoordinateService.cs` - Creates _CLUSTER.xml files
- `Services/UniversalSleevePlacerService.cs` - Places individual sleeves
- `Services/ClashZoneService.cs` - Detects intersections

## COMMON ISSUES SOLVED
- Pipe clustering not working (SleeveInstanceId = -1) → Fixed timing issue (increased wait time to 2 seconds)
- Wrong categories (Pipe Accessories vs Pipes) → Fixed category mapping
- Orientation showing "Z" → Fixed to use clash zone data
- Cache key mismatch → Use MepElementId consistently

## BEFORE ANY CLUSTERING CHANGES
1. Read `Docs/CLUSTERING_COMPLETE_GUIDE.md`
2. Check critical success factors
3. Verify architecture compliance
4. Update guide if changes are made

---
**Last Updated**: 2025-01-24
**Purpose**: Quick reference for clustering system context
