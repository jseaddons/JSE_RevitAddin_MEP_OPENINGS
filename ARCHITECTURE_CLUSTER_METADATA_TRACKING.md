# CLUSTER METADATA TRACKING FOR OPENING SCHEDULES

## 🎯 Purpose

Track MEP elements contained within each cluster opening to enable:
1. **Opening Schedules** - List all MEP elements per opening
2. **BIM Coordination** - Know what passes through each opening
3. **Construction Documentation** - Detailed opening information
4. **Clash Resolution** - Track which clashes were resolved by which cluster

---

## 📊 Required Data per Cluster

### Essential Information:

| Data Field | Source | Purpose |
|------------|--------|---------|
| **Cluster Opening ID** | Revit Element ID | Unique identifier |
| **Cluster Mark** | Instance parameter | User-facing identifier |
| **Host Element** | Wall/Floor/Framing | Where cluster is placed |
| **Location** | XYZ coordinates | Spatial reference |
| **Cluster Size** | Width × Height | Overall opening dimensions |
| **MEP Elements** | List of contained elements | What passes through |
| **System Abbreviations** | Per MEP element | Fire Protection, HVAC, etc. |
| **Element Sizes** | Per MEP element | Diameter, Width×Height |
| **Creation Date** | Timestamp | When cluster was created |
| **Individual Sleeves Replaced** | List of sleeve IDs | Traceability |

---

## 🗂️ Data Model

### ClusterMetadata.cs

```csharp
using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Metadata for a cluster opening - tracks all MEP elements contained within
    /// </summary>
    [XmlRoot("ClusterMetadata")]
    public class ClusterMetadata
    {
        /// <summary>
        /// Unique cluster identifier (Revit Element ID)
        /// </summary>
        public int ClusterId { get; set; }
        
        /// <summary>
        /// Cluster mark parameter (user-facing identifier like "CO-001")
        /// </summary>
        public string ClusterMark { get; set; } = string.Empty;
        
        /// <summary>
        /// Cluster family name (ClusterOpeningOnWallX, ClusterOpeningOnWallY, ClusterOpeningOnSlab)
        /// </summary>
        public string ClusterFamilyName { get; set; } = string.Empty;
        
        /// <summary>
        /// Host element type (Wall, Floor, Framing)
        /// </summary>
        public string HostType { get; set; } = string.Empty;
        
        /// <summary>
        /// Host element ID
        /// </summary>
        public int HostElementId { get; set; }
        
        /// <summary>
        /// Host element name/type
        /// </summary>
        public string HostName { get; set; } = string.Empty;
        
        /// <summary>
        /// Cluster center point (X, Y, Z in feet)
        /// </summary>
        public XyzPoint Location { get; set; } = new XyzPoint();
        
        /// <summary>
        /// Cluster bounding box width (in mm)
        /// </summary>
        public double ClusterWidthMm { get; set; }
        
        /// <summary>
        /// Cluster bounding box height (in mm)
        /// </summary>
        public double ClusterHeightMm { get; set; }
        
        /// <summary>
        /// Level name where cluster is located
        /// </summary>
        public string LevelName { get; set; } = string.Empty;
        
        /// <summary>
        /// When this cluster was created
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        
        /// <summary>
        /// List of MEP elements contained in this cluster
        /// </summary>
        [XmlArray("MepElements")]
        [XmlArrayItem("MepElement")]
        public List<MepElementInfo> MepElements { get; set; } = new List<MepElementInfo>();
        
        /// <summary>
        /// List of individual sleeve IDs that were replaced by this cluster
        /// </summary>
        [XmlArray("ReplacedSleeves")]
        [XmlArrayItem("SleeveId")]
        public List<int> ReplacedSleeveIds { get; set; } = new List<int>();
        
        /// <summary>
        /// Cluster tolerance used (JoinOpeningsDistance in mm)
        /// </summary>
        public double ClusterToleranceMm { get; set; }
        
        /// <summary>
        /// Number of MEP elements in this cluster
        /// </summary>
        public int MepElementCount => MepElements?.Count ?? 0;
        
        /// <summary>
        /// Total cross-sectional area of all MEP elements (in mm²)
        /// </summary>
        public double TotalMepAreaMm2
        {
            get
            {
                double total = 0;
                if (MepElements != null)
                {
                    foreach (var mep in MepElements)
                    {
                        total += mep.CrossSectionalAreaMm2;
                    }
                }
                return total;
            }
        }
        
        /// <summary>
        /// Cluster opening area (in mm²)
        /// </summary>
        public double ClusterAreaMm2 => ClusterWidthMm * ClusterHeightMm;
        
        /// <summary>
        /// Percentage of cluster area occupied by MEP elements
        /// </summary>
        public double OccupancyPercentage => ClusterAreaMm2 > 0 ? (TotalMepAreaMm2 / ClusterAreaMm2) * 100.0 : 0.0;
    }
    
    /// <summary>
    /// Information about a single MEP element within a cluster
    /// </summary>
    public class MepElementInfo
    {
        /// <summary>
        /// MEP element ID
        /// </summary>
        public int ElementId { get; set; }
        
        /// <summary>
        /// MEP element unique ID (persistent across sessions)
        /// </summary>
        public string UniqueId { get; set; } = string.Empty;
        
        /// <summary>
        /// MEP category (Ducts, Pipes, Cable Trays, Duct Accessories)
        /// </summary>
        public string Category { get; set; } = string.Empty;
        
        /// <summary>
        /// System abbreviation (HVAC, FP, DOM, etc.)
        /// </summary>
        public string SystemAbbreviation { get; set; } = string.Empty;
        
        /// <summary>
        /// System name (Supply Air, Domestic Hot Water, etc.)
        /// </summary>
        public string SystemName { get; set; } = string.Empty;
        
        /// <summary>
        /// System type (Supply Air, Return Air, Exhaust Air, etc.)
        /// </summary>
        public string SystemType { get; set; } = string.Empty;
        
        /// <summary>
        /// Element shape (Round, Rectangular, Oval)
        /// </summary>
        public string Shape { get; set; } = string.Empty;
        
        /// <summary>
        /// Width or Diameter in mm
        /// </summary>
        public double WidthMm { get; set; }
        
        /// <summary>
        /// Height in mm (for rectangular elements)
        /// </summary>
        public double HeightMm { get; set; }
        
        /// <summary>
        /// Diameter in mm (for round elements)
        /// </summary>
        public double DiameterMm { get; set; }
        
        /// <summary>
        /// Size formatted as string (e.g., "Ø300", "400×200")
        /// </summary>
        public string SizeFormatted
        {
            get
            {
                if (Shape == "Round")
                    return $"Ø{DiameterMm:F0}";
                else if (Shape == "Rectangular")
                    return $"{WidthMm:F0}×{HeightMm:F0}";
                else
                    return $"{WidthMm:F0}";
            }
        }
        
        /// <summary>
        /// Cross-sectional area in mm²
        /// </summary>
        public double CrossSectionalAreaMm2 { get; set; }
        
        /// <summary>
        /// Insulation type (None, Normal, Enhanced)
        /// </summary>
        public string InsulationType { get; set; } = "None";
        
        /// <summary>
        /// Insulation thickness in mm
        /// </summary>
        public double InsulationThicknessMm { get; set; }
        
        /// <summary>
        /// Family name
        /// </summary>
        public string FamilyName { get; set; } = string.Empty;
        
        /// <summary>
        /// Family type name
        /// </summary>
        public string FamilyTypeName { get; set; } = string.Empty;
        
        /// <summary>
        /// Comments from MEP element
        /// </summary>
        public string Comments { get; set; } = string.Empty;
        
        /// <summary>
        /// Location of this MEP element
        /// </summary>
        public XyzPoint Location { get; set; } = new XyzPoint();
    }
    
    /// <summary>
    /// Simple XYZ point for serialization
    /// </summary>
    public class XyzPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        
        public override string ToString()
        {
            return $"({X:F2}, {Y:F2}, {Z:F2})";
        }
    }
    
    /// <summary>
    /// Container for all cluster metadata
    /// </summary>
    [XmlRoot("ClusterMetadataCollection")]
    public class ClusterMetadataCollection
    {
        [XmlArray("Clusters")]
        [XmlArrayItem("Cluster")]
        public List<ClusterMetadata> Clusters { get; set; } = new List<ClusterMetadata>();
        
        /// <summary>
        /// When this collection was created/updated
        /// </summary>
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        
        /// <summary>
        /// Project name
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;
        
        /// <summary>
        /// Filter name that created these clusters
        /// </summary>
        public string FilterName { get; set; } = string.Empty;
        
        /// <summary>
        /// Total number of clusters
        /// </summary>
        public int TotalClusters => Clusters?.Count ?? 0;
        
        /// <summary>
        /// Total number of MEP elements across all clusters
        /// </summary>
        public int TotalMepElements
        {
            get
            {
                int total = 0;
                if (Clusters != null)
                {
                    foreach (var cluster in Clusters)
                    {
                        total += cluster.MepElementCount;
                    }
                }
                return total;
            }
        }
    }
}
```

---

## 🔧 Implementation in RectangularSleeveClusterCommandV2

### Add Metadata Tracking

```csharp
// In RectangularSleeveClusterCommandV2.cs

// Add field
private ClusterMetadataCollection _clusterMetadata = new ClusterMetadataCollection();

// After placing cluster instance (around line 420)
private void PlaceClusterAndTrackMetadata(
    Document doc,
    FamilySymbol clusterSymbol,
    XYZ clusterCenter,
    double clusterWidth,
    double clusterHeight,
    List<FamilyInstance> clusterSleeves,
    Element hostElement,
    string effectiveOrientation,
    string systemType)
{
    // Place cluster family instance
    var clusterInstance = doc.Create.NewFamilyInstance(
        clusterCenter, 
        clusterSymbol, 
        hostElement, 
        structuralType);
    
    // Set parameters
    SetParameterValue(clusterInstance, "Width", clusterWidth);
    SetParameterValue(clusterInstance, "Height", clusterHeight);
    SetParameterValue(clusterInstance, "HostOrientation", effectiveOrientation);
    SetParameterValue(clusterInstance, "SystemType", systemType);
    
    // ⚠️ NEW: Create metadata for this cluster
    var clusterMeta = CreateClusterMetadata(
        clusterInstance, 
        clusterSleeves, 
        hostElement, 
        clusterWidth, 
        clusterHeight);
    
    _clusterMetadata.Clusters.Add(clusterMeta);
    
    DebugLogger.Log($"[ClusterMetadata] Created metadata for cluster {clusterInstance.Id}: {clusterMeta.MepElementCount} MEP elements");
    
    placedCount++;
    
    // Delete individual sleeves
    foreach (var sleeve in clusterSleeves)
    {
        doc.Delete(sleeve.Id);
        deletedCount++;
    }
}

private ClusterMetadata CreateClusterMetadata(
    FamilyInstance clusterInstance,
    List<FamilyInstance> originalSleeves,
    Element hostElement,
    double clusterWidthFeet,
    double clusterHeightFeet)
{
    var metadata = new ClusterMetadata
    {
        ClusterId = clusterInstance.Id.IntegerValue,
        ClusterMark = GetParameterValueAsString(clusterInstance, "Mark") ?? $"CO-{clusterInstance.Id.IntegerValue}",
        ClusterFamilyName = clusterInstance.Symbol.Family.Name,
        HostType = GetHostTypeName(hostElement),
        HostElementId = hostElement.Id.IntegerValue,
        HostName = hostElement.Name ?? hostElement.Category?.Name ?? "Unknown",
        Location = new XyzPoint
        {
            X = ((LocationPoint)clusterInstance.Location).Point.X,
            Y = ((LocationPoint)clusterInstance.Location).Point.Y,
            Z = ((LocationPoint)clusterInstance.Location).Point.Z
        },
        ClusterWidthMm = UnitUtils.ConvertFromInternalUnits(clusterWidthFeet, UnitTypeId.Millimeters),
        ClusterHeightMm = UnitUtils.ConvertFromInternalUnits(clusterHeightFeet, UnitTypeId.Millimeters),
        LevelName = GetLevelName(clusterInstance),
        ClusterToleranceMm = ClusterConfigurationManager.Instance.JoinOpeningsDistance,
        ReplacedSleeveIds = originalSleeves.Select(s => s.Id.IntegerValue).ToList()
    };
    
    // Extract MEP elements from original sleeves
    foreach (var sleeve in originalSleeves)
    {
        var mepInfo = ExtractMepInfoFromSleeve(sleeve);
        if (mepInfo != null)
        {
            metadata.MepElements.Add(mepInfo);
        }
    }
    
    DebugLogger.Log($"[ClusterMetadata] Cluster {metadata.ClusterMark}: {metadata.MepElementCount} MEP elements, Total area: {metadata.TotalMepAreaMm2:F0}mm², Occupancy: {metadata.OccupancyPercentage:F1}%");
    
    return metadata;
}

private MepElementInfo ExtractMepInfoFromSleeve(FamilyInstance sleeve)
{
    try
    {
        // Get MEP element ID from sleeve parameter
        var mepIdParam = sleeve.LookupParameter("MEP_ElementId");
        if (mepIdParam == null || !mepIdParam.HasValue)
        {
            DebugLogger.Warning($"[ClusterMetadata] Sleeve {sleeve.Id} has no MEP_ElementId parameter");
            return null;
        }
        
        int mepId = mepIdParam.AsInteger();
        var mepElement = sleeve.Document.GetElement(new ElementId(mepId));
        
        if (mepElement == null)
        {
            DebugLogger.Warning($"[ClusterMetadata] MEP element {mepId} not found");
            return null;
        }
        
        var mepInfo = new MepElementInfo
        {
            ElementId = mepId,
            UniqueId = mepElement.UniqueId,
            Category = mepElement.Category?.Name ?? "Unknown",
            FamilyName = (mepElement as FamilyInstance)?.Symbol?.Family?.Name ?? "Unknown",
            FamilyTypeName = (mepElement as FamilyInstance)?.Symbol?.Name ?? mepElement.Name
        };
        
        // Extract system information
        ExtractSystemInfo(mepElement, mepInfo);
        
        // Extract size information
        ExtractSizeInfo(mepElement, mepInfo);
        
        // Extract insulation info
        ExtractInsulationInfo(mepElement, mepInfo);
        
        // Get location
        if (mepElement.Location is LocationCurve locCurve)
        {
            var midpoint = locCurve.Curve.Evaluate(0.5, true);
            mepInfo.Location = new XyzPoint { X = midpoint.X, Y = midpoint.Y, Z = midpoint.Z };
        }
        else if (mepElement.Location is LocationPoint locPoint)
        {
            mepInfo.Location = new XyzPoint { X = locPoint.Point.X, Y = locPoint.Point.Y, Z = locPoint.Point.Z };
        }
        
        // Get comments
        var commentsParam = mepElement.LookupParameter("Comments");
        mepInfo.Comments = commentsParam?.AsString() ?? string.Empty;
        
        return mepInfo;
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[ClusterMetadata] Error extracting MEP info from sleeve {sleeve.Id}: {ex.Message}");
        return null;
    }
}

private void ExtractSystemInfo(Element mepElement, MepElementInfo mepInfo)
{
    // Try to get system from MEP element
    MEPSystem mepSystem = null;
    
    if (mepElement is Duct duct)
    {
        mepSystem = duct.MEPSystem;
        mepInfo.Shape = GetDuctShape(duct);
    }
    else if (mepElement is Pipe pipe)
    {
        mepSystem = pipe.MEPSystem;
        mepInfo.Shape = "Round"; // Pipes are always round
    }
    else if (mepElement is CableTray cableTray)
    {
        mepInfo.Shape = "Rectangular"; // Cable trays are rectangular
        mepInfo.SystemName = "Cable Tray";
        mepInfo.SystemType = "Data/Power";
    }
    
    if (mepSystem != null)
    {
        mepInfo.SystemName = mepSystem.Name ?? "Unnamed System";
        mepInfo.SystemType = mepSystem.GetType().Name;
        
        // Extract system abbreviation
        var systemAbbrParam = mepSystem.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
        mepInfo.SystemAbbreviation = systemAbbrParam?.AsString() ?? GetDefaultSystemAbbreviation(mepInfo.Category);
    }
    else
    {
        mepInfo.SystemAbbreviation = GetDefaultSystemAbbreviation(mepInfo.Category);
    }
}

private void ExtractSizeInfo(Element mepElement, MepElementInfo mepInfo)
{
    if (mepElement is Duct duct)
    {
        if (mepInfo.Shape == "Round")
        {
            var diamParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            if (diamParam != null)
            {
                double diamFeet = diamParam.AsDouble();
                mepInfo.DiameterMm = UnitUtils.ConvertFromInternalUnits(diamFeet, UnitTypeId.Millimeters);
                mepInfo.WidthMm = mepInfo.DiameterMm;
                mepInfo.HeightMm = mepInfo.DiameterMm;
                mepInfo.CrossSectionalAreaMm2 = Math.PI * Math.Pow(mepInfo.DiameterMm / 2.0, 2);
            }
        }
        else // Rectangular
        {
            var widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            var heightParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            
            if (widthParam != null)
            {
                mepInfo.WidthMm = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), UnitTypeId.Millimeters);
            }
            if (heightParam != null)
            {
                mepInfo.HeightMm = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), UnitTypeId.Millimeters);
            }
            mepInfo.CrossSectionalAreaMm2 = mepInfo.WidthMm * mepInfo.HeightMm;
        }
    }
    else if (mepElement is Pipe pipe)
    {
        var diamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
        if (diamParam != null)
        {
            double diamFeet = diamParam.AsDouble();
            mepInfo.DiameterMm = UnitUtils.ConvertFromInternalUnits(diamFeet, UnitTypeId.Millimeters);
            mepInfo.WidthMm = mepInfo.DiameterMm;
            mepInfo.HeightMm = mepInfo.DiameterMm;
            mepInfo.CrossSectionalAreaMm2 = Math.PI * Math.Pow(mepInfo.DiameterMm / 2.0, 2);
        }
    }
    else if (mepElement is CableTray tray)
    {
        var widthParam = tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
        var heightParam = tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
        
        if (widthParam != null)
        {
            mepInfo.WidthMm = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), UnitTypeId.Millimeters);
        }
        if (heightParam != null)
        {
            mepInfo.HeightMm = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), UnitTypeId.Millimeters);
        }
        mepInfo.CrossSectionalAreaMm2 = mepInfo.WidthMm * mepInfo.HeightMm;
    }
}

private void ExtractInsulationInfo(Element mepElement, MepElementInfo mepInfo)
{
    var insulationParam = mepElement.LookupParameter("Insulation Thickness");
    if (insulationParam != null && insulationParam.HasValue)
    {
        double thicknessFeet = insulationParam.AsDouble();
        if (thicknessFeet > 0.001) // > 1mm
        {
            mepInfo.InsulationType = "Normal";
            mepInfo.InsulationThicknessMm = UnitUtils.ConvertFromInternalUnits(thicknessFeet, UnitTypeId.Millimeters);
        }
    }
}

private string GetDefaultSystemAbbreviation(string category)
{
    return category switch
    {
        "Ducts" => "HVAC",
        "Pipes" => "PLB",
        "Cable Trays" => "ELEC",
        "Duct Accessories" => "HVAC",
        _ => "GEN"
    };
}

// At the end of Execute method, save metadata to XML
private void SaveClusterMetadataToXml(Document doc)
{
    try
    {
        _clusterMetadata.ProjectName = doc.Title;
        _clusterMetadata.FilterName = "RectangularClusters"; // Or get from current filter
        _clusterMetadata.LastUpdated = DateTime.Now;
        
        string filtersDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "JSE_MEP_Openings", "Projects", "Default", "Filters");
        
        Directory.CreateDirectory(filtersDirectory);
        
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string xmlPath = Path.Combine(filtersDirectory, $"ClusterMetadata_{timestamp}.xml");
        
        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(ClusterMetadataCollection));
        using (var writer = new StreamWriter(xmlPath))
        {
            serializer.Serialize(writer, _clusterMetadata);
        }
        
        DebugLogger.Log($"[ClusterMetadata] ✓ Saved metadata to: {xmlPath}");
        DebugLogger.Log($"[ClusterMetadata] Total clusters: {_clusterMetadata.TotalClusters}, Total MEP elements: {_clusterMetadata.TotalMepElements}");
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[ClusterMetadata] Failed to save metadata: {ex.Message}");
    }
}
```

---

## 📋 Opening Schedule Creation

### Revit Schedule Format:

| Cluster Mark | Host Type | Level | Size (W×H) | MEP Count | Systems | Element Sizes | Total Area | Occupancy % |
|--------------|-----------|-------|------------|-----------|---------|---------------|------------|-------------|
| CO-001 | Wall (200mm) | Level 1 | 800×600 | 3 | HVAC, PLB | Ø300, Ø150, 100×100 | 480,000 mm² | 23.5% |
| CO-002 | Floor (300mm) | Level 2 | 1200×400 | 5 | HVAC, ELEC | Ø200, 400×200, 300×100, CT-200, CT-100 | 480,000 mm² | 45.2% |

### Export to Excel/CSV:

```csharp
public class ClusterScheduleExporter
{
    public void ExportToCSV(ClusterMetadataCollection metadata, string outputPath)
    {
        using (var writer = new StreamWriter(outputPath))
        {
            // Header
            writer.WriteLine("Cluster Mark,Host Type,Level,Width (mm),Height (mm),MEP Count,Systems,Element Sizes,Total MEP Area (mm²),Cluster Area (mm²),Occupancy %");
            
            // Data rows
            foreach (var cluster in metadata.Clusters)
            {
                var systems = string.Join("; ", cluster.MepElements.Select(m => m.SystemAbbreviation).Distinct());
                var sizes = string.Join("; ", cluster.MepElements.Select(m => m.SizeFormatted));
                
                writer.WriteLine($"{cluster.ClusterMark},{cluster.HostName},{cluster.LevelName},{cluster.ClusterWidthMm:F0},{cluster.ClusterHeightMm:F0},{cluster.MepElementCount},{systems},{sizes},{cluster.TotalMepAreaMm2:F0},{cluster.ClusterAreaMm2:F0},{cluster.OccupancyPercentage:F1}%");
            }
        }
    }
}
```

---

## 🔍 XML Output Example

```xml
<?xml version="1.0" encoding="utf-8"?>
<ClusterMetadataCollection>
  <LastUpdated>2025-10-07T15:30:00</LastUpdated>
  <ProjectName>Office Building Project</ProjectName>
  <FilterName>RectangularClusters</FilterName>
  <Clusters>
    <Cluster>
      <ClusterId>8756321</ClusterId>
      <ClusterMark>CO-001</ClusterMark>
      <ClusterFamilyName>ClusterOpeningOnWallX</ClusterFamilyName>
      <HostType>Wall</HostType>
      <HostElementId>2023648</HostElementId>
      <HostName>SSH-AR-WAL-BLK-SOL-200mm</HostName>
      <Location>
        <X>113.29</X>
        <Y>58.11</Y>
        <Z>44.77</Z>
      </Location>
      <ClusterWidthMm>800</ClusterWidthMm>
      <ClusterHeightMm>600</ClusterHeightMm>
      <LevelName>Level 1</LevelName>
      <CreatedDate>2025-10-07T15:30:00</CreatedDate>
      <ClusterToleranceMm>200</ClusterToleranceMm>
      <MepElements>
        <MepElement>
          <ElementId>2579789</ElementId>
          <UniqueId>abc123-def456-ghi789</UniqueId>
          <Category>Ducts</Category>
          <SystemAbbreviation>HVAC</SystemAbbreviation>
          <SystemName>Supply Air</SystemName>
          <SystemType>Supply Air</SystemType>
          <Shape>Round</Shape>
          <DiameterMm>300</DiameterMm>
          <WidthMm>300</WidthMm>
          <HeightMm>300</HeightMm>
          <CrossSectionalAreaMm2>70686</CrossSectionalAreaMm2>
          <InsulationType>Normal</InsulationType>
          <InsulationThicknessMm>25</InsulationThicknessMm>
          <FamilyName>D30-Duct_Round-G2</FamilyName>
          <FamilyTypeName>Ø300</FamilyTypeName>
          <Location>
            <X>113.29</X>
            <Y>58.11</Y>
            <Z>44.77</Z>
          </Location>
        </MepElement>
        <MepElement>
          <ElementId>2579790</ElementId>
          <UniqueId>xyz789-uvw456-rst123</UniqueId>
          <Category>Pipes</Category>
          <SystemAbbreviation>PLB</SystemAbbreviation>
          <SystemName>Domestic Hot Water</SystemName>
          <SystemType>Domestic Hot Water</SystemType>
          <Shape>Round</Shape>
          <DiameterMm>50</DiameterMm>
          <WidthMm>50</WidthMm>
          <HeightMm>50</HeightMm>
          <CrossSectionalAreaMm2>1963</CrossSectionalAreaMm2>
          <InsulationType>Enhanced</InsulationType>
          <InsulationThicknessMm>40</InsulationThicknessMm>
          <FamilyName>Pipe-Standard</FamilyName>
          <FamilyTypeName>50mm</FamilyTypeName>
          <Location>
            <X>113.35</X>
            <Y>58.15</Y>
            <Z>44.80</Z>
          </Location>
        </MepElement>
      </MepElements>
      <ReplacedSleeves>
        <SleeveId>8756301</SleeveId>
        <SleeveId>8756302</SleeveId>
      </ReplacedSleeves>
    </Cluster>
  </Clusters>
</ClusterMetadataCollection>
```

---

## ✅ Summary

**This solution provides:**

1. ✅ **Complete MEP element tracking** per cluster
2. ✅ **System abbreviations** (HVAC, PLB, ELEC, etc.)
3. ✅ **Element sizes** (formatted for readability)
4. ✅ **XML persistence** for opening schedules
5. ✅ **Export to CSV/Excel** for coordination
6. ✅ **Occupancy percentage** for structural review
7. ✅ **Traceability** - which sleeves were replaced

**Ready to create comprehensive opening schedules from cluster metadata!** 📊



## 🎯 Purpose

Track MEP elements contained within each cluster opening to enable:
1. **Opening Schedules** - List all MEP elements per opening
2. **BIM Coordination** - Know what passes through each opening
3. **Construction Documentation** - Detailed opening information
4. **Clash Resolution** - Track which clashes were resolved by which cluster

---

## 📊 Required Data per Cluster

### Essential Information:

| Data Field | Source | Purpose |
|------------|--------|---------|
| **Cluster Opening ID** | Revit Element ID | Unique identifier |
| **Cluster Mark** | Instance parameter | User-facing identifier |
| **Host Element** | Wall/Floor/Framing | Where cluster is placed |
| **Location** | XYZ coordinates | Spatial reference |
| **Cluster Size** | Width × Height | Overall opening dimensions |
| **MEP Elements** | List of contained elements | What passes through |
| **System Abbreviations** | Per MEP element | Fire Protection, HVAC, etc. |
| **Element Sizes** | Per MEP element | Diameter, Width×Height |
| **Creation Date** | Timestamp | When cluster was created |
| **Individual Sleeves Replaced** | List of sleeve IDs | Traceability |

---

## 🗂️ Data Model

### ClusterMetadata.cs

```csharp
using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace JSE_RevitAddin_MEP_OPENINGS.Models
{
    /// <summary>
    /// Metadata for a cluster opening - tracks all MEP elements contained within
    /// </summary>
    [XmlRoot("ClusterMetadata")]
    public class ClusterMetadata
    {
        /// <summary>
        /// Unique cluster identifier (Revit Element ID)
        /// </summary>
        public int ClusterId { get; set; }
        
        /// <summary>
        /// Cluster mark parameter (user-facing identifier like "CO-001")
        /// </summary>
        public string ClusterMark { get; set; } = string.Empty;
        
        /// <summary>
        /// Cluster family name (ClusterOpeningOnWallX, ClusterOpeningOnWallY, ClusterOpeningOnSlab)
        /// </summary>
        public string ClusterFamilyName { get; set; } = string.Empty;
        
        /// <summary>
        /// Host element type (Wall, Floor, Framing)
        /// </summary>
        public string HostType { get; set; } = string.Empty;
        
        /// <summary>
        /// Host element ID
        /// </summary>
        public int HostElementId { get; set; }
        
        /// <summary>
        /// Host element name/type
        /// </summary>
        public string HostName { get; set; } = string.Empty;
        
        /// <summary>
        /// Cluster center point (X, Y, Z in feet)
        /// </summary>
        public XyzPoint Location { get; set; } = new XyzPoint();
        
        /// <summary>
        /// Cluster bounding box width (in mm)
        /// </summary>
        public double ClusterWidthMm { get; set; }
        
        /// <summary>
        /// Cluster bounding box height (in mm)
        /// </summary>
        public double ClusterHeightMm { get; set; }
        
        /// <summary>
        /// Level name where cluster is located
        /// </summary>
        public string LevelName { get; set; } = string.Empty;
        
        /// <summary>
        /// When this cluster was created
        /// </summary>
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        
        /// <summary>
        /// List of MEP elements contained in this cluster
        /// </summary>
        [XmlArray("MepElements")]
        [XmlArrayItem("MepElement")]
        public List<MepElementInfo> MepElements { get; set; } = new List<MepElementInfo>();
        
        /// <summary>
        /// List of individual sleeve IDs that were replaced by this cluster
        /// </summary>
        [XmlArray("ReplacedSleeves")]
        [XmlArrayItem("SleeveId")]
        public List<int> ReplacedSleeveIds { get; set; } = new List<int>();
        
        /// <summary>
        /// Cluster tolerance used (JoinOpeningsDistance in mm)
        /// </summary>
        public double ClusterToleranceMm { get; set; }
        
        /// <summary>
        /// Number of MEP elements in this cluster
        /// </summary>
        public int MepElementCount => MepElements?.Count ?? 0;
        
        /// <summary>
        /// Total cross-sectional area of all MEP elements (in mm²)
        /// </summary>
        public double TotalMepAreaMm2
        {
            get
            {
                double total = 0;
                if (MepElements != null)
                {
                    foreach (var mep in MepElements)
                    {
                        total += mep.CrossSectionalAreaMm2;
                    }
                }
                return total;
            }
        }
        
        /// <summary>
        /// Cluster opening area (in mm²)
        /// </summary>
        public double ClusterAreaMm2 => ClusterWidthMm * ClusterHeightMm;
        
        /// <summary>
        /// Percentage of cluster area occupied by MEP elements
        /// </summary>
        public double OccupancyPercentage => ClusterAreaMm2 > 0 ? (TotalMepAreaMm2 / ClusterAreaMm2) * 100.0 : 0.0;
    }
    
    /// <summary>
    /// Information about a single MEP element within a cluster
    /// </summary>
    public class MepElementInfo
    {
        /// <summary>
        /// MEP element ID
        /// </summary>
        public int ElementId { get; set; }
        
        /// <summary>
        /// MEP element unique ID (persistent across sessions)
        /// </summary>
        public string UniqueId { get; set; } = string.Empty;
        
        /// <summary>
        /// MEP category (Ducts, Pipes, Cable Trays, Duct Accessories)
        /// </summary>
        public string Category { get; set; } = string.Empty;
        
        /// <summary>
        /// System abbreviation (HVAC, FP, DOM, etc.)
        /// </summary>
        public string SystemAbbreviation { get; set; } = string.Empty;
        
        /// <summary>
        /// System name (Supply Air, Domestic Hot Water, etc.)
        /// </summary>
        public string SystemName { get; set; } = string.Empty;
        
        /// <summary>
        /// System type (Supply Air, Return Air, Exhaust Air, etc.)
        /// </summary>
        public string SystemType { get; set; } = string.Empty;
        
        /// <summary>
        /// Element shape (Round, Rectangular, Oval)
        /// </summary>
        public string Shape { get; set; } = string.Empty;
        
        /// <summary>
        /// Width or Diameter in mm
        /// </summary>
        public double WidthMm { get; set; }
        
        /// <summary>
        /// Height in mm (for rectangular elements)
        /// </summary>
        public double HeightMm { get; set; }
        
        /// <summary>
        /// Diameter in mm (for round elements)
        /// </summary>
        public double DiameterMm { get; set; }
        
        /// <summary>
        /// Size formatted as string (e.g., "Ø300", "400×200")
        /// </summary>
        public string SizeFormatted
        {
            get
            {
                if (Shape == "Round")
                    return $"Ø{DiameterMm:F0}";
                else if (Shape == "Rectangular")
                    return $"{WidthMm:F0}×{HeightMm:F0}";
                else
                    return $"{WidthMm:F0}";
            }
        }
        
        /// <summary>
        /// Cross-sectional area in mm²
        /// </summary>
        public double CrossSectionalAreaMm2 { get; set; }
        
        /// <summary>
        /// Insulation type (None, Normal, Enhanced)
        /// </summary>
        public string InsulationType { get; set; } = "None";
        
        /// <summary>
        /// Insulation thickness in mm
        /// </summary>
        public double InsulationThicknessMm { get; set; }
        
        /// <summary>
        /// Family name
        /// </summary>
        public string FamilyName { get; set; } = string.Empty;
        
        /// <summary>
        /// Family type name
        /// </summary>
        public string FamilyTypeName { get; set; } = string.Empty;
        
        /// <summary>
        /// Comments from MEP element
        /// </summary>
        public string Comments { get; set; } = string.Empty;
        
        /// <summary>
        /// Location of this MEP element
        /// </summary>
        public XyzPoint Location { get; set; } = new XyzPoint();
    }
    
    /// <summary>
    /// Simple XYZ point for serialization
    /// </summary>
    public class XyzPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        
        public override string ToString()
        {
            return $"({X:F2}, {Y:F2}, {Z:F2})";
        }
    }
    
    /// <summary>
    /// Container for all cluster metadata
    /// </summary>
    [XmlRoot("ClusterMetadataCollection")]
    public class ClusterMetadataCollection
    {
        [XmlArray("Clusters")]
        [XmlArrayItem("Cluster")]
        public List<ClusterMetadata> Clusters { get; set; } = new List<ClusterMetadata>();
        
        /// <summary>
        /// When this collection was created/updated
        /// </summary>
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        
        /// <summary>
        /// Project name
        /// </summary>
        public string ProjectName { get; set; } = string.Empty;
        
        /// <summary>
        /// Filter name that created these clusters
        /// </summary>
        public string FilterName { get; set; } = string.Empty;
        
        /// <summary>
        /// Total number of clusters
        /// </summary>
        public int TotalClusters => Clusters?.Count ?? 0;
        
        /// <summary>
        /// Total number of MEP elements across all clusters
        /// </summary>
        public int TotalMepElements
        {
            get
            {
                int total = 0;
                if (Clusters != null)
                {
                    foreach (var cluster in Clusters)
                    {
                        total += cluster.MepElementCount;
                    }
                }
                return total;
            }
        }
    }
}
```

---

## 🔧 Implementation in RectangularSleeveClusterCommandV2

### Add Metadata Tracking

```csharp
// In RectangularSleeveClusterCommandV2.cs

// Add field
private ClusterMetadataCollection _clusterMetadata = new ClusterMetadataCollection();

// After placing cluster instance (around line 420)
private void PlaceClusterAndTrackMetadata(
    Document doc,
    FamilySymbol clusterSymbol,
    XYZ clusterCenter,
    double clusterWidth,
    double clusterHeight,
    List<FamilyInstance> clusterSleeves,
    Element hostElement,
    string effectiveOrientation,
    string systemType)
{
    // Place cluster family instance
    var clusterInstance = doc.Create.NewFamilyInstance(
        clusterCenter, 
        clusterSymbol, 
        hostElement, 
        structuralType);
    
    // Set parameters
    SetParameterValue(clusterInstance, "Width", clusterWidth);
    SetParameterValue(clusterInstance, "Height", clusterHeight);
    SetParameterValue(clusterInstance, "HostOrientation", effectiveOrientation);
    SetParameterValue(clusterInstance, "SystemType", systemType);
    
    // ⚠️ NEW: Create metadata for this cluster
    var clusterMeta = CreateClusterMetadata(
        clusterInstance, 
        clusterSleeves, 
        hostElement, 
        clusterWidth, 
        clusterHeight);
    
    _clusterMetadata.Clusters.Add(clusterMeta);
    
    DebugLogger.Log($"[ClusterMetadata] Created metadata for cluster {clusterInstance.Id}: {clusterMeta.MepElementCount} MEP elements");
    
    placedCount++;
    
    // Delete individual sleeves
    foreach (var sleeve in clusterSleeves)
    {
        doc.Delete(sleeve.Id);
        deletedCount++;
    }
}

private ClusterMetadata CreateClusterMetadata(
    FamilyInstance clusterInstance,
    List<FamilyInstance> originalSleeves,
    Element hostElement,
    double clusterWidthFeet,
    double clusterHeightFeet)
{
    var metadata = new ClusterMetadata
    {
        ClusterId = clusterInstance.Id.IntegerValue,
        ClusterMark = GetParameterValueAsString(clusterInstance, "Mark") ?? $"CO-{clusterInstance.Id.IntegerValue}",
        ClusterFamilyName = clusterInstance.Symbol.Family.Name,
        HostType = GetHostTypeName(hostElement),
        HostElementId = hostElement.Id.IntegerValue,
        HostName = hostElement.Name ?? hostElement.Category?.Name ?? "Unknown",
        Location = new XyzPoint
        {
            X = ((LocationPoint)clusterInstance.Location).Point.X,
            Y = ((LocationPoint)clusterInstance.Location).Point.Y,
            Z = ((LocationPoint)clusterInstance.Location).Point.Z
        },
        ClusterWidthMm = UnitUtils.ConvertFromInternalUnits(clusterWidthFeet, UnitTypeId.Millimeters),
        ClusterHeightMm = UnitUtils.ConvertFromInternalUnits(clusterHeightFeet, UnitTypeId.Millimeters),
        LevelName = GetLevelName(clusterInstance),
        ClusterToleranceMm = ClusterConfigurationManager.Instance.JoinOpeningsDistance,
        ReplacedSleeveIds = originalSleeves.Select(s => s.Id.IntegerValue).ToList()
    };
    
    // Extract MEP elements from original sleeves
    foreach (var sleeve in originalSleeves)
    {
        var mepInfo = ExtractMepInfoFromSleeve(sleeve);
        if (mepInfo != null)
        {
            metadata.MepElements.Add(mepInfo);
        }
    }
    
    DebugLogger.Log($"[ClusterMetadata] Cluster {metadata.ClusterMark}: {metadata.MepElementCount} MEP elements, Total area: {metadata.TotalMepAreaMm2:F0}mm², Occupancy: {metadata.OccupancyPercentage:F1}%");
    
    return metadata;
}

private MepElementInfo ExtractMepInfoFromSleeve(FamilyInstance sleeve)
{
    try
    {
        // Get MEP element ID from sleeve parameter
        var mepIdParam = sleeve.LookupParameter("MEP_ElementId");
        if (mepIdParam == null || !mepIdParam.HasValue)
        {
            DebugLogger.Warning($"[ClusterMetadata] Sleeve {sleeve.Id} has no MEP_ElementId parameter");
            return null;
        }
        
        int mepId = mepIdParam.AsInteger();
        var mepElement = sleeve.Document.GetElement(new ElementId(mepId));
        
        if (mepElement == null)
        {
            DebugLogger.Warning($"[ClusterMetadata] MEP element {mepId} not found");
            return null;
        }
        
        var mepInfo = new MepElementInfo
        {
            ElementId = mepId,
            UniqueId = mepElement.UniqueId,
            Category = mepElement.Category?.Name ?? "Unknown",
            FamilyName = (mepElement as FamilyInstance)?.Symbol?.Family?.Name ?? "Unknown",
            FamilyTypeName = (mepElement as FamilyInstance)?.Symbol?.Name ?? mepElement.Name
        };
        
        // Extract system information
        ExtractSystemInfo(mepElement, mepInfo);
        
        // Extract size information
        ExtractSizeInfo(mepElement, mepInfo);
        
        // Extract insulation info
        ExtractInsulationInfo(mepElement, mepInfo);
        
        // Get location
        if (mepElement.Location is LocationCurve locCurve)
        {
            var midpoint = locCurve.Curve.Evaluate(0.5, true);
            mepInfo.Location = new XyzPoint { X = midpoint.X, Y = midpoint.Y, Z = midpoint.Z };
        }
        else if (mepElement.Location is LocationPoint locPoint)
        {
            mepInfo.Location = new XyzPoint { X = locPoint.Point.X, Y = locPoint.Point.Y, Z = locPoint.Point.Z };
        }
        
        // Get comments
        var commentsParam = mepElement.LookupParameter("Comments");
        mepInfo.Comments = commentsParam?.AsString() ?? string.Empty;
        
        return mepInfo;
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[ClusterMetadata] Error extracting MEP info from sleeve {sleeve.Id}: {ex.Message}");
        return null;
    }
}

private void ExtractSystemInfo(Element mepElement, MepElementInfo mepInfo)
{
    // Try to get system from MEP element
    MEPSystem mepSystem = null;
    
    if (mepElement is Duct duct)
    {
        mepSystem = duct.MEPSystem;
        mepInfo.Shape = GetDuctShape(duct);
    }
    else if (mepElement is Pipe pipe)
    {
        mepSystem = pipe.MEPSystem;
        mepInfo.Shape = "Round"; // Pipes are always round
    }
    else if (mepElement is CableTray cableTray)
    {
        mepInfo.Shape = "Rectangular"; // Cable trays are rectangular
        mepInfo.SystemName = "Cable Tray";
        mepInfo.SystemType = "Data/Power";
    }
    
    if (mepSystem != null)
    {
        mepInfo.SystemName = mepSystem.Name ?? "Unnamed System";
        mepInfo.SystemType = mepSystem.GetType().Name;
        
        // Extract system abbreviation
        var systemAbbrParam = mepSystem.get_Parameter(BuiltInParameter.RBS_SYSTEM_ABBREVIATION_PARAM);
        mepInfo.SystemAbbreviation = systemAbbrParam?.AsString() ?? GetDefaultSystemAbbreviation(mepInfo.Category);
    }
    else
    {
        mepInfo.SystemAbbreviation = GetDefaultSystemAbbreviation(mepInfo.Category);
    }
}

private void ExtractSizeInfo(Element mepElement, MepElementInfo mepInfo)
{
    if (mepElement is Duct duct)
    {
        if (mepInfo.Shape == "Round")
        {
            var diamParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
            if (diamParam != null)
            {
                double diamFeet = diamParam.AsDouble();
                mepInfo.DiameterMm = UnitUtils.ConvertFromInternalUnits(diamFeet, UnitTypeId.Millimeters);
                mepInfo.WidthMm = mepInfo.DiameterMm;
                mepInfo.HeightMm = mepInfo.DiameterMm;
                mepInfo.CrossSectionalAreaMm2 = Math.PI * Math.Pow(mepInfo.DiameterMm / 2.0, 2);
            }
        }
        else // Rectangular
        {
            var widthParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
            var heightParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
            
            if (widthParam != null)
            {
                mepInfo.WidthMm = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), UnitTypeId.Millimeters);
            }
            if (heightParam != null)
            {
                mepInfo.HeightMm = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), UnitTypeId.Millimeters);
            }
            mepInfo.CrossSectionalAreaMm2 = mepInfo.WidthMm * mepInfo.HeightMm;
        }
    }
    else if (mepElement is Pipe pipe)
    {
        var diamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);
        if (diamParam != null)
        {
            double diamFeet = diamParam.AsDouble();
            mepInfo.DiameterMm = UnitUtils.ConvertFromInternalUnits(diamFeet, UnitTypeId.Millimeters);
            mepInfo.WidthMm = mepInfo.DiameterMm;
            mepInfo.HeightMm = mepInfo.DiameterMm;
            mepInfo.CrossSectionalAreaMm2 = Math.PI * Math.Pow(mepInfo.DiameterMm / 2.0, 2);
        }
    }
    else if (mepElement is CableTray tray)
    {
        var widthParam = tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM);
        var heightParam = tray.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM);
        
        if (widthParam != null)
        {
            mepInfo.WidthMm = UnitUtils.ConvertFromInternalUnits(widthParam.AsDouble(), UnitTypeId.Millimeters);
        }
        if (heightParam != null)
        {
            mepInfo.HeightMm = UnitUtils.ConvertFromInternalUnits(heightParam.AsDouble(), UnitTypeId.Millimeters);
        }
        mepInfo.CrossSectionalAreaMm2 = mepInfo.WidthMm * mepInfo.HeightMm;
    }
}

private void ExtractInsulationInfo(Element mepElement, MepElementInfo mepInfo)
{
    var insulationParam = mepElement.LookupParameter("Insulation Thickness");
    if (insulationParam != null && insulationParam.HasValue)
    {
        double thicknessFeet = insulationParam.AsDouble();
        if (thicknessFeet > 0.001) // > 1mm
        {
            mepInfo.InsulationType = "Normal";
            mepInfo.InsulationThicknessMm = UnitUtils.ConvertFromInternalUnits(thicknessFeet, UnitTypeId.Millimeters);
        }
    }
}

private string GetDefaultSystemAbbreviation(string category)
{
    return category switch
    {
        "Ducts" => "HVAC",
        "Pipes" => "PLB",
        "Cable Trays" => "ELEC",
        "Duct Accessories" => "HVAC",
        _ => "GEN"
    };
}

// At the end of Execute method, save metadata to XML
private void SaveClusterMetadataToXml(Document doc)
{
    try
    {
        _clusterMetadata.ProjectName = doc.Title;
        _clusterMetadata.FilterName = "RectangularClusters"; // Or get from current filter
        _clusterMetadata.LastUpdated = DateTime.Now;
        
        string filtersDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "JSE_MEP_Openings", "Projects", "Default", "Filters");
        
        Directory.CreateDirectory(filtersDirectory);
        
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string xmlPath = Path.Combine(filtersDirectory, $"ClusterMetadata_{timestamp}.xml");
        
        var serializer = new System.Xml.Serialization.XmlSerializer(typeof(ClusterMetadataCollection));
        using (var writer = new StreamWriter(xmlPath))
        {
            serializer.Serialize(writer, _clusterMetadata);
        }
        
        DebugLogger.Log($"[ClusterMetadata] ✓ Saved metadata to: {xmlPath}");
        DebugLogger.Log($"[ClusterMetadata] Total clusters: {_clusterMetadata.TotalClusters}, Total MEP elements: {_clusterMetadata.TotalMepElements}");
    }
    catch (Exception ex)
    {
        DebugLogger.Error($"[ClusterMetadata] Failed to save metadata: {ex.Message}");
    }
}
```

---

## 📋 Opening Schedule Creation

### Revit Schedule Format:

| Cluster Mark | Host Type | Level | Size (W×H) | MEP Count | Systems | Element Sizes | Total Area | Occupancy % |
|--------------|-----------|-------|------------|-----------|---------|---------------|------------|-------------|
| CO-001 | Wall (200mm) | Level 1 | 800×600 | 3 | HVAC, PLB | Ø300, Ø150, 100×100 | 480,000 mm² | 23.5% |
| CO-002 | Floor (300mm) | Level 2 | 1200×400 | 5 | HVAC, ELEC | Ø200, 400×200, 300×100, CT-200, CT-100 | 480,000 mm² | 45.2% |

### Export to Excel/CSV:

```csharp
public class ClusterScheduleExporter
{
    public void ExportToCSV(ClusterMetadataCollection metadata, string outputPath)
    {
        using (var writer = new StreamWriter(outputPath))
        {
            // Header
            writer.WriteLine("Cluster Mark,Host Type,Level,Width (mm),Height (mm),MEP Count,Systems,Element Sizes,Total MEP Area (mm²),Cluster Area (mm²),Occupancy %");
            
            // Data rows
            foreach (var cluster in metadata.Clusters)
            {
                var systems = string.Join("; ", cluster.MepElements.Select(m => m.SystemAbbreviation).Distinct());
                var sizes = string.Join("; ", cluster.MepElements.Select(m => m.SizeFormatted));
                
                writer.WriteLine($"{cluster.ClusterMark},{cluster.HostName},{cluster.LevelName},{cluster.ClusterWidthMm:F0},{cluster.ClusterHeightMm:F0},{cluster.MepElementCount},{systems},{sizes},{cluster.TotalMepAreaMm2:F0},{cluster.ClusterAreaMm2:F0},{cluster.OccupancyPercentage:F1}%");
            }
        }
    }
}
```

---

## 🔍 XML Output Example

```xml
<?xml version="1.0" encoding="utf-8"?>
<ClusterMetadataCollection>
  <LastUpdated>2025-10-07T15:30:00</LastUpdated>
  <ProjectName>Office Building Project</ProjectName>
  <FilterName>RectangularClusters</FilterName>
  <Clusters>
    <Cluster>
      <ClusterId>8756321</ClusterId>
      <ClusterMark>CO-001</ClusterMark>
      <ClusterFamilyName>ClusterOpeningOnWallX</ClusterFamilyName>
      <HostType>Wall</HostType>
      <HostElementId>2023648</HostElementId>
      <HostName>SSH-AR-WAL-BLK-SOL-200mm</HostName>
      <Location>
        <X>113.29</X>
        <Y>58.11</Y>
        <Z>44.77</Z>
      </Location>
      <ClusterWidthMm>800</ClusterWidthMm>
      <ClusterHeightMm>600</ClusterHeightMm>
      <LevelName>Level 1</LevelName>
      <CreatedDate>2025-10-07T15:30:00</CreatedDate>
      <ClusterToleranceMm>200</ClusterToleranceMm>
      <MepElements>
        <MepElement>
          <ElementId>2579789</ElementId>
          <UniqueId>abc123-def456-ghi789</UniqueId>
          <Category>Ducts</Category>
          <SystemAbbreviation>HVAC</SystemAbbreviation>
          <SystemName>Supply Air</SystemName>
          <SystemType>Supply Air</SystemType>
          <Shape>Round</Shape>
          <DiameterMm>300</DiameterMm>
          <WidthMm>300</WidthMm>
          <HeightMm>300</HeightMm>
          <CrossSectionalAreaMm2>70686</CrossSectionalAreaMm2>
          <InsulationType>Normal</InsulationType>
          <InsulationThicknessMm>25</InsulationThicknessMm>
          <FamilyName>D30-Duct_Round-G2</FamilyName>
          <FamilyTypeName>Ø300</FamilyTypeName>
          <Location>
            <X>113.29</X>
            <Y>58.11</Y>
            <Z>44.77</Z>
          </Location>
        </MepElement>
        <MepElement>
          <ElementId>2579790</ElementId>
          <UniqueId>xyz789-uvw456-rst123</UniqueId>
          <Category>Pipes</Category>
          <SystemAbbreviation>PLB</SystemAbbreviation>
          <SystemName>Domestic Hot Water</SystemName>
          <SystemType>Domestic Hot Water</SystemType>
          <Shape>Round</Shape>
          <DiameterMm>50</DiameterMm>
          <WidthMm>50</WidthMm>
          <HeightMm>50</HeightMm>
          <CrossSectionalAreaMm2>1963</CrossSectionalAreaMm2>
          <InsulationType>Enhanced</InsulationType>
          <InsulationThicknessMm>40</InsulationThicknessMm>
          <FamilyName>Pipe-Standard</FamilyName>
          <FamilyTypeName>50mm</FamilyTypeName>
          <Location>
            <X>113.35</X>
            <Y>58.15</Y>
            <Z>44.80</Z>
          </Location>
        </MepElement>
      </MepElements>
      <ReplacedSleeves>
        <SleeveId>8756301</SleeveId>
        <SleeveId>8756302</SleeveId>
      </ReplacedSleeves>
    </Cluster>
  </Clusters>
</ClusterMetadataCollection>
```

---

## ✅ Summary

**This solution provides:**

1. ✅ **Complete MEP element tracking** per cluster
2. ✅ **System abbreviations** (HVAC, PLB, ELEC, etc.)
3. ✅ **Element sizes** (formatted for readability)
4. ✅ **XML persistence** for opening schedules
5. ✅ **Export to CSV/Excel** for coordination
6. ✅ **Occupancy percentage** for structural review
7. ✅ **Traceability** - which sleeves were replaced

**Ready to create comprehensive opening schedules from cluster metadata!** 📊



















