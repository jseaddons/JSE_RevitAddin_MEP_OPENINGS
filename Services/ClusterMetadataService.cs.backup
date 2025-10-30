using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for creating and managing cluster metadata for opening schedules
    /// </summary>
    public class ClusterMetadataService
    {
        private readonly Document _document;
        
        public ClusterMetadataService(Document document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }
        
        /// <summary>
        /// Create metadata for a cluster opening
        /// </summary>
        public ClusterMetadata CreateClusterMetadata(
            FamilyInstance clusterInstance,
            List<FamilyInstance> originalSleeves,
            Element hostElement,
            double clusterWidthFeet,
            double clusterHeightFeet,
            string category)
        {
            var metadata = new ClusterMetadata
            {
                ClusterId = clusterInstance.Id.IntegerValue,
                ClusterMark = GetParameterValueAsString(clusterInstance, "Mark") ?? $"CO-{clusterInstance.Id.IntegerValue}",
                ClusterFamilyName = clusterInstance.Symbol.Family.Name,
                MepCategory = category, // Single category (Ducts, Pipes, Cable Trays, Duct Accessories)
                HostType = GetHostTypeName(hostElement),
                HostElementId = hostElement.Id.IntegerValue,
                HostName = hostElement.Name ?? hostElement.Category?.Name ?? "Unknown",
                Location = GetXyzPoint((clusterInstance.Location as LocationPoint)?.Point),
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
        
        /// <summary>
        /// Extract MEP element information from a sleeve instance
        /// </summary>
        private MepElementInfo ExtractMepInfoFromSleeve(FamilyInstance sleeve)
        {
            try
            {
                // Get MEP element ID from sleeve parameter (adjust parameter name as needed)
                var mepIdParam = sleeve.LookupParameter("MEP_ElementId") 
                              ?? sleeve.LookupParameter("SourceElementId")
                              ?? sleeve.LookupParameter("MepElementId");
                
                if (mepIdParam == null || !mepIdParam.HasValue)
                {
                    DebugLogger.Warning($"[ClusterMetadata] Sleeve {sleeve.Id} has no MEP element ID parameter");
                    return null;
                }
                
                int mepId = mepIdParam.AsInteger();
                var mepElement = _document.GetElement(new ElementId(mepId));
                
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
                    mepInfo.Location = GetXyzPoint(midpoint);
                }
                else if (mepElement.Location is LocationPoint locPoint)
                {
                    mepInfo.Location = GetXyzPoint(locPoint.Point);
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
        
        /// <summary>
        /// Extract system information from MEP element
        /// </summary>
        private void ExtractSystemInfo(Element mepElement, MepElementInfo mepInfo)
        {
            MEPSystem mepSystem = null;
            
            if (mepElement is Duct duct)
            {
                mepSystem = duct.MEPSystem;
                mepInfo.Shape = GetDuctShape(duct);
            }
            else if (mepElement is Pipe pipe)
            {
                mepSystem = pipe.MEPSystem;
                mepInfo.Shape = "Round";
            }
            else if (mepElement is CableTray cableTray)
            {
                mepInfo.Shape = "Rectangular";
                mepInfo.SystemName = "Cable Tray";
                mepInfo.SystemType = "Data/Power";
            }
            else if (mepElement is FamilyInstance famInst && famInst.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_DuctAccessory)
            {
                // Duct accessory (damper)
                mepInfo.Shape = "Rectangular"; // Most dampers are rectangular
                mepInfo.SystemName = "Duct Accessory";
                mepInfo.SystemType = "Damper";
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
        
        /// <summary>
        /// Extract size information from MEP element
        /// </summary>
        private void ExtractSizeInfo(Element mepElement, MepElementInfo mepInfo)
        {
            if (mepElement is Duct duct)
            {
                if (mepInfo.Shape == "Round" || mepInfo.Shape == "Circular")
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
        
        /// <summary>
        /// Extract insulation information
        /// </summary>
        private void ExtractInsulationInfo(Element mepElement, MepElementInfo mepInfo)
        {
            var insulationParam = mepElement.LookupParameter("Insulation Thickness") 
                               ?? mepElement.LookupParameter("InsulationThickness");
            
            if (insulationParam != null && insulationParam.HasValue)
            {
                double thicknessFeet = insulationParam.AsDouble();
                if (thicknessFeet > 0.001) // > ~0.3mm
                {
                    mepInfo.InsulationType = "Normal";
                    mepInfo.InsulationThicknessMm = UnitUtils.ConvertFromInternalUnits(thicknessFeet, UnitTypeId.Millimeters);
                }
            }
        }
        
        /// <summary>
        /// Save cluster metadata collection to XML file
        /// </summary>
        public void SaveClusterMetadata(ClusterMetadataCollection metadata, string filterName, string category)
        {
            try
            {
                metadata.ProjectName = _document.Title;
                metadata.FilterName = filterName;
                metadata.Category = category;
                metadata.LastUpdated = DateTime.Now;
                
                string filtersDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "JSE_MEP_Openings", "Projects", "Default", "Filters");
                
                Directory.CreateDirectory(filtersDirectory);
                
                // Fixed naming: {FilterName}_{Category}_CLUSTER.xml
                string xmlFileName = $"{filterName}_{category}_CLUSTER.xml";
                string xmlPath = Path.Combine(filtersDirectory, xmlFileName);
                
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(ClusterMetadataCollection));
                using (var writer = new StreamWriter(xmlPath))
                {
                    serializer.Serialize(writer, metadata);
                }
                
                DebugLogger.Log($"[ClusterMetadata] ✓ Saved metadata to: {xmlPath}");
                DebugLogger.Log($"[ClusterMetadata] Total clusters: {metadata.TotalClusters}, Total MEP elements: {metadata.TotalMepElements}");
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[ClusterMetadata] Failed to save metadata: {ex.Message}");
            }
        }
        
        // Helper methods
        
        private string GetDuctShape(Duct duct)
        {
            try
            {
                var diamParam = duct.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM);
                if (diamParam != null && diamParam.HasValue && diamParam.AsDouble() > 0.001)
                {
                    return "Round";
                }
                return "Rectangular";
            }
            catch
            {
                return "Rectangular";
            }
        }
        
        private string GetHostTypeName(Element hostElement)
        {
            if (hostElement == null) return "Unknown";
            
            var category = hostElement.Category;
            if (category == null) return "Unknown";
            
            return category.Id.IntegerValue switch
            {
                (int)BuiltInCategory.OST_Walls => "Wall",
                (int)BuiltInCategory.OST_Floors => "Floor",
                (int)BuiltInCategory.OST_StructuralFraming => "Framing",
                _ => category.Name ?? "Unknown"
            };
        }
        
        private string GetLevelName(FamilyInstance instance)
        {
            try
            {
                var levelParam = instance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
                              ?? instance.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                
                if (levelParam != null && levelParam.HasValue)
                {
                    var levelId = levelParam.AsElementId();
                    var level = _document.GetElement(levelId) as Level;
                    return level?.Name ?? "Unknown";
                }
                
                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }
        
        private string GetParameterValueAsString(Element element, string parameterName)
        {
            try
            {
                var param = element.LookupParameter(parameterName);
                if (param != null && param.HasValue)
                {
                    return param.AsString() ?? param.AsValueString() ?? string.Empty;
                }
                return null;
            }
            catch
            {
                return null;
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
        
        private XyzPoint GetXyzPoint(XYZ xyz)
        {
            if (xyz == null) return new XyzPoint();
            return new XyzPoint { X = xyz.X, Y = xyz.Y, Z = xyz.Z };
        }
    }
}
