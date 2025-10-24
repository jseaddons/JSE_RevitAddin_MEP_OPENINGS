using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service to handle sleeve data operations for clustering
    /// </summary>
    public class SleeveDataService
    {
        private readonly Document _doc;
        
        public SleeveDataService(Document doc)
        {
            _doc = doc;
        }
        
        /// <summary>
        /// Save sleeve data to _CLUSTER.xml file
        /// </summary>
        public void SaveSleeveDataToClusterXml(List<SleeveData> sleeveDataList, string filterName, string categoryName)
        {
            try
            {
                // ✅ DYNAMIC PATH: Use same pattern as other services
                var filtersDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "JSE_MEP_Openings", 
                    "Projects", 
                    "Default", 
                    "Filters");
                
                // Create directory if it doesn't exist
                if (!Directory.Exists(filtersDirectory))
                {
                    Directory.CreateDirectory(filtersDirectory);
                }
                
                // ✅ DYNAMIC FILENAME: {FilterName}_{CategoryName}_CLUSTER.xml
                var fileName = $"{filterName}_{categoryName}_CLUSTER.xml";
                var filePath = Path.Combine(filtersDirectory, fileName);
                
                // Create XML document
                var xmlDoc = new XmlDocument();
                var rootElement = xmlDoc.CreateElement("SleeveDataList");
                xmlDoc.AppendChild(rootElement);
                
                foreach (var sleeveData in sleeveDataList)
                {
                    var sleeveElement = xmlDoc.CreateElement("SleeveData");
                    
                    // Add all properties
                    AddXmlElement(xmlDoc, sleeveElement, "SleeveInstanceId", sleeveData.SleeveInstanceId.ToString());
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1X", sleeveData.Corner1.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1Y", sleeveData.Corner1.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner1Z", sleeveData.Corner1.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2X", sleeveData.Corner2.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2Y", sleeveData.Corner2.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner2Z", sleeveData.Corner2.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3X", sleeveData.Corner3.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3Y", sleeveData.Corner3.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner3Z", sleeveData.Corner3.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4X", sleeveData.Corner4.X.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4Y", sleeveData.Corner4.Y.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Corner4Z", sleeveData.Corner4.Z.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Width", sleeveData.Width.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Height", sleeveData.Height.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "Depth", sleeveData.Depth.ToString("F6"));
                    AddXmlElement(xmlDoc, sleeveElement, "HostType", sleeveData.HostType);
                    AddXmlElement(xmlDoc, sleeveElement, "Orientation", sleeveData.Orientation);
                    AddXmlElement(xmlDoc, sleeveElement, "Category", sleeveData.Category);
                    AddXmlElement(xmlDoc, sleeveElement, "CreatedAt", sleeveData.CreatedAt);
                    
                    rootElement.AppendChild(sleeveElement);
                }
                
                // Save XML file
                xmlDoc.Save(filePath);
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\sleeve_data.log",
                    $"[SAVE-CLUSTER-XML] Saved {sleeveDataList.Count} sleeves to {fileName}\n");
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\sleeve_data.log",
                    $"[ERROR] Failed to save cluster XML: {ex.Message}\n");
            }
        }
        
        /// <summary>
        /// Load sleeve data from _CLUSTER.xml file
        /// </summary>
        public List<SleeveData> LoadSleeveDataFromClusterXml(string filterName, string categoryName)
        {
            var sleeveDataList = new List<SleeveData>();
            
            try
            {
                // ✅ DYNAMIC PATH: Use same pattern as other services
                var filtersDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "JSE_MEP_Openings", 
                    "Projects", 
                    "Default", 
                    "Filters");
                
                // ✅ DYNAMIC FILENAME: {FilterName}_{CategoryName}_CLUSTER.xml
                var fileName = $"{filterName}_{categoryName}_CLUSTER.xml";
                var filePath = Path.Combine(filtersDirectory, fileName);
                
                if (!File.Exists(filePath))
                {
                    System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\sleeve_data.log",
                        $"[LOAD-CLUSTER-XML] File not found: {fileName}\n");
                    return sleeveDataList;
                }
                
                var xmlDoc = new XmlDocument();
                xmlDoc.Load(filePath);
                
                var sleeveNodes = xmlDoc.SelectNodes("//SleeveData");
                if (sleeveNodes != null)
                {
                    foreach (XmlNode node in sleeveNodes)
                    {
                        var sleeveData = new SleeveData();
                        
                        // Load all properties
                        if (int.TryParse(node.SelectSingleNode("SleeveInstanceId")?.InnerText, out int sleeveId))
                            sleeveData.SleeveInstanceId = sleeveId;
                        
                        // Load corners
                        if (double.TryParse(node.SelectSingleNode("Corner1X")?.InnerText, out double c1x) &&
                            double.TryParse(node.SelectSingleNode("Corner1Y")?.InnerText, out double c1y) &&
                            double.TryParse(node.SelectSingleNode("Corner1Z")?.InnerText, out double c1z))
                            sleeveData.Corner1 = new XYZ(c1x, c1y, c1z);
                        
                        if (double.TryParse(node.SelectSingleNode("Corner2X")?.InnerText, out double c2x) &&
                            double.TryParse(node.SelectSingleNode("Corner2Y")?.InnerText, out double c2y) &&
                            double.TryParse(node.SelectSingleNode("Corner2Z")?.InnerText, out double c2z))
                            sleeveData.Corner2 = new XYZ(c2x, c2y, c2z);
                        
                        if (double.TryParse(node.SelectSingleNode("Corner3X")?.InnerText, out double c3x) &&
                            double.TryParse(node.SelectSingleNode("Corner3Y")?.InnerText, out double c3y) &&
                            double.TryParse(node.SelectSingleNode("Corner3Z")?.InnerText, out double c3z))
                            sleeveData.Corner3 = new XYZ(c3x, c3y, c3z);
                        
                        if (double.TryParse(node.SelectSingleNode("Corner4X")?.InnerText, out double c4x) &&
                            double.TryParse(node.SelectSingleNode("Corner4Y")?.InnerText, out double c4y) &&
                            double.TryParse(node.SelectSingleNode("Corner4Z")?.InnerText, out double c4z))
                            sleeveData.Corner4 = new XYZ(c4x, c4y, c4z);
                        
                        // Load dimensions
                        if (double.TryParse(node.SelectSingleNode("Width")?.InnerText, out double width))
                            sleeveData.Width = width;
                        if (double.TryParse(node.SelectSingleNode("Height")?.InnerText, out double height))
                            sleeveData.Height = height;
                        if (double.TryParse(node.SelectSingleNode("Depth")?.InnerText, out double depth))
                            sleeveData.Depth = depth;
                        
                        // Load metadata
                        sleeveData.HostType = node.SelectSingleNode("HostType")?.InnerText ?? "";
                        sleeveData.Orientation = node.SelectSingleNode("Orientation")?.InnerText ?? "";
                        sleeveData.Category = node.SelectSingleNode("Category")?.InnerText ?? "";
                        
                        sleeveDataList.Add(sleeveData);
                    }
                }
                
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\sleeve_data.log",
                    $"[LOAD-CLUSTER-XML] Loaded {sleeveDataList.Count} sleeves from {fileName}\n");
            }
            catch (Exception ex)
            {
                System.IO.File.AppendAllText(@"C:\JSE_CSharp_Projects\JSE_MEPOPENING_23\Log\sleeve_data.log",
                    $"[ERROR] Failed to load cluster XML: {ex.Message}\n");
            }
            
            return sleeveDataList;
        }
        
        private void AddXmlElement(XmlDocument xmlDoc, XmlElement parentElement, string elementName, string value)
        {
            var element = xmlDoc.CreateElement(elementName);
            element.InnerText = value;
            parentElement.AppendChild(element);
        }
    }
}
