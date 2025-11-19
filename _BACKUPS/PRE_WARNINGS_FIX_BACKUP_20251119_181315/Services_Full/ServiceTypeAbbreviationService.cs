using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for managing service type abbreviations based on user input
    /// </summary>
    public class ServiceTypeAbbreviationService
    {
        private Dictionary<string, string> _serviceAbbreviations;
        private Dictionary<string, string> _parameterMappings;
        
        public ServiceTypeAbbreviationService()
        {
            _serviceAbbreviations = new Dictionary<string, string>();
            _parameterMappings = new Dictionary<string, string>();
            LoadDefaultAbbreviations();
        }
        
        /// <summary>
        /// Get abbreviation for a service type
        /// </summary>
        public string GetAbbreviation(string serviceType, string parameterName = "")
        {
            if (string.IsNullOrEmpty(serviceType)) return serviceType;
            
            // First try exact match
            if (_serviceAbbreviations.ContainsKey(serviceType))
            {
                return _serviceAbbreviations[serviceType];
            }
            
            // Try case-insensitive match
            var caseInsensitiveMatch = _serviceAbbreviations.FirstOrDefault(kvp => 
                string.Equals(kvp.Key, serviceType, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(caseInsensitiveMatch.Key))
            {
                return caseInsensitiveMatch.Value;
            }
            
            // Try partial match (contains)
            var partialMatch = _serviceAbbreviations.FirstOrDefault(kvp => 
                serviceType.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(partialMatch.Key))
            {
                return partialMatch.Value;
            }
            
            // If no match found, return original value
            return serviceType;
        }
        
        /// <summary>
        /// Add or update a service type abbreviation
        /// </summary>
        public void AddAbbreviation(string serviceType, string abbreviation)
        {
            if (string.IsNullOrEmpty(serviceType) || string.IsNullOrEmpty(abbreviation))
                return;
                
            _serviceAbbreviations[serviceType] = abbreviation;
        }
        
        /// <summary>
        /// Remove a service type abbreviation
        /// </summary>
        public bool RemoveAbbreviation(string serviceType)
        {
            return _serviceAbbreviations.Remove(serviceType);
        }
        
        /// <summary>
        /// Get all service type abbreviations
        /// </summary>
        public Dictionary<string, string> GetAllAbbreviations()
        {
            return new Dictionary<string, string>(_serviceAbbreviations);
        }
        
        /// <summary>
        /// Clear all abbreviations
        /// </summary>
        public void ClearAllAbbreviations()
        {
            _serviceAbbreviations.Clear();
        }
        
        /// <summary>
        /// Import abbreviations from CSV file
        /// </summary>
        public List<ServiceTypeAbbreviation> ImportFromCsv(string csvPath)
        {
            var importedAbbreviations = new List<ServiceTypeAbbreviation>();
            
            try
            {
                if (!File.Exists(csvPath))
                {
                    System.Diagnostics.Debug.WriteLine($"CSV file not found: {csvPath}");
                    return importedAbbreviations;
                }
                
                var lines = File.ReadAllLines(csvPath);
                if (lines.Length == 0) return importedAbbreviations;
                
                // Skip header line if it exists
                var startIndex = 0;
                if (lines[0].Contains("Service Type") || lines[0].Contains("Abbreviation"))
                {
                    startIndex = 1;
                }
                
                for (int i = startIndex; i < lines.Length; i++)
                {
                    var line = lines[i].Trim();
                    if (string.IsNullOrEmpty(line)) continue;
                    
                    var parts = ParseCsvLine(line);
                    if (parts.Count >= 2)
                    {
                        var abbreviation = new ServiceTypeAbbreviation
                        {
                            ServiceType = parts[0],
                            Abbreviation = parts[1],
                            ParameterName = parts.Count > 2 ? parts[2] : string.Empty,
                            IsEnabled = true
                        };
                        
                        importedAbbreviations.Add(abbreviation);
                        AddAbbreviation(abbreviation.ServiceType, abbreviation.Abbreviation);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error importing CSV: {ex.Message}");
            }
            
            return importedAbbreviations;
        }
        
        /// <summary>
        /// Export abbreviations to CSV file
        /// </summary>
        public bool ExportToCsv(List<ServiceTypeAbbreviation> abbreviations, string csvPath)
        {
            try
            {
                if (abbreviations == null || abbreviations.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("No abbreviations to export");
                    return false;
                }
                
                var lines = new List<string>();
                
                // Add header
                lines.Add("Service Type,Abbreviation,Parameter Name");
                
                // Add abbreviations
                foreach (var abbreviation in abbreviations)
                {
                    var line = $"{EscapeCsvValue(abbreviation.ServiceType)}," +
                              $"{EscapeCsvValue(abbreviation.Abbreviation)}," +
                              $"{EscapeCsvValue(abbreviation.ParameterName)}";
                    lines.Add(line);
                }
                
                File.WriteAllLines(csvPath, lines);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error exporting CSV: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Get service type from MEP element
        /// </summary>
        public string GetServiceTypeFromElement(Element mepElement)
        {
            try
            {
                // Try different parameters to get service type
                var parametersToTry = new List<string>
                {
                    "System Abbreviation",
                    "System Name", 
                    "System Type",
                    "Family Name",
                    "Type Name",
                    "Comments"
                };
                
                foreach (var paramName in parametersToTry)
                {
                    var param = mepElement.LookupParameter(paramName);
                    if (param != null)
                    {
                        var value = param.AsString();
                        if (!string.IsNullOrEmpty(value))
                        {
                            return value;
                        }
                    }
                }
                
                // Fallback to category name
                return mepElement.Category?.Name ?? "Unknown";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting service type from element: {ex.Message}");
                return "Unknown";
            }
        }
        
        /// <summary>
        /// Get abbreviated service type from MEP element
        /// </summary>
        public string GetAbbreviatedServiceTypeFromElement(Element mepElement, string parameterName = "")
        {
            var serviceType = GetServiceTypeFromElement(mepElement);
            return GetAbbreviation(serviceType, parameterName);
        }
        
        /// <summary>
        /// Combine multiple service types with abbreviations
        /// </summary>
        public string CombineServiceTypes(List<Element> mepElements, string separator = " & ")
        {
            var abbreviatedTypes = new List<string>();
            
            foreach (var element in mepElements)
            {
                var abbreviatedType = GetAbbreviatedServiceTypeFromElement(element);
                if (!string.IsNullOrEmpty(abbreviatedType) && !abbreviatedTypes.Contains(abbreviatedType))
                {
                    abbreviatedTypes.Add(abbreviatedType);
                }
            }
            
            return string.Join(separator, abbreviatedTypes);
        }
        
        /// <summary>
        /// Load default abbreviations
        /// </summary>
        private void LoadDefaultAbbreviations()
        {
            // Common electrical abbreviations
            AddAbbreviation("Electrical Distribution Board", "EDB");
            AddAbbreviation("Distribution Board", "DB");
            AddAbbreviation("Main Distribution Board", "MDB");
            AddAbbreviation("Sub Distribution Board", "SDB");
            AddAbbreviation("Extra-Low Voltage", "ELV");
            AddAbbreviation("Low Voltage", "LV");
            AddAbbreviation("High Voltage", "HV");
            AddAbbreviation("Power", "P");
            AddAbbreviation("Lighting", "L");
            AddAbbreviation("Electrical", "E");
            
            // Common mechanical abbreviations
            AddAbbreviation("Supply Air", "SA");
            AddAbbreviation("Return Air", "RA");
            AddAbbreviation("Exhaust Air", "EA");
            AddAbbreviation("Fresh Air", "FA");
            AddAbbreviation("Ventilation", "V");
            AddAbbreviation("Air Conditioning", "AC");
            AddAbbreviation("Heating", "H");
            AddAbbreviation("Cooling", "C");
            AddAbbreviation("Mechanical", "M");
            
            // Common plumbing abbreviations
            AddAbbreviation("Cold Water", "CW");
            AddAbbreviation("Hot Water", "HW");
            AddAbbreviation("Domestic Cold Water", "DCW");
            AddAbbreviation("Domestic Hot Water", "DHW");
            AddAbbreviation("Sanitary", "S");
            AddAbbreviation("Hydronic", "H");
            AddAbbreviation("Sprinklers", "SPR");
            AddAbbreviation("Fire Protection", "FP");
            AddAbbreviation("Plumbing", "P");
            
            // Common communication abbreviations
            AddAbbreviation("Telecommunications", "TELECOM");
            AddAbbreviation("Data", "D");
            AddAbbreviation("Voice", "V");
            AddAbbreviation("Network", "N");
            AddAbbreviation("Cable Tray", "CT");
            AddAbbreviation("Conduit", "C");
            AddAbbreviation("Global System for Mobile", "GSM");
            AddAbbreviation("WiFi", "WIFI");
            AddAbbreviation("Fiber Optic", "FO");
            
            // Common material abbreviations
            AddAbbreviation("Concrete", "C");
            AddAbbreviation("Steel", "S");
            AddAbbreviation("Aluminum", "AL");
            AddAbbreviation("Wood", "W");
            AddAbbreviation("Glass", "GL");
            AddAbbreviation("Plastic", "PL");
            AddAbbreviation("Copper", "CU");
            AddAbbreviation("Iron", "FE");
        }
        
        /// <summary>
        /// Validate abbreviation
        /// </summary>
        public List<string> ValidateAbbreviation(ServiceTypeAbbreviation abbreviation)
        {
            var errors = new List<string>();
            
            if (string.IsNullOrEmpty(abbreviation.ServiceType))
            {
                errors.Add("Service type cannot be empty");
            }
            
            if (string.IsNullOrEmpty(abbreviation.Abbreviation))
            {
                errors.Add("Abbreviation cannot be empty");
            }
            
            if (abbreviation.ServiceType.Equals(abbreviation.Abbreviation, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add("Service type and abbreviation cannot be the same");
            }
            
            return errors;
        }
        
        #region Private Helper Methods
        
        private List<string> ParseCsvLine(string line)
        {
            var parts = new List<string>();
            var currentPart = string.Empty;
            var inQuotes = false;
            
            for (int i = 0; i < line.Length; i++)
            {
                var c = line[i];
                
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes)
                {
                    parts.Add(currentPart.Trim());
                    currentPart = string.Empty;
                }
                else
                {
                    currentPart += c;
                }
            }
            
            parts.Add(currentPart.Trim());
            return parts;
        }
        
        private string EscapeCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            
            // Escape quotes and wrap in quotes if contains comma or quote
            if (value.Contains(",") || value.Contains("\""))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            
            return value;
        }
        
        #endregion
    }
    
    /// <summary>
    /// Data model for service type abbreviation
    /// </summary>
    public class ServiceTypeAbbreviation
    {
        public string ServiceType { get; set; } = string.Empty;
        public string Abbreviation { get; set; } = string.Empty;
        public string ParameterName { get; set; } = string.Empty;
        public bool IsEnabled { get; set; } = true;
        public string Description { get; set; } = string.Empty;
        
        public ServiceTypeAbbreviation()
        {
        }
        
        public ServiceTypeAbbreviation(string serviceType, string abbreviation)
        {
            ServiceType = serviceType;
            Abbreviation = abbreviation;
            IsEnabled = true;
        }
        
        public ServiceTypeAbbreviation(string serviceType, string abbreviation, string parameterName)
        {
            ServiceType = serviceType;
            Abbreviation = abbreviation;
            ParameterName = parameterName;
            IsEnabled = true;
        }
    }
}
