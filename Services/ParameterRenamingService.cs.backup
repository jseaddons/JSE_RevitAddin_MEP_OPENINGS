using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for managing parameter value renaming conditions
    /// </summary>
    public class ParameterRenamingService
    {
        private List<RenamingCondition> _renamingConditions;
        
        public ParameterRenamingService()
        {
            _renamingConditions = new List<RenamingCondition>();
        }
        
        /// <summary>
        /// Apply renaming conditions to a parameter value
        /// </summary>
        public string ApplyRenaming(string originalValue, string parameterName)
        {
            if (string.IsNullOrEmpty(originalValue)) return originalValue;
            
            try
            {
                // Find applicable renaming conditions
                var applicableConditions = _renamingConditions
                    .Where(c => c.IsEnabled && 
                               (string.IsNullOrEmpty(c.ParameterName) || 
                                c.ParameterName.Equals(parameterName, StringComparison.OrdinalIgnoreCase)) &&
                               c.OriginalValue.Equals(originalValue, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                
                if (applicableConditions.Count == 0)
                    return originalValue;
                
                // Use the first applicable condition
                return applicableConditions.First().NewValue;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying renaming: {ex.Message}");
                return originalValue;
            }
        }
        
        /// <summary>
        /// Add a renaming condition
        /// </summary>
        public void AddRenamingCondition(RenamingCondition condition)
        {
            if (condition == null) return;
            
            try
            {
                // Check if condition already exists
                var existingCondition = _renamingConditions.FirstOrDefault(c =>
                    c.OriginalValue.Equals(condition.OriginalValue, StringComparison.OrdinalIgnoreCase) &&
                    c.ParameterName.Equals(condition.ParameterName, StringComparison.OrdinalIgnoreCase));
                
                if (existingCondition != null)
                {
                    // Update existing condition
                    existingCondition.NewValue = condition.NewValue;
                    existingCondition.IsEnabled = condition.IsEnabled;
                }
                else
                {
                    // Add new condition
                    _renamingConditions.Add(condition);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adding renaming condition: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Remove a renaming condition
        /// </summary>
        public bool RemoveRenamingCondition(string originalValue, string parameterName)
        {
            try
            {
                var condition = _renamingConditions.FirstOrDefault(c =>
                    c.OriginalValue.Equals(originalValue, StringComparison.OrdinalIgnoreCase) &&
                    c.ParameterName.Equals(parameterName, StringComparison.OrdinalIgnoreCase));
                
                if (condition != null)
                {
                    _renamingConditions.Remove(condition);
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error removing renaming condition: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Get all renaming conditions
        /// </summary>
        public List<RenamingCondition> GetAllRenamingConditions()
        {
            return _renamingConditions.ToList();
        }
        
        /// <summary>
        /// Get renaming conditions for a specific parameter
        /// </summary>
        public List<RenamingCondition> GetRenamingConditionsForParameter(string parameterName)
        {
            return _renamingConditions
                .Where(c => string.IsNullOrEmpty(c.ParameterName) || 
                           c.ParameterName.Equals(parameterName, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }
        
        /// <summary>
        /// Clear all renaming conditions
        /// </summary>
        public void ClearAllRenamingConditions()
        {
            _renamingConditions.Clear();
        }
        
        /// <summary>
        /// Enable or disable a renaming condition
        /// </summary>
        public bool SetRenamingConditionEnabled(string originalValue, string parameterName, bool enabled)
        {
            try
            {
                var condition = _renamingConditions.FirstOrDefault(c =>
                    c.OriginalValue.Equals(originalValue, StringComparison.OrdinalIgnoreCase) &&
                    c.ParameterName.Equals(parameterName, StringComparison.OrdinalIgnoreCase));
                
                if (condition != null)
                {
                    condition.IsEnabled = enabled;
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting renaming condition enabled: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// Import renaming conditions from CSV file
        /// </summary>
        public List<RenamingCondition> ImportFromCsv(string csvPath)
        {
            var importedConditions = new List<RenamingCondition>();
            
            try
            {
                if (!File.Exists(csvPath))
                {
                    System.Diagnostics.Debug.WriteLine($"CSV file not found: {csvPath}");
                    return importedConditions;
                }
                
                var lines = File.ReadAllLines(csvPath);
                if (lines.Length == 0) return importedConditions;
                
                // Skip header line if it exists
                var startIndex = 0;
                if (lines[0].Contains("Original Value") || lines[0].Contains("New Value"))
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
                        var condition = new RenamingCondition
                        {
                            OriginalValue = parts[0],
                            NewValue = parts[1],
                            ParameterName = parts.Count > 2 ? parts[2] : string.Empty,
                            IsEnabled = true
                        };
                        
                        importedConditions.Add(condition);
                        AddRenamingCondition(condition);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error importing CSV: {ex.Message}");
            }
            
            return importedConditions;
        }
        
        /// <summary>
        /// Export renaming conditions to CSV file
        /// </summary>
        public bool ExportToCsv(List<RenamingCondition> conditions, string csvPath)
        {
            try
            {
                if (conditions == null || conditions.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("No conditions to export");
                    return false;
                }
                
                var lines = new List<string>();
                
                // Add header
                lines.Add("Original Value,New Value,Parameter Name");
                
                // Add conditions
                foreach (var condition in conditions)
                {
                    var line = $"{EscapeCsvValue(condition.OriginalValue)}," +
                              $"{EscapeCsvValue(condition.NewValue)}," +
                              $"{EscapeCsvValue(condition.ParameterName)}";
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
        /// Get predefined renaming conditions for common scenarios
        /// </summary>
        public List<RenamingCondition> GetPredefinedRenamingConditions()
        {
            var conditions = new List<RenamingCondition>();
            
            // Common system abbreviation mappings
            conditions.Add(new RenamingCondition("Sprinklers", "SPR", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Hydronic", "S", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Sanitary", "S", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Supply Air", "V", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Exhaust Air", "V", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Return Air", "R", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Electrical", "E", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Lighting", "L", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Power", "P", "System Abbreviation"));
            conditions.Add(new RenamingCondition("Fire Protection", "FP", "System Abbreviation"));
            
            // Common material mappings
            conditions.Add(new RenamingCondition("Concrete", "C", "Material"));
            conditions.Add(new RenamingCondition("Steel", "S", "Material"));
            conditions.Add(new RenamingCondition("Aluminum", "AL", "Material"));
            conditions.Add(new RenamingCondition("Wood", "W", "Material"));
            conditions.Add(new RenamingCondition("Glass", "GL", "Material"));
            
            // Common fire rating mappings
            conditions.Add(new RenamingCondition("1 Hour", "1HR", "Fire Rating"));
            conditions.Add(new RenamingCondition("2 Hour", "2HR", "Fire Rating"));
            conditions.Add(new RenamingCondition("3 Hour", "3HR", "Fire Rating"));
            conditions.Add(new RenamingCondition("4 Hour", "4HR", "Fire Rating"));
            
            return conditions;
        }
        
        /// <summary>
        /// Load predefined renaming conditions
        /// </summary>
        public void LoadPredefinedRenamingConditions()
        {
            try
            {
                var predefinedConditions = GetPredefinedRenamingConditions();
                foreach (var condition in predefinedConditions)
                {
                    AddRenamingCondition(condition);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading predefined conditions: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Validate renaming conditions
        /// </summary>
        public List<string> ValidateRenamingConditions(List<RenamingCondition> conditions)
        {
            var errors = new List<string>();
            
            try
            {
                foreach (var condition in conditions)
                {
                    if (string.IsNullOrEmpty(condition.OriginalValue))
                    {
                        errors.Add("Original value cannot be empty");
                    }
                    
                    if (string.IsNullOrEmpty(condition.NewValue))
                    {
                        errors.Add("New value cannot be empty");
                    }
                    
                    if (condition.OriginalValue.Equals(condition.NewValue, StringComparison.OrdinalIgnoreCase))
                    {
                        errors.Add($"Original and new values cannot be the same: {condition.OriginalValue}");
                    }
                }
                
                // Check for duplicates
                var duplicates = conditions
                    .GroupBy(c => new { c.OriginalValue, c.ParameterName })
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key);
                
                foreach (var duplicate in duplicates)
                {
                    errors.Add($"Duplicate condition found: {duplicate.OriginalValue} for parameter {duplicate.ParameterName}");
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Error validating conditions: {ex.Message}");
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
}
