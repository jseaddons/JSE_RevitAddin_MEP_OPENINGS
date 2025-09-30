using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class MarkParameterAddValue : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Initialize a custom log file for this execution
            DebugLogger.InitCustomLogFile("MarkParameterAddValue_Debug");

            // Ask user for prefix
            string prefix = UtilityClass.GetPrefixFromUser();
            if (string.IsNullOrEmpty(prefix))
            {
                return Result.Cancelled;
            }

            return ExecuteWithPrefix(doc, uiDoc, prefix);
        }

        /// <summary>
        /// ExecuteImpl method for orchestrator integration
        /// </summary>
        public CommandExecutionResult ExecuteImpl(Document doc, UIDocument uiDoc, string prefix)
        {
            var result = new CommandExecutionResult();
            
            try
            {
                DebugLogger.Info($"[MARK_PARAMETER] Using prefix from UI: {prefix}");
                
                var commandResult = ExecuteWithPrefix(doc, uiDoc, prefix);
                result.Success = commandResult == Result.Succeeded;
                result.ErrorMessage = result.Success ? null : $"MarkParameterAddValue failed with result: {commandResult}";
                
                return result;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[MARK_PARAMETER] ExecuteImpl failed: {ex.Message}");
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
        }

        /// <summary>
        /// Executes the mark parameter assignment with a provided prefix
        /// </summary>
        public Result ExecuteWithPrefix(Document doc, UIDocument uiDoc, string prefix)
        {
            // Initialize a custom log file for this execution
            DebugLogger.InitCustomLogFile("MarkParameterAddValue_Debug");

            if (string.IsNullOrEmpty(prefix))
            {
                return Result.Cancelled;
            }

            using (Transaction t = new Transaction(doc, "Add Mark Parameter Values"))
            {
                t.Start();

                // Assign marks to existing openings
                var allOpenings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_GenericModel)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .Where(fi => fi.Symbol.Family.Name.IndexOf("OpeningOnWall", StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                int pipeIndex = 1, ductIndex = 1, damperIndex = 1, cableTrayIndex = 1, clusterPipeIndex = 1, clusterDuctIndex = 1, clusterCableTrayIndex = 1;
                foreach (var fi in allOpenings)
                {
                    var markParam = fi.LookupParameter("Mark");
                    if (markParam == null || markParam.IsReadOnly)
                    {
                        DebugLogger.Warning($"Skipping element ID {fi.Id.IntegerValue}: Mark parameter is null or read-only.");
                        continue;
                    }

                    string existingMark = markParam.AsString();
                    DebugLogger.Debug($"Processing element ID {fi.Id.IntegerValue}: Existing Mark = '{existingMark}', Family Name = '{fi.Symbol.Family.Name}', Family Type Name = '{fi.Symbol.Name}'");

                    // Remove previously added prefix if present
                    if (!string.IsNullOrEmpty(existingMark) && existingMark.StartsWith(prefix))
                    {
                        existingMark = existingMark.Substring(prefix.Length);
                    }

                    // Clear the existing mark value entirely
                    existingMark = string.Empty;

                    string familyName = fi.Symbol.Family.Name;
                    string typeName = fi.Symbol.Name;

                    DebugLogger.Debug($"Matching logic: Family Name = '{familyName}', Type Name = '{typeName}'");

                    if (familyName.Contains("Pipe", StringComparison.OrdinalIgnoreCase))
                    {
                        markParam.Set($"{prefix}P{pipeIndex:000}");
                        DebugLogger.Info($"Assigned Mark: {prefix}P{pipeIndex:000}");
                        pipeIndex++;
                    }
                    else if (familyName.Contains("Duct", StringComparison.OrdinalIgnoreCase))
                    {
                        markParam.Set($"{prefix}D{ductIndex:000}");
                        DebugLogger.Info($"Assigned Mark: {prefix}D{ductIndex:000}");
                        ductIndex++;
                    }
                    else if (familyName.Contains("Damper", StringComparison.OrdinalIgnoreCase))
                    {
                        markParam.Set($"{prefix}D{damperIndex:000}");
                        DebugLogger.Info($"Assigned Mark: {prefix}D{damperIndex:000}");
                        damperIndex++;
                    }
                    else if (familyName.Contains("CableTray", StringComparison.OrdinalIgnoreCase))
                    {
                        markParam.Set($"{prefix}C{cableTrayIndex:000}");
                        DebugLogger.Info($"Assigned Mark: {prefix}C{cableTrayIndex:000}");
                        cableTrayIndex++;
                    }
                    // Cluster opening sleeve families: category-specific prefix
                    else if (familyName.StartsWith("Cluster", StringComparison.OrdinalIgnoreCase) ||
                             familyName.IndexOf("Cluster", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             typeName.EndsWith("Rect", StringComparison.OrdinalIgnoreCase))
                    {
                        // Determine cluster category based on context or family name
                        if (familyName.Contains("Pipe", StringComparison.OrdinalIgnoreCase) || 
                            typeName.Contains("Pipe", StringComparison.OrdinalIgnoreCase))
                        {
                            markParam.Set($"{prefix}CP{clusterPipeIndex:000}");
                            DebugLogger.Info($"Assigned Mark: {prefix}CP{clusterPipeIndex:000}");
                            clusterPipeIndex++;
                        }
                        else if (familyName.Contains("Duct", StringComparison.OrdinalIgnoreCase) || 
                                 familyName.Contains("Damper", StringComparison.OrdinalIgnoreCase) ||
                                 typeName.Contains("Duct", StringComparison.OrdinalIgnoreCase) ||
                                 typeName.Contains("Damper", StringComparison.OrdinalIgnoreCase))
                        {
                            markParam.Set($"{prefix}CD{clusterDuctIndex:000}");
                            DebugLogger.Info($"Assigned Mark: {prefix}CD{clusterDuctIndex:000}");
                            clusterDuctIndex++;
                        }
                        else if (familyName.Contains("CableTray", StringComparison.OrdinalIgnoreCase) || 
                                 familyName.Contains("Cable", StringComparison.OrdinalIgnoreCase) ||
                                 typeName.Contains("CableTray", StringComparison.OrdinalIgnoreCase) ||
                                 typeName.Contains("Cable", StringComparison.OrdinalIgnoreCase))
                        {
                            markParam.Set($"{prefix}CC{clusterCableTrayIndex:000}");
                            DebugLogger.Info($"Assigned Mark: {prefix}CC{clusterCableTrayIndex:000}");
                            clusterCableTrayIndex++;
                        }
                        else
                        {
                            // Default cluster marking if category cannot be determined
                            markParam.Set($"{prefix}CD{clusterDuctIndex:000}");
                            DebugLogger.Info($"Assigned Mark (default cluster): {prefix}CD{clusterDuctIndex:000}");
                            clusterDuctIndex++;
                        }
                    }
                    else
                    {
                        DebugLogger.Warning($"Element ID {fi.Id.IntegerValue} does not match any known family name patterns. Family Name: {familyName}, Type Name: {typeName}");
                    }
                }

                t.Commit();
            }

            TaskDialog.Show("Schedule", "Mark parameters updated successfully.");
            return Result.Succeeded;
        }
    }
}
