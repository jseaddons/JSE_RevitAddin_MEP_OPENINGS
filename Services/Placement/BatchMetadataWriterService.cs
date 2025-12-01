using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// Batch metadata writer service for deferred parameter writes.
    /// Writes non-critical metadata to sleeves after placement to optimize performance.
    /// Part of Phase 2 Medium-Risk optimization strategy.
    /// </summary>
    public class BatchMetadataWriterService
    {
        private readonly Document _doc;

        public BatchMetadataWriterService(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// Write deferred metadata parameters to all placed sleeves in a single batch operation.
        /// Non-critical parameters: MEP_UniqueId, MEP_Size, System_Abbreviation, MEP_Count, Bottom of Opening, Host Parameters.
        /// </summary>
        /// <param name="sleeveDataList">List of tuples containing sleeve instance and associated clash zone data</param>
        public void WriteDeferredMetadata(List<(FamilyInstance sleeve, ClashZone clashZone)> sleeveDataList)
        {
            if (sleeveDataList == null || sleeveDataList.Count == 0)
                return;

            var sw = Stopwatch.StartNew();
            int totalParams = 0;
            int failedParams = 0;

            try
            {
                foreach (var (sleeve, clashZone) in sleeveDataList)
                {
                    try
                    {
                        // Write non-critical MEP metadata
                        totalParams += WriteNonCriticalMepMetadata(sleeve, clashZone, out int failed);
                        failedParams += failed;

                        // Write Bottom of Opening (calculated parameter)
                        totalParams += WriteBottomOfOpening(sleeve, clashZone, out failed);
                        failedParams += failed;

                        // Write host parameters transferred from XML
                        totalParams += WriteHostParameters(sleeve, clashZone, out failed);
                        failedParams += failed;
                    }
                    catch (Exception ex)
                    {
                        if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Warning($"[BatchMetadataWriter] Error writing deferred metadata for sleeve {sleeve.Id}: {ex.Message}");
                    }
                }

                sw.Stop();

                if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                {
                    DebugLogger.Info($"[BatchMetadataWriter] Wrote {totalParams} deferred parameters to {sleeveDataList.Count} sleeves in {sw.ElapsedMilliseconds}ms ({failedParams} failed)");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Error($"[BatchMetadataWriter] Fatal error in batch write: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Write non-critical MEP metadata parameters.
        /// </summary>
        private int WriteNonCriticalMepMetadata(FamilyInstance sleeve, ClashZone clashZone, out int failed)
        {
            failed = 0;
            int written = 0;

            try
            {
                // MEP_UniqueId
                var param = sleeve.LookupParameter("MEP_UniqueId");
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(clashZone.MepElementUniqueId);
                    written++;
                }
                else failed++;

                // MEP_Size - ✅ CRITICAL FIX: Ensure we set as STRING value (text) for all categories
                // ✅ SIZE PARAMETER VALUE: Use MepElementSizeParameterValue (raw Size parameter value) instead of MepElementFormattedSize
                // MepElementSizeParameterValue contains the exact text from the Size parameter (e.g., "20 mmø", "200 mm dia symbol")
                // Fallback to MepElementFormattedSize if MepElementSizeParameterValue is empty
                param = sleeve.LookupParameter("MEP_Size");
                if (param != null && !param.IsReadOnly)
                {
                    var sizeValue = !string.IsNullOrWhiteSpace(clashZone.MepElementSizeParameterValue) 
                        ? clashZone.MepElementSizeParameterValue 
                        : (clashZone.MepElementFormattedSize ?? string.Empty);
                    
                    // ✅ DIAGNOSTIC: Log what we're trying to set
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [BatchMetadataWriter] 🔍 Attempting to set MEP_Size on sleeve {sleeve.Id}: StorageType={param.StorageType}, MepElementSizeParameterValue='{clashZone.MepElementSizeParameterValue}', MepElementFormattedSize='{clashZone.MepElementFormattedSize}', sizeValue='{sizeValue}', Category={clashZone.MepElementCategory}\n");
                    }
                    
                    // ✅ FIX: Check StorageType - if it's String, set directly; if Double, we should NOT set it (it's for calculation only)
                    // The MEP_Size parameter on sleeve should be a String parameter to display text values
                    if (param.StorageType == StorageType.String)
                    {
                        if (!string.IsNullOrWhiteSpace(sizeValue))
                        {
                            try
                            {
                                param.Set(sizeValue);
                                written++;
                                
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [BatchMetadataWriter] ✅ SUCCESS: Set MEP_Size = '{sizeValue}' (String) for sleeve {sleeve.Id}\n");
                                    DebugLogger.Info($"[BatchMetadataWriter] ✅ Set MEP_Size = '{sizeValue}' (String) for sleeve {sleeve.Id}");
                                }
                            }
                            catch (Exception setEx)
                            {
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [BatchMetadataWriter] ❌ ERROR: Failed to set MEP_Size='{sizeValue}' on sleeve {sleeve.Id}: {setEx.Message}\n");
                                    DebugLogger.Warning($"[BatchMetadataWriter] Failed to set MEP_Size='{sizeValue}' on sleeve {sleeve.Id}: {setEx.Message}");
                                }
                                failed++;
                            }
                        }
                        else
                        {
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [BatchMetadataWriter] ⚠️ WARNING: MEP_Size is empty/null for sleeve {sleeve.Id} (MepElementSizeParameterValue='{clashZone.MepElementSizeParameterValue}', MepElementFormattedSize='{clashZone.MepElementFormattedSize}')\n");
                                DebugLogger.Warning($"[BatchMetadataWriter] MEP_Size is empty/null for sleeve {sleeve.Id} (MepElementSizeParameterValue='{clashZone.MepElementSizeParameterValue}', MepElementFormattedSize='{clashZone.MepElementFormattedSize}')");
                            }
                            failed++;
                        }
                    }
                    else
                    {
                        // ⚠️ WARNING: MEP_Size parameter is not a String type - cannot set text value
                        // This should be a String parameter to display formatted sizes like "20 mmø"
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [BatchMetadataWriter] ⚠️ CRITICAL: MEP_Size parameter on sleeve {sleeve.Id} is {param.StorageType} (expected String). Cannot set text value '{sizeValue}'. Parameter should be String type to display formatted sizes.\n");
                            DebugLogger.Warning($"[BatchMetadataWriter] ⚠️ MEP_Size parameter on sleeve {sleeve.Id} is {param.StorageType} (expected String). Cannot set text value '{sizeValue}'. Parameter should be String type to display formatted sizes.");
                        }
                        failed++;
                    }
                }
                else 
                {
                    if (param == null && !DeploymentConfiguration.DeploymentMode)
                    {
                        SafeFileLogger.SafeAppendText("parameter_service_debug.log",
                            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [BatchMetadataWriter] ⚠️ MEP_Size parameter not found on sleeve {sleeve.Id}\n");
                        DebugLogger.Warning($"[BatchMetadataWriter] MEP_Size parameter not found on sleeve {sleeve.Id}");
                    }
                    failed++;
                }

                // System_Abbreviation
                param = sleeve.LookupParameter("System_Abbreviation");
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(clashZone.MepElementSystemAbbreviation);
                    written++;
                }
                else failed++;

                // MEP_Count (always 1 for individual sleeves)
                param = sleeve.LookupParameter("MEP_Count");
                if (param != null && !param.IsReadOnly)
                {
                    param.Set(1);
                    written++;
                }
                else failed++;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[BatchMetadataWriter] Error writing MEP metadata: {ex.Message}");
                failed++;
            }

            return written;
        }

        /// <summary>
        /// Calculate and write Bottom of Opening parameter.
        /// Formula: Bottom of Opening = Elevation from Level - (Height / 2)
        /// </summary>
        private int WriteBottomOfOpening(FamilyInstance sleeve, ClashZone clashZone, out int failed)
        {
            failed = 0;
            int written = 0;

            try
            {
                // Only calculate for rectangular openings on walls and framing
                bool isWallHost = clashZone.StructuralElementType == "Wall" || clashZone.StructuralElementType == "Walls";
                bool isFramingHost = string.Equals(clashZone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);

                if (!isWallHost && !isFramingHost)
                    return 0;

                var elevationFromLevelParam = sleeve.LookupParameter("Elevation from Level")
                                           ?? sleeve.LookupParameter("Schedule Level Elevation")
                                           ?? sleeve.LookupParameter("Elevation from Level Offset");

                var heightParam = sleeve.LookupParameter("Height");
                var bottomOfOpeningParam = sleeve.LookupParameter("Bottom of Opening");

                if (elevationFromLevelParam != null && heightParam != null && bottomOfOpeningParam != null && !bottomOfOpeningParam.IsReadOnly)
                {
                    double elevationFromLevel = elevationFromLevelParam.AsDouble();
                    double height = heightParam.AsDouble();

                    if (height > 0 && Math.Abs(elevationFromLevel) < 10000) // Reasonable bounds check
                    {
                        double bottomOfOpening = elevationFromLevel - (height / 2.0);
                        bottomOfOpeningParam.Set(bottomOfOpening);
                        written++;
                    }
                    else
                    {
                        failed++;
                    }
                }
                else
                {
                    failed++;
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[BatchMetadataWriter] Error calculating Bottom of Opening: {ex.Message}");
                failed++;
            }

            return written;
        }

        /// <summary>
        /// Transfer host parameters from XML intersection data to sleeve parameters.
        /// </summary>
        private int WriteHostParameters(FamilyInstance sleeve, ClashZone clashZone, out int failed)
        {
            failed = 0;
            int written = 0;

            if (clashZone.HostParameterValues == null || clashZone.HostParameterValues.Count == 0)
                return 0;

            try
            {
                foreach (var hostParam in clashZone.HostParameterValues)
                {
                    try
                    {
                        var param = sleeve.LookupParameter(hostParam.Key);
                        if (param != null && !param.IsReadOnly)
                        {
                            if (param.StorageType == StorageType.String)
                            {
                                param.Set(hostParam.Value);
                                written++;
                            }
                            else if (param.StorageType == StorageType.Integer)
                            {
                                if (int.TryParse(hostParam.Value, out int intValue))
                                {
                                    param.Set(intValue);
                                    written++;
                                }
                                else
                                {
                                    failed++;
                                }
                            }
                            else if (param.StorageType == StorageType.Double)
                            {
                                if (double.TryParse(hostParam.Value, System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture, out double doubleValue))
                                {
                                    param.Set(doubleValue);
                                    written++;
                                }
                                else
                                {
                                    failed++;
                                }
                            }
                        }
                        else
                        {
                            failed++;
                        }
                    }
                    catch (Exception paramEx)
                    {
                        if (!DeploymentConfiguration.DeploymentMode && OptimizationFlags.UseDiagnosticMode)
                        {
                            DebugLogger.Warning($"[BatchMetadataWriter] Error setting host parameter '{hostParam.Key}': {paramEx.Message}");
                        }
                        failed++;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[BatchMetadataWriter] Error writing host parameters: {ex.Message}");
                failed++;
            }

            return written;
        }
    }
}
