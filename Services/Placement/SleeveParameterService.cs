using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Services.Placement;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ SRP COMPLIANCE: Service responsible for setting all parameters on sleeve elements.
    /// Single Responsibility: Parameter setting and deferred parameter management only.
    /// 
    /// Preserves all 28 features from COMPREHENSIVE_ARCHITECTURE_PLAN.md:
    /// - ✅ POINT 10: Parameter Batching (4-6× faster placement)
    /// - ✅ Performance Monitoring (tracks operation timings)
    /// - ✅ Safe Element Validation (avoids document mismatch bugs)
    /// - ✅ Global Settings (rounding based on configuration)
    /// - ✅ Diagnostic Logging (multi-level logging support)
    /// - ✅ Deployment Mode (reduces logging in production)
    /// - ✅ Transaction Safety (safe parameter writes)
    /// - ✅ Crash-Safe Execution (exception handling)
    /// - ✅ Flag-Based Control (respects optimization flags)
    /// - ✅ And all other features from comprehensive architecture
    /// </summary>
    public class SleeveParameterService
    {
        private readonly Document _doc;
        private readonly bool _isReplayPath;
        
        // ✅ PERFORMANCE MONITORING: Performance monitor for tracking operations
        private readonly PlacementPerformanceMonitor? _performanceMonitor;
        
        // ✅ PARAMETER BATCHING: Deferred parameter writes (4-6× faster placement)
        // Accumulates parameter values during placement loop, writes all after single regeneration
        // Key: ElementId of sleeve instance
        // Value: Dictionary of parameter name → value (double or string)
        private Dictionary<ElementId, Dictionary<string, object>> _deferredParameters = 
            new Dictionary<ElementId, Dictionary<string, object>>();
        
        // ✅ SAFETY FLAG: Prevents multiple flushes (critical for performance)
        private bool _hasFlushedParameters = false;

        public SleeveParameterService(
            Document doc,
            bool isReplayPath = false,
            PlacementPerformanceMonitor? performanceMonitor = null)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
            _isReplayPath = isReplayPath;
            _performanceMonitor = performanceMonitor;
        }

        /// <summary>
        /// ✅ MAIN METHOD: Set all parameters on a sleeve instance.
        /// Handles dimensions, metadata, clearances, and depth parameters.
        /// Preserves all optimization features: batching, monitoring, validation, logging.
        /// </summary>
        public void SetSleeveParameters(
            FamilyInstance instance, 
            double width, 
            double height, 
            double diameter, 
            bool isCircular, 
            ClashZone zone)
        {
            // ✅ PERFORMANCE MONITORING: Track parameter setting
            using (var tracker = _performanceMonitor?.TrackOperation("Set Sleeve Parameters"))
            {
                if (instance == null) return;

                // ✅ SAFE ELEMENT VALIDATION: Validate instance is still valid (avoids document mismatch bug)
                if (OptimizationFlags.UseSafeElementValidation)
                {
                    if (!ValidateElement(instance))
                        return;
                }

                var currentSleeveId = instance.Id;

                // ✅ GLOBAL SETTINGS: Apply rounding based on global configuration (RoundingValue and RoundAlwaysUp)
                var (roundedWidth, roundedHeight, roundedDiameter) = ApplyRounding(
                    width, height, diameter, zone);

                // Set dimensions (Width/Height or Diameter)
                if (isCircular)
                {
                    SetParameter(instance, "Diameter", roundedDiameter, currentSleeveId, 
                        fallbackName: "Sleeve Diameter");
                }
                else
                {
                    SetParameter(instance, "Width", roundedWidth, currentSleeveId, 
                        fallbackName: "Sleeve Width");
                    SetParameter(instance, "Height", roundedHeight, currentSleeveId, 
                        fallbackName: "Sleeve Height");
                    tracker?.SetItemCount(1);
                }

                // ✅ SRP COMPLIANCE: Delegate depth parameter setting to dedicated method
                if (zone != null)
                {
                    SetDepthParameter(instance, zone, currentSleeveId);
                }

                // ✅ FLAG MANAGEMENT SUPPORT: Set Sleeve Instance ID IMMEDIATELY (not deferred)
                // Flag management reads this parameter from Revit elements to identify individual sleeves
                SetSleeveInstanceId(instance, currentSleeveId);

                // ✅ CLUSTERING SUPPORT: Set MEP_ElementId and MEP_Category for clustering
                if (zone != null)
                {
                    SetMepMetadata(instance, zone, currentSleeveId);
                    SetDamperClearances(instance, zone, currentSleeveId);
                }
            }
        }

        /// <summary>
        /// ✅ SRP COMPLIANCE: Dedicated method for setting Depth/Wall Width parameter based on host type.
        /// Single Responsibility: Calculate and set structural thickness parameter only.
        /// Maintains all optimization features: batching, performance monitoring, safe validation, diagnostic logging.
        /// </summary>
        public void SetDepthParameter(FamilyInstance instance, ClashZone zone, ElementId currentSleeveId)
        {
            if (instance == null || zone == null) return;
            
            bool isWallHost = zone.StructuralElementType == "Wall" || zone.StructuralElementType == "Walls";
            bool isFramingHost = string.Equals(zone.StructuralElementType, "Structural Framing", StringComparison.OrdinalIgnoreCase);
            
            // Get the correct thickness based on host type
            double thickness = GetThickness(zone, isWallHost, isFramingHost);
            
            // ✅ CRITICAL FIX: For PATH 3 (Non-Fresh), if thickness is 0, retrieve from linked file
            if (!_isReplayPath && thickness <= 0.0 && zone.StructuralElementIdValue > 0)
            {
                thickness = RetrieveThicknessFromLinkedFile(zone, thickness);
            }
            
            // Set Depth or Wall Width parameter
            bool depthSetSuccess = false;
            if (isWallHost)
            {
                depthSetSuccess = SetParameter(instance, "Wall Width", thickness, currentSleeveId);
                if (depthSetSuccess && !DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[SleeveParameterService] [DEPTH-SET] Zone={zone.Id}, Sleeve={instance.Id}: Set Wall Width={thickness * 304.8:F1}mm");
                }
            }
            
            if (!depthSetSuccess)
            {
                depthSetSuccess = SetParameter(instance, "Depth", thickness, currentSleeveId);
                if (depthSetSuccess && !DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[SleeveParameterService] [DEPTH-SET] Zone={zone.Id}, Sleeve={instance.Id}: Set Depth={thickness * 304.8:F1}mm");
                }
            }
            
            if (!depthSetSuccess && !DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Warning($"[SleeveParameterService] [DEPTH-SET] ❌ Zone={zone.Id}, Sleeve={instance.Id}: Could not set Depth or Wall Width parameter");
            }
        }

        /// <summary>
        /// ✅ PARAMETER BATCHING: Flush all deferred parameters to Revit elements.
        /// Applies all accumulated parameter values after regeneration.
        /// Preserves all safety features: duplicate flush prevention, error handling, logging.
        /// </summary>
        public int FlushDeferredParameters()
        {
            // ✅ SAFETY FLAG: Prevent multiple flushes (critical for performance)
            if (_hasFlushedParameters)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    var stackTrace = new System.Diagnostics.StackTrace(skipFrames: 1, fNeedFileInfo: false);
                    var caller = stackTrace.GetFrame(0)?.GetMethod()?.Name ?? "Unknown";
                    DebugLogger.Warning($"[SleeveParameterService] [BATCH-PARAMS] ⚠️ SAFETY: FlushDeferredParameters called AGAIN from {caller} - IGNORING (already flushed once). This indicates a bug - parameters should only flush once at the end!");
                }
                return 0; // ✅ CRITICAL: Exit early to prevent duplicate flushes
            }
            
            if (_deferredParameters == null || _deferredParameters.Count == 0)
            {
                _hasFlushedParameters = true; // Mark as flushed even if empty
                return 0;
            }
            
            int successCount = 0;
            int failCount = 0;
            var errorLog = new System.Text.StringBuilder();
            
            if (!DeploymentConfiguration.DeploymentMode)
            {
                int totalParams = _deferredParameters.Values.Sum(d => d.Count);
                DebugLogger.Info($"[SleeveParameterService] [BATCH-PARAMS] 🔄 Flushing {_deferredParameters.Count} individual sleeves with {totalParams} total parameters...");
            }
            
            try
            {
                foreach (var kvp in _deferredParameters)
                {
                    var sleeveId = kvp.Key;
                    var paramValues = kvp.Value;
                    
                    var sleeve = _doc.GetElement(sleeveId) as FamilyInstance;
                    if (sleeve == null) continue;
                    
                    foreach (var paramKvp in paramValues)
                    {
                        var param = sleeve.LookupParameter(paramKvp.Key);
                        if (param != null && !param.IsReadOnly)
                        {
                            try
                            {
                                if (paramKvp.Value is double dVal)
                                    param.Set(dVal);
                                else if (paramKvp.Value is string sVal)
                                    param.Set(sVal);
                                else if (paramKvp.Value is int iVal)
                                    param.Set(iVal);
                                
                                successCount++;
                            }
                            catch (Exception ex)
                            {
                                failCount++;
                                errorLog.AppendLine($"[SleeveParameterService] Failed to set parameter '{paramKvp.Key}' on sleeve {sleeveId.IntegerValue}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[SleeveParameterService] [BATCH-PARAMS] Error during flush: {ex.Message}");
                }
            }
            finally
            {
                _hasFlushedParameters = true;
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    int totalParams = _deferredParameters.Values.Sum(d => d.Count);
                    DebugLogger.Info($"[SleeveParameterService] [BATCH-PARAMS] ✅ Flushed {successCount} parameters for {_deferredParameters.Count} sleeves, {failCount} failed");
                    if (errorLog.Length > 0)
                    {
                        SafeFileLogger.SafeAppendText("parameter_batching_errors.log", errorLog.ToString());
                    }
                }
                
                // Clear deferred parameters after flush (ready for next placement batch)
                _deferredParameters.Clear();
            }
            
            return successCount;
        }

        /// <summary>
        /// ✅ RESET: Reset the flush flag for a new placement batch.
        /// Called at the start of each placement run.
        /// </summary>
        public void ResetFlushFlag()
        {
            _hasFlushedParameters = false;
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Read parameter from deferred cache first, then fallback to Revit element.
        /// This prevents stale reads during corner placement calculations when batching is enabled.
        /// </summary>
        public double GetParameterValueWithBatchingSupport(
            FamilyInstance sleeve, 
            string parameterName, 
            double fallbackValue)
        {
            if (OptimizationFlags.UseBatchedParameterWrites && _deferredParameters != null)
            {
                var sleeveId = sleeve.Id;
                if (_deferredParameters.ContainsKey(sleeveId) && 
                    _deferredParameters[sleeveId].ContainsKey(parameterName))
                {
                    var cachedValue = _deferredParameters[sleeveId][parameterName];
                    if (cachedValue is double dVal)
                        return dVal;
                }
            }
            
            // Fallback to Revit element parameter
            var param = sleeve.LookupParameter(parameterName);
            if (param != null && param.StorageType == StorageType.Double)
            {
                return param.AsDouble();
            }
            
            return fallbackValue;
        }

        // ============================================================================
        // PRIVATE HELPER METHODS (SRP: Each method has single responsibility)
        // ============================================================================

        /// <summary>
        /// ✅ SAFE ELEMENT VALIDATION: Validate instance is still valid (avoids document mismatch bug)
        /// </summary>
        private bool ValidateElement(FamilyInstance instance)
        {
            try
            {
                // ⚠️ CRITICAL: Do NOT compare documents by reference (causes false positives like in parameter transfer)
                // Instead, validate by element ID - if doc.GetElement() succeeds, element is in correct document
                var validationElement = _doc.GetElement(instance.Id);
                if (validationElement == null || !validationElement.IsValidObject)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[SleeveParameterService] ⚠️ Instance {instance.Id.IntegerValue} is invalid - skipping parameter setting");
                    }
                    return false;
                }

                // Additional validation: ensure element IDs match (not just document reference)
                if (validationElement.Id != instance.Id)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        DebugLogger.Warning($"[SleeveParameterService] ⚠️ Element ID mismatch for instance {instance.Id.IntegerValue}");
                    }
                    return false;
                }
                
                return true;
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"[SleeveParameterService] Error validating instance {instance.Id.IntegerValue}: {ex.Message}");
                }
                return false;
            }
        }

        /// <summary>
        /// ✅ GLOBAL SETTINGS: Apply rounding based on global configuration
        /// </summary>
        private (double roundedWidth, double roundedHeight, double roundedDiameter) ApplyRounding(
            double width, double height, double diameter, ClashZone zone)
        {
            // Dampers are excluded from rounding (preserve exact calculated dimensions)
            bool isDamper = zone?.MepElementCategory != null && 
                           (zone.MepElementCategory.IndexOf("Damper", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            zone.MepElementCategory.IndexOf("Duct Accessories", StringComparison.OrdinalIgnoreCase) >= 0);
            
            double roundedWidth = width;
            double roundedHeight = height;
            double roundedDiameter = diameter;
            
            if (!isDamper)
            {
                // ✅ Apply rounding for non-damper elements using global settings
                // OpeningSettingsHelper reads RoundingValue and RoundAlwaysUp from ApplicationProfileService
                (roundedWidth, roundedHeight) = OpeningSettingsHelper.RoundDimensionsToNearest5mm(width, height);
                roundedDiameter = OpeningSettingsHelper.RoundDiameterToNearest5mm(diameter);
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    // Log rounding if values changed
                    if (Math.Abs(width - roundedWidth) > 1e-6 || Math.Abs(height - roundedHeight) > 1e-6 || Math.Abs(diameter - roundedDiameter) > 1e-6)
                    {
                        DebugLogger.Info($"[SleeveParameterService] [ROUNDING] Zone {zone?.Id}: " +
                            $"Width {RevitUnitConversionService.Instance.FromInternalMillimeters(width):F1}mm → {RevitUnitConversionService.Instance.FromInternalMillimeters(roundedWidth):F1}mm, " +
                            $"Height {RevitUnitConversionService.Instance.FromInternalMillimeters(height):F1}mm → {RevitUnitConversionService.Instance.FromInternalMillimeters(roundedHeight):F1}mm, " +
                            $"Diameter {RevitUnitConversionService.Instance.FromInternalMillimeters(diameter):F1}mm → {RevitUnitConversionService.Instance.FromInternalMillimeters(roundedDiameter):F1}mm");
                    }
                }
            }
            else
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Info($"[SleeveParameterService] [ROUNDING] Zone {zone?.Id}: DAMPER - No rounding applied, using exact calculated dimensions");
                }
            }
            
            return (roundedWidth, roundedHeight, roundedDiameter);
        }

        /// <summary>
        /// ✅ PARAMETER BATCHING: Set parameter value with batching support and diagnostic logging.
        /// Deferred when batching enabled, immediate when disabled.
        /// </summary>
        private bool SetParameter(
            FamilyInstance instance, 
            string parameterName, 
            double value, 
            ElementId currentSleeveId,
            string fallbackName = null)
        {
            var param = instance.LookupParameter(parameterName) ?? 
                       (fallbackName != null ? instance.LookupParameter(fallbackName) : null);
            
            if (param == null || param.IsReadOnly) return false;
            
            if (OptimizationFlags.UseBatchedParameterWrites)
            {
                if (!_deferredParameters.ContainsKey(currentSleeveId))
                    _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                
                // ✅ CRITICAL DIAGNOSTIC: Log if parameter is being overwritten
                bool isOverwrite = _deferredParameters[currentSleeveId].ContainsKey(parameterName);
                if (isOverwrite && (parameterName == "Width" || parameterName == "Height" || 
                                   parameterName == "Depth" || parameterName == "Wall Width"))
                {
                    var oldValue = _deferredParameters[currentSleeveId][parameterName];
                    SafeFileLogger.SafeAppendText("parameter_overwrite_debug.log",
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ⚠️ PARAMETER OVERWRITE: Sleeve {currentSleeveId.IntegerValue}, " +
                        $"Parameter='{parameterName}', OldValue={oldValue}, NewValue={value}\n");
                }
                
                _deferredParameters[currentSleeveId][parameterName] = value;
                return true;
            }
            else
            {
                param.Set(value);
                return true;
            }
        }

        /// <summary>
        /// ✅ FLAG MANAGEMENT SUPPORT: Set Sleeve Instance ID IMMEDIATELY (not deferred)
        /// </summary>
        private void SetSleeveInstanceId(FamilyInstance instance, ElementId currentSleeveId)
        {
            var sleeveInstanceIdParam = instance.LookupParameter("Sleeve Instance ID");
            if (sleeveInstanceIdParam != null && !sleeveInstanceIdParam.IsReadOnly)
            {
                // ✅ ALWAYS set immediately, regardless of batching flag - this is critical for flag management
                sleeveInstanceIdParam.Set(currentSleeveId.IntegerValue);
                
                // If batching is enabled, also add to deferred for consistency, but immediate set is critical
                if (OptimizationFlags.UseBatchedParameterWrites)
                {
                    if (!_deferredParameters.ContainsKey(currentSleeveId))
                        _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                    _deferredParameters[currentSleeveId]["Sleeve Instance ID"] = currentSleeveId.IntegerValue;
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SleeveParameterService] ✅ IMMEDIATE (CRITICAL): Set 'Sleeve Instance ID'={currentSleeveId.IntegerValue} for sleeve {currentSleeveId.IntegerValue}\n");
                }
            }
            else
            {
                // ⚠️ CRITICAL WARNING: Sleeve Instance ID parameter not found or read-only
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SleeveParameterService] ⚠️⚠️⚠️ CRITICAL: Cannot set 'Sleeve Instance ID' for sleeve {currentSleeveId.IntegerValue} - parameter not found or read-only!\n" +
                        $"  This will prevent flag management from identifying individual sleeves!\n");
                }
            }
        }

        /// <summary>
        /// ✅ CLUSTERING SUPPORT: Set MEP_ElementId and MEP_Category for clustering
        /// </summary>
        private void SetMepMetadata(FamilyInstance instance, ClashZone zone, ElementId currentSleeveId)
        {
            // Set MEP_Category - Required for clustering
            var mepCategoryParam = instance.LookupParameter("MEP_Category");
            if (mepCategoryParam != null && !mepCategoryParam.IsReadOnly)
            {
                if (OptimizationFlags.UseBatchedParameterWrites)
                {
                    if (!_deferredParameters.ContainsKey(currentSleeveId))
                        _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                    _deferredParameters[currentSleeveId]["MEP_Category"] = zone.MepElementCategory;
                }
                else
                {
                    mepCategoryParam.Set(zone.MepElementCategory);
                }
            }

            // ✅ CRITICAL: Set MEP_ElementId IMMEDIATELY (not deferred) - Required for clustering and corner retrieval
            var mepElementIdParam = instance.LookupParameter("MEP_ElementId");
            if (mepElementIdParam != null && !mepElementIdParam.IsReadOnly)
            {
                // ✅ ALWAYS set immediately, regardless of batching flag - this is critical for clustering
                mepElementIdParam.Set(zone.MepElementId.IntegerValue);
                
                // If batching is enabled, also add to deferred for consistency, but immediate set is critical
                if (OptimizationFlags.UseBatchedParameterWrites)
                {
                    if (!_deferredParameters.ContainsKey(currentSleeveId))
                        _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                    _deferredParameters[currentSleeveId]["MEP_ElementId"] = zone.MepElementId.IntegerValue;
                }
                
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SleeveParameterService] ✅ IMMEDIATE: Set 'MEP_ElementId'={zone.MepElementId.IntegerValue} for sleeve {currentSleeveId}\n");
                }
            }
            else
            {
                // ⚠️ CRITICAL WARNING: MEP_ElementId parameter not found or read-only
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [SleeveParameterService] ⚠️⚠️⚠️ CRITICAL: Cannot set 'MEP_ElementId' for sleeve {currentSleeveId} - parameter not found or read-only!\n" +
                        $"  This will prevent cluster sizing from finding corners in database!\n");
                }
            }
        }

        /// <summary>
        /// ✅ DAMPER ASYMMETRIC CLEARANCE: Set individual clearance parameters from zone
        /// </summary>
        private void SetDamperClearances(FamilyInstance instance, ClashZone zone, ElementId currentSleeveId)
        {
            // Set Clearance_Left
            if (zone.ClearanceLeft > 0)
            {
                SetClearanceParameter(instance, "Clearance_Left", zone.ClearanceLeft, currentSleeveId);
            }
            
            // Set Clearance_Right
            if (zone.ClearanceRight > 0)
            {
                SetClearanceParameter(instance, "Clearance_Right", zone.ClearanceRight, currentSleeveId);
            }
            
            // Set Clearance_Top
            if (zone.ClearanceTop > 0)
            {
                SetClearanceParameter(instance, "Clearance_Top", zone.ClearanceTop, currentSleeveId);
            }
            
            // Set Clearance_Bottom
            if (zone.ClearanceBottom > 0)
            {
                SetClearanceParameter(instance, "Clearance_Bottom", zone.ClearanceBottom, currentSleeveId);
            }
            
            if (!DeploymentConfiguration.DeploymentMode && 
                (zone.ClearanceLeft > 0 || zone.ClearanceRight > 0 || zone.ClearanceTop > 0 || zone.ClearanceBottom > 0))
            {
                DebugLogger.Info($"[SleeveParameterService] ✅ DAMPER CLEARANCES SET: Zone {zone.Id}, " +
                    $"L={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.ClearanceLeft):F1}mm, " +
                    $"R={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.ClearanceRight):F1}mm, " +
                    $"T={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.ClearanceTop):F1}mm, " +
                    $"B={RevitUnitConversionService.Instance.FromInternalMillimeters(zone.ClearanceBottom):F1}mm");
            }
        }

        /// <summary>
        /// Helper method to set clearance parameters with batching support
        /// </summary>
        private void SetClearanceParameter(FamilyInstance instance, string paramName, double value, ElementId currentSleeveId)
        {
            var param = instance.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly)
            {
                if (OptimizationFlags.UseBatchedParameterWrites)
                {
                    if (!_deferredParameters.ContainsKey(currentSleeveId))
                        _deferredParameters[currentSleeveId] = new Dictionary<string, object>();
                    _deferredParameters[currentSleeveId][paramName] = value;
                }
                else
                {
                    param.Set(value);
                }
            }
        }

        /// <summary>
        /// Get thickness based on host type
        /// </summary>
        private double GetThickness(ClashZone zone, bool isWallHost, bool isFramingHost)
        {
            if (isWallHost)
            {
                return zone.WallThickness > 0 ? zone.WallThickness : zone.StructuralElementThickness;
            }
            else if (isFramingHost)
            {
                return zone.FramingThickness > 0 ? zone.FramingThickness : zone.StructuralElementThickness;
            }
            else
            {
                return zone.StructuralElementThickness;
            }
        }

        /// <summary>
        /// ✅ CRITICAL FIX: Retrieve thickness from linked file if missing (PATH 3 fallback)
        /// </summary>
        private double RetrieveThicknessFromLinkedFile(ClashZone zone, double currentThickness)
        {
            try
            {
                // Try to get structural element from linked files
                Element structuralElement = null;
                var linkInstances = new FilteredElementCollector(_doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>();
                
                foreach (var linkInstance in linkInstances)
                {
                    var linkDoc = linkInstance.GetLinkDocument();
                    if (linkDoc != null)
                    {
                        try
                        {
                            structuralElement = linkDoc.GetElement(new ElementId(zone.StructuralElementIdValue));
                            if (structuralElement != null)
                            {
                                // Retrieve thickness from linked file element
                                if (structuralElement is Wall wall)
                                {
                                    currentThickness = wall.Width;
                                }
                                else if (structuralElement is Floor floor)
                                {
                                    currentThickness = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)?.AsDouble() ?? 0.0;
                                }
                                else if (structuralElement?.Category?.Id?.IntegerValue == (int)BuiltInCategory.OST_StructuralFraming)
                                {
                                    // Get framing thickness from type parameter 'b'
                                    var typeId = (structuralElement as FamilyInstance)?.GetTypeId() ?? structuralElement.GetTypeId();
                                    var typeElem = linkDoc.GetElement(typeId);
                                    if (typeElem != null)
                                    {
                                        var param = typeElem.LookupParameter("b") ?? typeElem.LookupParameter("B");
                                        if (param != null) currentThickness = param.AsDouble();
                                    }
                                }
                                
                                if (currentThickness > 0.0)
                                {
                                    if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Info($"[SleeveParameterService] [DEPTH-FALLBACK] Retrieved thickness={currentThickness * 304.8:F1}mm from linked file for structural element {zone.StructuralElementIdValue} (Zone={zone.Id})");
                                    break;
                                }
                            }
                        }
                        catch { continue; }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Warning($"[SleeveParameterService] [DEPTH-FALLBACK] Error retrieving thickness from linked file for Zone={zone.Id}: {ex.Message}");
            }
            
            return currentThickness;
        }
    }
}

