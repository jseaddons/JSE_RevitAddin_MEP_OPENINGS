using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Batches parameter writes for 4-6× performance improvement.
    /// Implements Single Responsibility Principle - only handles parameter batching.
    /// Thread-safe for concurrent parameter queueing.
    /// </summary>
    public class ParameterBatchingService : IParameterBatchingService
    {
        // Key: ElementId of sleeve instance
        // Value: Dictionary of parameter name → value (double or string)
        private readonly Dictionary<ElementId, Dictionary<string, object>> _deferredParameters;
        private readonly object _lock = new object();
        
        // ✅ SAFETY FLAG: Prevents multiple flushes (critical for performance)
        private bool _hasFlushed = false;

        public ParameterBatchingService()
        {
            _deferredParameters = new Dictionary<ElementId, Dictionary<string, object>>();
        }

        public bool IsBatchingEnabled => OptimizationFlags.UseBatchedParameterWrites;

        public int DeferredElementCount
        {
            get
            {
                lock (_lock)
                {
                    return _deferredParameters.Count;
                }
            }
        }

        public int DeferredParameterCount
        {
            get
            {
                lock (_lock)
                {
                    return _deferredParameters.Values.Sum(d => d.Count);
                }
            }
        }

        public void DeferParameter(ElementId elementId, string parameterName, object value)
        {
            if (elementId == null || elementId == ElementId.InvalidElementId)
                return;

            if (string.IsNullOrEmpty(parameterName))
                return;

            lock (_lock)
            {
                if (!_deferredParameters.ContainsKey(elementId))
                {
                    _deferredParameters[elementId] = new Dictionary<string, object>();
                }

                _deferredParameters[elementId][parameterName] = value;

                // Diagnostic logging for small batches
                if (!DeploymentConfiguration.DeploymentMode && _deferredParameters.Count <= 5)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[BATCH-PARAMS] 📝 Deferred parameter '{parameterName}' = {value} for element {elementId.GetIntegerValue()}");
                }
            }
        }

        public int FlushDeferredParameters(Document doc)
        {
            if (doc == null)
                throw new ArgumentNullException(nameof(doc));

            // ✅ SAFETY FLAG: Prevent multiple flushes (critical for performance - only flush once!)
            lock (_lock)
            {
                if (_hasFlushed)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                        var stackTrace = new System.Diagnostics.StackTrace(skipFrames: 1, fNeedFileInfo: false);
                        var caller = stackTrace.GetFrame(0)?.GetMethod()?.Name ?? "Unknown";
                        DebugLogger.Warning($"[BATCH-PARAMS] ⚠️ SAFETY: FlushDeferredParameters called AGAIN from {caller} - IGNORING (already flushed once). This indicates a bug - parameters should only flush once!");
                    }
                    return 0; // ✅ CRITICAL: Exit early to prevent duplicate flushes
                }
            }

            Dictionary<ElementId, Dictionary<string, object>> parametersToFlush;
            
            lock (_lock)
            {
                if (_deferredParameters.Count == 0)
                {
                    if (!DeploymentConfiguration.DeploymentMode)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Info("[BATCH-PARAMS] ⚠️ FlushDeferredParameters called but no parameters deferred");
                    }
                    _hasFlushed = true; // Mark as flushed even if empty to prevent retries
                    return 0;
                }

                // Create a copy to avoid holding lock during Revit API calls
                parametersToFlush = new Dictionary<ElementId, Dictionary<string, object>>(_deferredParameters);
            }

            int totalParams = parametersToFlush.Values.Sum(d => d.Count);
            int successCount = 0;
            int failCount = 0;

            if (!DeploymentConfiguration.DeploymentMode)
            {
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[BATCH-PARAMS] 🔄 Flushing {parametersToFlush.Count} elements with {totalParams} total parameters...");
                                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info($"[BATCH-PARAMS] 📊 Element IDs: [{string.Join(", ", parametersToFlush.Keys.Select(id => id.GetIntegerValue()))}]");
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                // ✅ OPTIMIZATION 5: Pre-cache all elements and parameters to avoid repeated API calls
                var elementCache = new Dictionary<ElementId, Element>();
                var parameterCache = new Dictionary<ElementId, Dictionary<string, Parameter>>();

                foreach (var elementId in parametersToFlush.Keys)
                {
                    try
                    {
                        var element = doc.GetElement(elementId);
                        if (element != null && element.IsValidObject)
                        {
                            elementCache[elementId] = element;
                            
                            // Pre-cache parameters for this element
                            if (parametersToFlush.TryGetValue(elementId, out var paramsDict))
                            {
                                var paramDict = new Dictionary<string, Parameter>();
                                foreach (var paramName in paramsDict.Keys)
                                {
                                    var param = element.LookupParameter(paramName);
                                    if (param != null)
                                    {
                                        paramDict[paramName] = param;
                                    }
                                }
                                if (paramDict.Count > 0)
                                {
                                    parameterCache[elementId] = paramDict;
                                }
                            }
                        }
                    }
                    catch { }
                }

                if (!DeploymentConfiguration.DeploymentMode && elementCache.Count > 0)
                {
                    DebugLogger.Info($"[BATCH-PARAMS] ✅ Pre-cached {elementCache.Count} elements and {parameterCache.Count} parameter sets for flush");
                }

                foreach (var kvp in parametersToFlush)
                {
                    var elementId = kvp.Key;
                    var parameters = kvp.Value;

                    try
                    {
                        // ✅ OPTIMIZATION 5: Use cached element
                        if (!elementCache.TryGetValue(elementId, out var element))
                        {
                            failCount += parameters.Count;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                DebugLogger.Warning($"[BATCH-PARAMS] ❌ Element {elementId.GetIntegerValue()} not found, skipping {parameters.Count} parameters");
                            }
                            continue;
                        }

                        foreach (var paramKvp in parameters)
                        {
                            try
                            {
                                // ✅ OPTIMIZATION 5: Use cached parameter if available
                                Parameter param = null;
                                if (parameterCache.TryGetValue(elementId, out var paramDict) && paramDict.TryGetValue(paramKvp.Key, out param))
                                {
                                    // Use cached parameter
                                }
                                else
                                {
                                    param = element.LookupParameter(paramKvp.Key);
                                }

                                if (param == null || param.IsReadOnly)
                                {
                                    failCount++;
                                    continue;
                                }

                                bool setSuccess = false;

                                if (paramKvp.Value is double doubleValue)
                                {
                                    setSuccess = param.Set(doubleValue);
                                }
                                else if (paramKvp.Value is string stringValue)
                                {
                                    setSuccess = param.Set(stringValue);
                                }
                                else if (paramKvp.Value is int intValue)
                                {
                                    setSuccess = param.Set(intValue);
                                }

                                if (setSuccess)
                                {
                                    successCount++;
                                }
                                else
                                {
                                    failCount++;
                                }
                            }
                            catch (Exception paramEx)
                            {
                                failCount++;
                                if (!DeploymentConfiguration.DeploymentMode)
                                {
                                    DebugLogger.Warning($"[BATCH-PARAMS] Failed to set parameter '{paramKvp.Key}' on element {elementId.GetIntegerValue()}: {paramEx.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception elementEx)
                    {
                        failCount += parameters.Count;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Error($"[BATCH-PARAMS] Failed to process element {elementId.GetIntegerValue()}: {elementEx.Message}");
                        }
                    }
                }

                sw.Stop();

                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[BATCH-PARAMS] ✅ Flush complete: {successCount} successful, {failCount} failed in {sw.ElapsedMilliseconds}ms");
                    
                    // Log performance to separate file
                    SafeFileLogger.SafeAppendText("parameter_batching_performance.log",
                        $"[{DateTime.Now:HH:mm:ss}] Flushed {parametersToFlush.Count} elements, {successCount} params, {sw.ElapsedMilliseconds}ms");
                }
            }
            catch (Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Error($"[BATCH-PARAMS] Critical error during flush: {ex.Message}");
                }
                throw;
            }
            finally
            {
                // ✅ SAFETY FLAG: Mark as flushed to prevent duplicate flushes (CRITICAL for performance)
                lock (_lock)
                {
                    _hasFlushed = true;
                }
                
                // Clear deferred parameters after flush attempt
                Clear();
            }

            return successCount;
        }

        public void Clear()
        {
            lock (_lock)
            {
                if (!DeploymentConfiguration.DeploymentMode && _deferredParameters.Count > 0)
                {
                                        if (!DeploymentConfiguration.DeploymentMode)
                        DebugLogger.Info($"[BATCH-PARAMS] 🗑️ Clearing {_deferredParameters.Count} deferred elements");
                }

                _deferredParameters.Clear();
                // ✅ CRITICAL FIX: DO NOT reset _hasFlushed here - it should only be reset at the START of a new placement run
                // Resetting it here allows multiple flushes during the same placement run, breaking batching optimization
                // The flag will be reset when a new placement run starts (in UniversalSleevePlacerService.PlaceSleeves)
            }
        }
        
        /// <summary>
        /// Reset the flush flag - should only be called at the start of a new placement run
        /// </summary>
        public void ResetFlushFlag()
        {
            lock (_lock)
            {
                _hasFlushed = false;
                if (!DeploymentConfiguration.DeploymentMode)
                    DebugLogger.Info("[BATCH-PARAMS] 🔄 Reset flush flag (new placement run starting)");
            }
        }
    }
}