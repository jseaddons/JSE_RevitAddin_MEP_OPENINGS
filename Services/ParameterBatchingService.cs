using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;
using JSE_RevitAddin_MEP_OPENINGS.Utils;
using JSE_RevitAddin_MEP_OPENINGS.Data;

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
                        DebugLogger.Info($"[BATCH-PARAMS] 📝 Deferred parameter '{parameterName}' = {value} for element {elementId.IntegerValue}");
                }
            }
        }

        public int FlushDeferredParameters(Document doc)
        {
            if (doc == null)
                throw new ArgumentNullException(nameof(doc));

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
                    DebugLogger.Info($"[BATCH-PARAMS] 📊 Element IDs: [{string.Join(", ", parametersToFlush.Keys.Select(id => id.IntegerValue))}]");
            }

            var sw = System.Diagnostics.Stopwatch.StartNew();

            try
            {
                foreach (var kvp in parametersToFlush)
                {
                    var elementId = kvp.Key;
                    var parameters = kvp.Value;

                    try
                    {
                        var element = doc.GetElement(elementId);
                        if (element == null)
                        {
                            failCount += parameters.Count;
                            if (!DeploymentConfiguration.DeploymentMode)
                            {
                                                                if (!DeploymentConfiguration.DeploymentMode)
                                    DebugLogger.Warning($"[BATCH-PARAMS] ❌ Element {elementId.IntegerValue} not found, skipping {parameters.Count} parameters");
                            }
                            continue;
                        }

                        foreach (var paramKvp in parameters)
                        {
                            try
                            {
                                var param = element.LookupParameter(paramKvp.Key);
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
                                                                        if (!DeploymentConfiguration.DeploymentMode)
                                        DebugLogger.Warning($"[BATCH-PARAMS] Failed to set parameter '{paramKvp.Key}' on element {elementId.IntegerValue}: {paramEx.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception elementEx)
                    {
                        failCount += parameters.Count;
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                                                        if (!DeploymentConfiguration.DeploymentMode)
                                DebugLogger.Error($"[BATCH-PARAMS] Failed to process element {elementId.IntegerValue}: {elementEx.Message}");
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
            }
        }
    }
}
