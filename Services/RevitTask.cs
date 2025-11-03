using System;
using System.Collections.Concurrent;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Industry-standard queue from any thread into Revit's main thread.
    /// Copy-paste from major GitHub projects (pyRevit, RevitMCPSDK, etc.)
    /// </summary>
    public static class RevitTask
    {
        private static ExternalEvent? _exEvent;
        private static readonly ConcurrentQueue<Action<UIApplication>> _queue = new();

        public static void Init(ExternalEvent exEvent)
        {
            _exEvent = exEvent;
        }

        public static void Run(Action<UIApplication> action)
        {
            if (_exEvent == null)
                throw new InvalidOperationException("RevitTask not initialized. Call Init() first.");

            _queue.Enqueue(action);
            _exEvent.Raise();   // → Revit will call Execute on main thread
        }

        public static bool IsInitialized()
        {
            return _exEvent != null;
        }

        // ExternalEventHandler implementation
        public class Handler : IExternalEventHandler
        {
            public void Execute(UIApplication app)
            {
                while (_queue.TryDequeue(out var action))
                {
                    try
                    {
                        action(app);
                    }
                    catch (Exception ex)
                    {
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[RevitTask] Exception in queued action: {ex.Message}");
                                                if (!DeploymentConfiguration.DeploymentMode)
                            DebugLogger.Error($"[RevitTask] Stack trace: {ex.StackTrace}");
                        // Continue processing other queued actions
                    }
                }
            }

            public string GetName() => "RevitTask";
        }
    }
}
