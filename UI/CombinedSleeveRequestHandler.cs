using System;
using System.Collections.Generic;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.UI
{
    /// <summary>
    /// Handles external events from the modeless Combined Sleeve UI to ensure valid Revit API context.
    /// </summary>
    public class CombinedSleeveRequestHandler : IExternalEventHandler
    {
        private Action<UIApplication> _action;

        public CombinedSleeveRequestHandler()
        {
        }

        /// <summary>
        /// Set the action to be executed in the next Raise() call.
        /// </summary>
        public void SetAction(Action<UIApplication> action)
        {
            _action = action;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                if (_action != null)
                {
                    _action(app);
                }
            }
            catch (Exception ex)
            {
                // Basic logging if DebugLogger is available, otherwise silent
                // DebugLogger.Error("External Event Error: " + ex.Message);
            }
        }

        public string GetName()
        {
            return "CombinedSleeveRequestHandler";
        }
    }
}
