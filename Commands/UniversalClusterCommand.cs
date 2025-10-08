using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    /// <summary>
    /// Universal cluster command wrapper - triggers RectangularSleeveClusterCommandV2 manually
    /// 
    /// ⚠️ WORKAROUND: Cannot construct ExternalCommandData from IExternalEventHandler
    /// Solution: User must manually trigger cluster command from Revit ribbon
    /// OR: Extract clustering logic to a service class (future enhancement)
    /// 
    /// For now: This is a placeholder that logs that clustering should happen
    /// </summary>
    public class UniversalClusterCommand : ICommand
    {
        private readonly string _targetCategory;
        
        public UniversalClusterCommand(string targetCategory)
        {
            _targetCategory = targetCategory ?? throw new ArgumentNullException(nameof(targetCategory));
        }
        
        public void Execute(UIApplication app)
        {
            try
            {
                DebugLogger.Info($"[UniversalClusterCommand] ⚠️ Clustering for {_targetCategory} - MANUAL TRIGGER REQUIRED");
                DebugLogger.Info($"[UniversalClusterCommand] Please run 'Rectangular Cluster Command V2' from Revit ribbon to cluster {_targetCategory} sleeves");
                DebugLogger.Info($"[UniversalClusterCommand] Automated clustering will be available in next update");
                
                // TODO: Extract clustering logic from RectangularSleeveClusterCommandV2 into a service
                // Then call: var service = new UniversalClusterService();
                //           service.ClusterSleeves(app.ActiveUIDocument.Document, _targetCategory);
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[UniversalClusterCommand] Error: {ex.Message}");
            }
        }
    }
}

