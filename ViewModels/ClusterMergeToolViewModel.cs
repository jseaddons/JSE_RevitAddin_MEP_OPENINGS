using System.Windows.Input;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.ViewModels
{
    public class ClusterMergeToolViewModel
    {
        public ICommand ClusterMergeCommand { get; }

        public ClusterMergeToolViewModel()
        {
            ClusterMergeCommand = new RelayCommand(ExecuteClusterMerge);
        }

        private void ExecuteClusterMerge()
        {
            // TODO: Call the ClusterMergeCommand
            TaskDialog.Show("Cluster Merge", "Cluster merge tool executed from ViewModel.");
        }
    }
}
