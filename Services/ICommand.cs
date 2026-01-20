using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Base interface for all Revit commands
    /// </summary>
    public interface ICommand
    {
        void Execute(UIApplication app);
    }
}

