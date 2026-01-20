using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refresh
{
    /// <summary>
    /// Focused interface for document-related context.
    /// Follows Interface Segregation Principle - clients only depend on what they need.
    /// </summary>
    public interface IRefreshDocumentContext
    {
        Document Document { get; }
        UIDocument UIDocument { get; }
    }
    
    /// <summary>
    /// Focused interface for caching-related context.
    /// Follows Interface Segregation Principle - clients only depend on what they need.
    /// </summary>
    public interface IRefreshCacheContext
    {
        StringPool StringPool { get; }
        GeometryCache GeometryCache { get; }
        XmlCache XmlCache { get; }
    }
    
    /// <summary>
    /// Focused interface for UI selection context.
    /// Follows Interface Segregation Principle - clients only depend on what they need.
    /// </summary>
    public interface IRefreshSelectionContext
    {
        List<string> SelectedFilterNames { get; }
        List<string> SelectedMepCategories { get; }
        List<string> SelectedReferenceFiles { get; }
        List<string> SelectedHostFiles { get; }
        List<string> SelectedHostTypes { get; }
        Dictionary<string, double> ClearanceSettings { get; }
    }
    
    /// <summary>
    /// Focused interface for settings context.
    /// Follows Interface Segregation Principle - clients only depend on what they need.
    /// </summary>
    public interface IRefreshSettingsContext
    {
        bool EnableThreePointValidation { get; set; }
        bool IsDeploymentMode { get; set; }
    }
}
