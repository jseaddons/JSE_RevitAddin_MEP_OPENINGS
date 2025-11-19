/// This is a backup of the IntersectionResult class from MEP-Structural Intersection Detector.cs (Line 511-527)
/// Backup created before removing unused IntersectionResult references
/// If the app fails to work after removal, this can be restored.
/// 
private class IntersectionResult
{
    public Element MepElement { get; set; }
    public Element StructuralElement { get; set; }
    public string MepCategory { get; set; }
    public string StructuralCategory { get; set; }
    public ElementId MepId { get; set; }
    public ElementId WallId { get; set; }
    public XYZ IntersectionCenter { get; set; }
}
