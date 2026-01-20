using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.InsulationDetection
{
    /// <summary>
    /// Detects insulation status and thickness from MEP elements.
    /// Follows Single Responsibility Principle - only responsible for insulation detection.
    /// Follows Open/Closed Principle - can be extended for new insulation types without modification.
    /// </summary>
    public class InsulationDetector : IInsulationDetector
    {
        private const double InsulationThreshold = 0.001; // ~0.3mm threshold for considering element insulated

        /// <summary>
        /// Detects if an element is insulated by checking insulation parameters.
        /// Returns true if insulation thickness > threshold (0.001ft ≈ 0.3mm).
        /// </summary>
        public bool IsInsulated(Element element)
        {
            if (element == null)
                return false;

            double thickness = GetInsulationThickness(element);
            return thickness > InsulationThreshold;
        }

        /// <summary>
        /// Gets the insulation thickness for an element (in Revit internal units - feet).
        /// Returns 0.0 if element is not insulated or thickness cannot be determined.
        /// </summary>
        public double GetInsulationThickness(Element element)
        {
            if (element == null)
                return 0.0;

            // Check for InsulationThickness parameter (common for ducts and pipes)
            // Try multiple parameter name variations
            var insulationParam = element.LookupParameter("InsulationThickness") ??
                                 element.LookupParameter("Insulation Thickness") ??
                                 element.LookupParameter("Insulation");

            if (insulationParam != null && insulationParam.HasValue)
            {
                double thickness = insulationParam.AsDouble();
                return thickness > 0.0 ? thickness : 0.0;
            }

            return 0.0;
        }

        /// <summary>
        /// Gets insulation information from MepElementSize if available, otherwise detects from element.
        /// Returns (isInsulated, thickness) tuple.
        /// Priority: MepElementSize (already analyzed by strategy) > Direct element detection
        /// </summary>
        public (bool isInsulated, double thickness) GetInsulationInfo(Element element, MepElementSize mepElementSize)
        {
            // ✅ OOP PRINCIPLE: Prefer strategy-analyzed data over direct inspection
            // MepElementSize is populated by strategies which have already done proper analysis
            if (mepElementSize != null && mepElementSize.IsInsulated && mepElementSize.InsulationThickness > InsulationThreshold)
            {
                return (true, mepElementSize.InsulationThickness);
            }

            // Fallback: Direct element detection if MepElementSize doesn't have insulation info
            double thickness = GetInsulationThickness(element);
            bool isInsulated = thickness > InsulationThreshold;

            return (isInsulated, thickness);
        }
    }
}

