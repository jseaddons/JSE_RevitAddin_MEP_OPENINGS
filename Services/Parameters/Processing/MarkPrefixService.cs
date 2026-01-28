using System;
using JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Configuration;
using Autodesk.Revit.DB;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Parameters.Processing
{
    // Dummy service to satisfy UI references
    public class MarkPrefixService
    {
        private Document _doc;
        private Action<string> _logger;

        public MarkPrefixService(Document doc, Action<string> logger = null)
        {
            _doc = doc;
            _logger = logger;
        }

        public static MarkPrefixSettings GetCurrentPrefixes()
        {
            return new MarkPrefixSettings();
        }

        // Dummy methods to satisfy any remaining MarkParameterCommand logic
        public (int, int) ApplyPrefixesOnly(Document doc, string category, string projectPrefix, string disciplinePrefix, bool remarkFlag, object settings)
        {
            return (0, 0); 
        }

        public (int, int) ApplyNumbersBatch(Document doc, string format, object filter, object settings)
        {
            return (0, 0);
        }
    }
}
