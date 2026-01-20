using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Services.Interfaces;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Refactored
{
    /// <summary>
    /// ✅ SOLID REFACTORED: Service for validating document state before operations.
    /// Extracted from UniversalSleevePlacementCommand to adhere to SRP.
    /// </summary>
    public class DocumentValidatorService : IDocumentValidator
    {
        public bool ValidateDocument(Document doc, string logPrefix)
        {
            if (!DeploymentConfiguration.DeploymentMode)
            {
                DebugLogger.Info($"{logPrefix} Document validation - Path: {doc.PathName}");
                DebugLogger.Info($"{logPrefix} Document validation - IsModifiable: {doc.IsModifiable}");
                DebugLogger.Info($"{logPrefix} Document validation - IsLinked: {doc.IsLinked}");
                DebugLogger.Info($"{logPrefix} Document validation - IsWorkshared: {doc.IsWorkshared}");
            }

            // Test if we can modify
            bool canModify = false;
            try
            {
                using (var testTransaction = new Transaction(doc, "Test Modification"))
                {
                    if (testTransaction.Start() == TransactionStatus.Started)
                    {
                        canModify = true;
                        testTransaction.RollBack();
                        if (!DeploymentConfiguration.DeploymentMode)
                        {
                            DebugLogger.Info($"{logPrefix} Document can be modified");
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    DebugLogger.Error($"{logPrefix} Document modification test failed: {ex.Message}");
                }
            }

            return canModify;
        }
    }
}

