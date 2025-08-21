using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Security
{
    public static class LicenseValidator
    {
        private const string REQUIRED_DOMAIN = "jse24";
        private const string REQUIRED_DOMAIN_ALTERNATIVE = "jse";

        public static bool ValidateLicense()
        {
            try
            {
                // Check if we're in the required domain
                if (!ValidateDomain())
                {
                    // Show error message to user
                    TaskDialog.Show("License Error", 
                        "Not Licensed\n\nPlease contact admin@jseeng.com");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                // Show error message to user even on exception
                TaskDialog.Show("License Error", 
                    "Not Licensed\n\nPlease contact admin@jseeng.com");
                return false;
            }
        }

        private static bool ValidateDomain()
        {
            try
            {
                string domain = Environment.UserDomainName;
                
                // Check for exact domain match first
                if (string.Equals(domain, REQUIRED_DOMAIN, StringComparison.OrdinalIgnoreCase))
                    return true;
                
                // Check for alternative domain
                if (string.Equals(domain, REQUIRED_DOMAIN_ALTERNATIVE, StringComparison.OrdinalIgnoreCase))
                    return true;
                
                // Check if domain contains required text
                if (domain.ToUpper().Contains(REQUIRED_DOMAIN.ToUpper()))
                    return true;
                
                return false;
            }
            catch
            {
                return false;
            }
        }

        public static bool ValidateHardwareId()
        {
            try
            {
                // Simple machine name check - allow any machine on JSE domain
                string machineName = Environment.MachineName;
                string domain = Environment.UserDomainName;
                
                // Allow if on JSE domain
                return domain.ToUpper().Contains("JSE");
            }
            catch
            {
                return false;
            }
        }
    }
}
