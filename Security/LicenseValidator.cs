using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace JSE_RevitAddin_MEP_OPENINGS.Security
{
    public static class LicenseValidator
    {
        public static bool ValidateLicense()
        {
            try
            {
                // Simple domain check - only allow computers on JSE domain
                string domain = Environment.UserDomainName;
                return domain.ToUpper().Contains("JSE");
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
