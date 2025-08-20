using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Net;
using System.Net.NetworkInformation;
using System.Linq;

namespace JSE_RevitAddin_MEP_OPENINGS.Security
{
    public static class LicenseValidator
    {
        private const string REQUIRED_DOMAIN = "jse24";
        private const string REQUIRED_DOMAIN_ALTERNATIVE = "jse"; // Fallback for backward compatibility
        
        /// <summary>
        /// Validates that the application is running on a computer in the required domain
        /// </summary>
        public static bool ValidateLicense()
        {
            try
            {
                // Primary check: Must be on jse24 domain
                string domain = Environment.UserDomainName;
                string machineName = Environment.MachineName;
                
                LogValidationAttempt(domain, machineName);
                
                // Check for exact jse24 domain match
                if (domain.Equals(REQUIRED_DOMAIN, StringComparison.OrdinalIgnoreCase))
                {
                    LogValidationSuccess(domain, machineName, "Exact domain match");
                    return true;
                }
                
                // Check for jse24 as part of domain name
                if (domain.ToUpper().Contains(REQUIRED_DOMAIN.ToUpper()))
                {
                    LogValidationSuccess(domain, machineName, "Domain contains required name");
                    return true;
                }
                
                // Fallback check for backward compatibility (contains "jse")
                if (domain.ToUpper().Contains(REQUIRED_DOMAIN_ALTERNATIVE.ToUpper()))
                {
                    LogValidationSuccess(domain, machineName, "Fallback domain check");
                    return true;
                }
                
                LogValidationFailure(domain, machineName, "Domain validation failed");
                return false;
            }
            catch (Exception ex)
            {
                LogValidationError(ex);
                return false;
            }
        }

        /// <summary>
        /// Validates hardware and network configuration
        /// </summary>
        public static bool ValidateHardwareId()
        {
            try
            {
                string domain = Environment.UserDomainName;
                string machineName = Environment.MachineName;
                
                // Additional network validation - check if we can resolve jse24 domain
                if (CanResolveDomain(REQUIRED_DOMAIN))
                {
                    LogValidationSuccess(domain, machineName, "Domain resolution successful");
                    return true;
                }
                
                // If domain resolution fails, fall back to basic domain check
                return ValidateLicense();
            }
            catch (Exception ex)
            {
                LogValidationError(ex);
                return false;
            }
        }
        
        /// <summary>
        /// Checks if the specified domain can be resolved
        /// </summary>
        private static bool CanResolveDomain(string domainName)
        {
            try
            {
                // Try to resolve the domain
                IPHostEntry hostEntry = Dns.GetHostEntry(domainName);
                return hostEntry.AddressList.Length > 0;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Gets minimal domain information for logging only
        /// </summary>
        public static string GetDomainInfo()
        {
            try
            {
                string domain = Environment.UserDomainName;
                return domain;
            }
            catch
            {
                return "Unknown";
            }
        }
        
        /// <summary>
        /// Logs validation attempts for security auditing
        /// </summary>
        private static void LogValidationAttempt(string domain, string machineName)
        {
            string logMessage = $"[DOMAIN_VALIDATION] Attempting validation - Domain: {domain}, Machine: {machineName}";
            WriteToSecurityLog(logMessage);
        }
        
        /// <summary>
        /// Logs successful validation
        /// </summary>
        private static void LogValidationSuccess(string domain, string machineName, string reason)
        {
            string logMessage = $"[DOMAIN_VALIDATION] SUCCESS - Domain: {domain}, Machine: {machineName}, Reason: {reason}";
            WriteToSecurityLog(logMessage);
        }
        
        /// <summary>
        /// Logs validation failure
        /// </summary>
        private static void LogValidationFailure(string domain, string machineName, string reason)
        {
            string logMessage = $"[DOMAIN_VALIDATION] FAILED - Domain: {domain}, Machine: {machineName}, Reason: {reason}";
            WriteToSecurityLog(logMessage);
        }
        
        /// <summary>
        /// Logs validation errors
        /// </summary>
        private static void LogValidationError(Exception ex)
        {
            string logMessage = $"[DOMAIN_VALIDATION] ERROR - {ex.Message}";
            WriteToSecurityLog(logMessage);
        }
        
        /// <summary>
        /// Writes security events to a log file
        /// </summary>
        private static void WriteToSecurityLog(string message)
        {
            try
            {
                string logDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "JSE_RevitAddin", "Security");
                if (!Directory.Exists(logDir))
                    Directory.CreateDirectory(logDir);
                
                string logFile = Path.Combine(logDir, "domain_validation.log");
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string logEntry = $"[{timestamp}] {message}";
                
                File.AppendAllText(logFile, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently fail if logging fails
            }
        }
    }
}
