using System;
using System.Net;
using System.Net.NetworkInformation;
using System.Linq; // Added for .Where()

namespace JSE_RevitAddin_MEP_OPENINGS.Security
{
    /// <summary>
    /// Comprehensive domain validation for jse24 domain requirements
    /// </summary>
    public static class DomainValidator
    {
        private const string REQUIRED_DOMAIN = "jse24";
        private const string REQUIRED_DOMAIN_ALTERNATIVE = "jse";
        
        /// <summary>
        /// Performs comprehensive domain validation
        /// </summary>
        public static DomainValidationResult ValidateDomain()
        {
            var result = new DomainValidationResult();
            
            try
            {
                // Basic domain information
                result.CurrentDomain = Environment.UserDomainName;
                result.MachineName = Environment.MachineName;
                result.UserName = Environment.UserName;
                result.ValidationTime = DateTime.Now;
                
                // Check if we're on a domain at all
                if (string.IsNullOrEmpty(result.CurrentDomain))
                {
                    result.IsValid = false;
                    result.FailureReason = "Not connected to any domain";
                    return result;
                }
                
                // Primary validation: exact jse24 domain match
                if (result.CurrentDomain.Equals(REQUIRED_DOMAIN, StringComparison.OrdinalIgnoreCase))
                {
                    result.IsValid = true;
                    result.ValidationMethod = "Exact domain match";
                    return result;
                }
                
                // Secondary validation: domain contains jse24
                if (result.CurrentDomain.ToUpper().Contains(REQUIRED_DOMAIN.ToUpper()))
                {
                    result.IsValid = true;
                    result.ValidationMethod = "Domain contains required name";
                    return result;
                }
                
                // Fallback validation: domain contains jse (for backward compatibility)
                if (result.CurrentDomain.ToUpper().Contains(REQUIRED_DOMAIN_ALTERNATIVE.ToUpper()))
                {
                    result.IsValid = true;
                    result.ValidationMethod = "Fallback domain check";
                    return result;
                }
                
                // Additional network validation
                if (CanResolveJse24Domain())
                {
                    result.IsValid = true;
                    result.ValidationMethod = "Network domain resolution";
                    return result;
                }
                
                // All validations failed
                result.IsValid = false;
                result.FailureReason = $"Domain '{result.CurrentDomain}' does not meet requirements";
                return result;
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.FailureReason = $"Validation error: {ex.Message}";
                result.Exception = ex;
                return result;
            }
        }
        
        /// <summary>
        /// Checks if the jse24 domain can be resolved via DNS
        /// </summary>
        private static bool CanResolveJse24Domain()
        {
            try
            {
                // Try to resolve the jse24 domain
                IPHostEntry hostEntry = Dns.GetHostEntry(REQUIRED_DOMAIN);
                return hostEntry.AddressList.Length > 0;
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// Attempts to validate domain membership using Active Directory
        /// Note: This method is disabled due to namespace compatibility issues
        /// </summary>
        public static bool ValidateActiveDirectoryMembership()
        {
            // Active Directory validation disabled for compatibility
            return false;
        }
        
        /// <summary>
        /// Gets minimal network information for logging only
        /// </summary>
        public static string GetNetworkInfo()
        {
            try
            {
                string info = $"Domain: {Environment.UserDomainName}";
                return info;
            }
            catch
            {
                return "Unknown";
            }
        }
    }
    
    /// <summary>
    /// Result of domain validation
    /// </summary>
    public class DomainValidationResult
    {
        public bool IsValid { get; set; }
        public string CurrentDomain { get; set; } = string.Empty;
        public string MachineName { get; set; } = string.Empty;
        public string UserName { get; set; } = string.Empty;
        public string ValidationMethod { get; set; } = string.Empty;
        public string FailureReason { get; set; } = string.Empty;
        public DateTime ValidationTime { get; set; }
        public Exception? Exception { get; set; }
        
        public override string ToString()
        {
            if (IsValid)
            {
                return $"Domain validation SUCCESS - Domain: {CurrentDomain}, Method: {ValidationMethod}, Time: {ValidationTime}";
            }
            else
            {
                return $"Domain validation FAILED - Domain: {CurrentDomain}, Reason: {FailureReason}, Time: {ValidationTime}";
            }
        }
    }
}
