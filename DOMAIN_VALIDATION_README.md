# Domain Validation System for JSE24 Domain

## Overview

This application has been configured to run **ONLY** on computers that are members of domains with the name "jse24". This is a security measure to ensure the application is only used within authorized network environments.

## Domain Requirements

### Primary Domain
- **jse24** - Exact domain name match (preferred)

### Alternative Domains (for backward compatibility)
- Any domain containing "jse24" in the name
- Any domain containing "jse" in the name

### Examples of Valid Domains
- `jse24.local`
- `jse24.corp`
- `jse24.company.com`
- `jse.corp` (legacy support)
- `jse.local` (legacy support)

### Examples of Invalid Domains
- `company.local`
- `office.corp`
- `workgroup` (no domain)
- `localhost` (local machine only)

## How It Works

### 1. Silent Application Startup Validation
When the Revit add-in starts, it automatically:
- Checks the current computer's domain membership
- Validates against the required domain names
- Logs all validation attempts for security auditing
- **Silently prevents the application from loading if validation fails**
- **No error messages, no user notifications, completely invisible to users**

### 2. Validation Methods
The system uses multiple validation approaches:

#### Primary Validation
- **Environment Check**: `Environment.UserDomainName`
- **Exact Match**: Direct comparison with "jse24"
- **Contains Check**: Domain name contains "jse24"

#### Secondary Validation
- **Network Resolution**: DNS lookup for jse24 domain
- **Active Directory**: Domain controller connectivity check
- **Fallback**: Legacy "jse" domain support

### 3. Security Logging
All validation attempts are logged to:
```
%APPDATA%\JSE_RevitAddin\Security\domain_validation.log
```

Log entries include:
- Timestamp
- Domain name
- Machine name
- Validation result
- Validation method
- Failure reasons (if applicable)

## Silent Operation

### No User Notifications
If domain validation fails:
- **No error dialogs are shown**
- **No user notifications appear**
- **The add-in simply doesn't load**
- **Users will not see any buttons or commands**
- **Completely invisible validation process**

### What Users Experience
- **On jse24 domain**: Add-in loads normally with all commands visible
- **On other domains**: Add-in appears to not exist (no ribbon buttons, no commands)
- **No error messages**: Users won't know why the add-in isn't working

## Testing Domain Validation

### Silent Validation Only
The domain validation system is completely invisible to users:
- **No test buttons or commands visible**
- **No user interface for validation testing**
- **Validation happens automatically in the background**
- **Only administrators can check logs for validation status**

### Manual Testing
You can also test domain validation programmatically:
```csharp
var result = DomainValidator.ValidateDomain();
if (result.IsValid)
{
    // Domain validation successful
    Console.WriteLine($"Success: {result.ValidationMethod}");
}
else
{
    // Domain validation failed
    Console.WriteLine($"Failed: {result.FailureReason}");
}
```

## Troubleshooting

### Common Issues

#### 1. "Not connected to any domain"
- **Cause**: Computer is not joined to any domain
- **Solution**: Join computer to jse24 domain or authorized domain

#### 2. "Domain validation failed"
- **Cause**: Computer is on unauthorized domain
- **Solution**: Contact IT to move computer to jse24 domain

#### 3. "Validation error: [exception]"
- **Cause**: System error during validation
- **Solution**: Check system logs and contact administrator

### Network Issues
- Ensure computer can reach domain controllers
- Check DNS resolution for jse24 domain
- Verify network connectivity

### Active Directory Issues
- Ensure computer account is properly configured
- Check domain trust relationships
- Verify user permissions

## Security Features

### 1. Stealth Validation
- Multiple validation methods prevent circumvention
- Network-level validation ensures domain connectivity
- Active Directory integration for enterprise environments
- **Completely invisible to end users**

### 2. Silent Audit Logging
- All validation attempts are logged to security files
- Timestamp and user information recorded
- Security events tracked for compliance
- **No user notifications or error dialogs**

### 3. Invisible Protection
- Multiple validation layers
- Graceful degradation for network issues
- **Silent failure - users never know validation occurred**
- **Add-in simply doesn't appear on unauthorized domains**

## Configuration

### Domain Names
To modify required domain names, edit:
```csharp
// In Security/DomainValidator.cs
private const string REQUIRED_DOMAIN = "jse24";
private const string REQUIRED_DOMAIN_ALTERNATIVE = "jse";
```

### Logging
Security logs are stored in:
```
%APPDATA%\JSE_RevitAddin\Security\
```

### Validation Methods
You can enable/disable specific validation methods by modifying the `ValidateDomain()` method in `DomainValidator.cs`.

## Compliance

This domain validation system ensures:
- **Licensing Compliance**: Application only runs in authorized environments
- **Security**: Prevents unauthorized use outside company network
- **Audit Trail**: Complete logging of all validation attempts
- **Enterprise Integration**: Works with Active Directory environments

## Support

For issues with domain validation:
1. Check the security logs for detailed error information
2. Use the "Test Domain" command to diagnose issues
3. Contact your system administrator
4. Verify network and domain configuration

## Version History

- **v1.0**: Initial domain validation with jse24 requirement
- **v1.1**: Added comprehensive logging and error reporting
- **v1.2**: Enhanced validation methods and troubleshooting tools
