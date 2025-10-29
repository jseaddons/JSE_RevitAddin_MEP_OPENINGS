# Rollback Mechanism Guide

## 🔄 **Rollback Mechanisms Implemented**

### **1. Feature Flags Rollback** ⭐⭐⭐ **EASIEST**

**Location**: `Services/OptimizationFlags.cs`

**How to Rollback:**
```csharp
// Disable all optimizations instantly (safest rollback)
OptimizationFlags.ResetToSafeDefaults();
OptimizationFlags.UseFamilySymbolCache = false;      // Disable symbol caching
OptimizationFlags.UseXmlValidation = false;             // Disable validation
OptimizationFlags.UseOptimizedXmlSaves = false;        // Use old save method
OptimizationFlags.UseIncrementalCache = false;         // Use full cache rebuilds
```

**When to Use:**
- If any optimization causes issues
- To quickly revert to old behavior
- For troubleshooting

**Impact:** Instant rollback - changes take effect on next operation

---

### **2. XML Backup/Restore Rollback** ⭐⭐⭐ **AUTOMATIC**

**Location**: `Services/UniversalSleevePlacerService.cs`

**Automatic Process:**
1. **Before Save**: Creates `.backup` file (e.g., `ventilation.xml.backup`)
2. **After Save**: Validates save was successful (checks 95%+ zones updated)
3. **On Failure**: Automatically restores from `.backup` file

**Manual Restore (if needed):**
```csharp
// Programmatically restore any XML file
var service = new UniversalSleevePlacerService(...);
service.RestoreXmlFromBackup(xmlFilePath);
```

**Or Manually:**
1. Navigate to: `%AppData%\JSE_MEP_Openings\Projects\Default\Filters\`
2. Find your XML file (e.g., `ventilation_ducts.xml`)
3. Find backup file (e.g., `ventilation_ducts.xml.backup`)
4. Copy backup over original file

**When Backup is Created:**
- Every time `SaveUpdatedXmlFiles()` runs
- Only if `OptimizationFlags.UseXmlValidation = true`

**Backup Files:**
- Location: Same directory as XML files
- Format: `{OriginalFileName}.backup`
- Example: `ventilation_ducts.xml.backup`

---

### **3. Validation Failure Auto-Restore** ⭐⭐⭐ **AUTOMATIC**

**What Happens:**
1. XML save completes
2. Validation runs automatically
3. If validation fails (<95% success rate):
   - Backup is restored immediately
   - Error is thrown with clear message
   - Logs are written for debugging

**Example Error Message:**
```
[XML-VERIFY] Validation failed: Expected 100 updates, got 85 (85.0%). Threshold: 95%
[XML-RESTORE] ✓ Restored ventilation_ducts.xml from backup
InvalidOperationException: XML save verification failed for ventilation_ducts.xml - restored from backup
```

---

## 📋 **Quick Rollback Commands**

### **Emergency Rollback (All Optimizations):**
```csharp
// Disable everything - revert to original behavior
OptimizationFlags.UseFamilySymbolCache = false;
OptimizationFlags.UseXmlValidation = false;
OptimizationFlags.UseOptimizedXmlSaves = false;
OptimizationFlags.UseIncrementalCache = false;
```

### **Selective Rollback (Specific Feature):**
```csharp
// Disable only family symbol caching
OptimizationFlags.UseFamilySymbolCache = false;

// Disable only XML validation
OptimizationFlags.UseXmlValidation = false;
```

---

## 🔍 **Verification**

**Check Current Flags:**
```csharp
DebugLogger.Info($"Family Symbol Cache: {OptimizationFlags.UseFamilySymbolCache}");
DebugLogger.Info($"XML Validation: {OptimizationFlags.UseXmlValidation}");
DebugLogger.Info($"Optimized XML Saves: {OptimizationFlags.UseOptimizedXmlSaves}");
```

**Check Backup Files:**
```csharp
var filtersDirectory = Path.Combine(
    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
    "JSE_MEP_Openings", "Projects", "Default", "Filters");
var backupFiles = Directory.GetFiles(filtersDirectory, "*.backup");
DebugLogger.Info($"Found {backupFiles.Length} backup files");
```

---

## ✅ **Protection Layers**

1. **Layer 1**: Feature Flags (instant disable)
2. **Layer 2**: Automatic backup before save
3. **Layer 3**: Validation after save
4. **Layer 4**: Automatic restore on validation failure
5. **Layer 5**: Error logging for manual recovery

---

## 🚨 **Emergency Procedures**

### **If XML File Gets Corrupted:**
1. Check for `.backup` file in filters directory
2. Copy `.backup` file over corrupted `.xml` file
3. Restart operation

### **If Optimization Causes Crashes:**
1. Disable all feature flags:
   ```csharp
   OptimizationFlags.ResetToSafeDefaults();
   // Then set all to false
   OptimizationFlags.UseFamilySymbolCache = false;
   OptimizationFlags.UseXmlValidation = false;
   // etc.
   ```
2. Restart application
3. Report issue with flags disabled

---

## 📝 **Log Files**

All rollback operations are logged:
- `xml_save_errors.log` - Save failures
- `xml_restore.log` - Restore operations
- `memory_manager.log` - Memory/performance issues

---

## 🎯 **Summary**

✅ **Feature Flags**: Allow instant rollback by disabling features  
✅ **Auto Backup**: Every XML save creates backup  
✅ **Auto Validation**: Every save is validated  
✅ **Auto Restore**: Failed validation triggers automatic restore  
✅ **Manual Restore**: Backup files can be manually restored  
✅ **Error Logging**: All operations logged for debugging  

**Result**: Multiple layers of rollback protection ensure data safety and easy recovery from any issues.

