# 📘 Complete Guide: MEP Mark, Remark Selected, and Parameter Transfer

## 📋 Table of Contents

1. [Overview](#overview)
2. [Prerequisites](#prerequisites)
3. [MEP Mark Requirements](#mep-mark-requirements)
4. [Remark Selected Requirements](#remark-selected-requirements)
5. [Parameter Transfer Requirements](#parameter-transfer-requirements)
6. [Database Requirements](#database-requirements)
7. [Parameter Naming Conventions](#parameter-naming-conventions)
8. [Document Structure Requirements](#document-structure-requirements)
9. [Configuration Requirements](#configuration-requirements)
10. [Troubleshooting](#troubleshooting)
11. [Best Practices](#best-practices)

---

## 🎯 Overview

This document details all necessary aspects for three critical features to work properly:

1. **MEP Mark** - Automatic marking of sleeves with discipline-specific prefixes
2. **Remark Selected** - Re-applying marks to selected categories with new prefixes
3. **Parameter Transfer** - Transferring MEP element parameters to sleeves

---

## ✅ Prerequisites

### 1. **Universal Family Files**

All 4 universal family files must be loaded in the project:

- `RectangularOpeningOnWall.rfa`
- `CircularOpeningOnWall.rfa`
- `RectangularOpeningOnSlab.rfa`
- `CircularOpeningOnSlab.rfa`

**Location:** `Resources/` folder

**Verification:**
- Open Revit → Insert → Load Family
- Check that all 4 families are available in the project

### 2. **Shared Parameter File**

The shared parameter file must exist and contain required parameters:

**Location:** `Resources/Opening family shared parameter.txt`

**Required Parameters:**
- `MEP Mark` (or `MEPMARK`) - Text parameter for marking
- `MEP_ElementId` - Integer parameter linking to MEP element
- `MEP System Type` - Text parameter for system type
- `MEP Size` - Text parameter for MEP element size
- `Sleeve Instance ID` - Integer parameter for individual sleeve tracking
- `Cluster Sleeve Instance ID` - Integer parameter for cluster sleeve tracking

**Verification:**
- Check that shared parameter file exists
- Verify all parameters are added to universal families

### 3. **Database Setup**

SQLite database must be initialized with required tables:

**Required Tables:**
- `ClashZones` - Contains clash zone data with MEP element information
- `ClusterSleeves` - Contains cluster sleeve data
- `SleeveSnapshots` - Contains parameter snapshots for sleeves

**Verification:**
- Run Refresh command to initialize database
- Check `database_operations.log` for table creation messages

### 4. **Revit Document Structure**

**Active Document:**
- Sleeves must be placed in the **active document** (not linked files)
- Active document must be modifiable (not read-only)

**Linked Documents:**
- MEP elements can be in **linked documents** (supported)
- Linked documents must be loaded and accessible

**Verification:**
- Ensure active document is not read-only
- Verify linked documents are loaded (Manage Links)

---

## 🏷️ MEP Mark Requirements

### 1. **Parameter Setup**

**Parameter Name:** `MEP Mark` or `MEPMARK` (case-insensitive)

**Parameter Type:** Shared Parameter (Text/Instance)

**Required On:**
- All 4 universal family files
- Must be instance parameter (not type parameter)
- Must be writable (not read-only)

### 2. **Mark Format**

#### **Individual Sleeve Format:**
```
{UserPrefix}-{DisciplinePrefix}-{Number:000}
```

**Components:**
- `UserPrefix`: Optional project prefix from UI (e.g., "JSE", "PROJ")
- `DisciplinePrefix`: Category-specific prefix (DCT, DMP, PLU, ELE)
- `Number`: Sequential number with zero-padding (001, 002, 003...)

**Examples:**
- `JSE-DCT-001` (User prefix + Duct individual #1)
- `DCT-001` (No user prefix, Duct individual #1)
- `JSE-PLU-015` (User prefix + Pipe individual #15)

#### **Cluster Sleeve Format:**
```
{UserPrefix}-{DisciplinePrefix}_C_-{Number:000}
```

**Components:**
- Same as individual, but with `_C_` cluster indicator
- `_C_` is literal text (not a variable)

**Examples:**
- `JSE-DCT_C_-001` (User prefix + Duct cluster #1)
- `DCT_C_-001` (No user prefix, Duct cluster #1)

### 3. **Discipline Prefixes**

| Category | Discipline Prefix | Example Individual | Example Cluster |
|----------|------------------|-------------------|-----------------|
| Ducts | `DCT` | `DCT-001` | `DCT_C_-001` |
| Duct Accessories (Dampers) | `DMP` | `DMP-001` | `DMP_C_-001` |
| Pipes | `PLU` | `PLU-001` | `PLU_C_-001` |
| Cable Trays | `ELE` | `ELE-001` | `ELE_C_-001` |

### 4. **Numbering Strategy**

**Per Category:**
- Each category has separate counters
- Individual and cluster counters are separate
- Counters reset each run (start at 001)

**Example Sequence:**
```
Ducts (22 individual, 3 clusters):
  Individual: DCT-001, DCT-002, ..., DCT-022
  Clusters: DCT_C_-001, DCT_C_-002, DCT_C_-003

Pipes (15 individual, 2 clusters):
  Individual: PLU-001, PLU-002, ..., PLU-015
  Clusters: PLU_C_-001, PLU_C_-002
```

### 5. **UI Requirements**

**Project Prefix Input:**
- TextBox for user prefix (optional)
- Stored in `MarkPrefixSettings.ProjectPrefix`
- Case-insensitive (converted to uppercase)

**Number Format:**
- Options: `00`, `000`, `0000`
- Default: `000` (001, 002, 003...)

**Category Selection:**
- Must select at least one category
- Categories: Ducts, Pipes, Cable Trays, Dampers

### 6. **Database Requirements for MEP Mark**

**ClashZone Table:**
- Must contain clash zone data with `MepElementCategory`
- Required for determining discipline prefix

**ClusterSleeves Table:**
- Must contain cluster sleeve data
- Required for cluster mark generation

**Verification:**
- Run Refresh before applying marks
- Check `Refresh_debug.log` for clash zone detection

---

## 🔄 Remark Selected Requirements

### 1. **Purpose**

Re-apply marks to sleeves that already have marks, updating the prefix while preserving the number.

### 2. **Remark Flags**

**Per-Category Remark Flags:**
- `RemarkProject` - Re-mark all sleeves
- `RemarkDuct` - Re-mark Duct sleeves only
- `RemarkPipe` - Re-mark Pipe sleeves only
- `RemarkCableTray` - Re-mark Cable Tray sleeves only
- `RemarkDamper` - Re-mark Damper sleeves only

**Behavior:**
- If `RemarkDuct = true`, only Duct sleeves are re-marked
- Other categories are skipped (preserve existing marks)
- If `RemarkProject = true`, all categories are re-marked

### 3. **Number Preservation**

**When Remarking:**
- Extract existing number from current mark
- Preserve the number
- Update only the prefix

**Example:**
```
Existing mark: "OLD-DCT-005"
New prefix: "JSE"
Result: "JSE-DCT-005" (number preserved)
```

### 4. **Requirements**

**Existing Mark Required:**
- Sleeve must have existing `MEP Mark` value
- If no existing mark, sleeve is skipped

**Mark Format Validation:**
- Must match expected format to extract number
- Invalid formats are skipped

**Category Matching:**
- Sleeve category must match selected category
- Determined from `MEP_Category` parameter or clash zone data

---

## 🔄 Parameter Transfer Requirements

### 1. **Parameter Mappings**

**Source Parameters (from MEP elements):**
- `Size` → Maps to `MEP Size` on sleeve
- `System Type` → Maps to `MEP System Type` on sleeve
- `Service Type` (Cable Trays) → Maps to `MEP System Type` on sleeve
- `MEP System Abbreviation` → Maps to `MEP System Abbreviation` on sleeve

**Target Parameters (on sleeves):**
- `MEP Size` - Text parameter (read from Revit MEP element)
- `MEP System Type` - Text parameter (from snapshot or Revit)
- `MEP System Abbreviation` - Text parameter (from snapshot)
- `MEP_ElementId` - Integer parameter (critical, must exist)

### 2. **Parameter Naming Conventions**

**Critical:** Parameter names use **spaces**, not underscores (except for specific parameters):

| Parameter Name | Format | Notes |
|---------------|--------|-------|
| `MEP Size` | Space | NOT `MEP_Size` |
| `MEP System Type` | Space | NOT `MEP_System_Type` |
| `MEP System Abbreviation` | Space | NOT `MEP_System_Abbreviation` |
| `MEP_ElementId` | Underscore | Exception - uses underscore |
| `MEP_UniqueId` | Underscore | Exception - uses underscore |
| `MEP_Count` | Underscore | Exception - uses underscore |
| `MEP_Category` | Underscore | Exception - uses underscore |

### 3. **Transfer Types**

**ReferenceToOpening:**
- Transfers from MEP elements (ducts, pipes, cable trays)
- Uses `MEP_ElementId` to find MEP element
- MEP elements can be in linked documents

**HostToOpening:**
- Transfers from host elements (walls, floors, ceilings)
- Uses host element parameters

**LevelToOpening:**
- Transfers from level elements
- Uses level name

### 4. **MEP Size Transfer (Special Case)**

**For Individual Sleeves:**
- **ALWAYS** read from Revit MEP element (not database snapshot)
- Uses `MEP_ElementId` to find MEP element
- Reads `Size` parameter from MEP element
- Fallback to snapshot only if `MEP_ElementId` is invalid

**For Cluster Sleeves:**
- **ALWAYS** read from database snapshot (aggregated data)
- Does NOT read from individual MEP elements
- Snapshot contains aggregated MEP Size

**Verification:**
- Check `transfer_debug.log` for MEP Size source
- Look for: `Read 'MEP Size'='...' from Revit MEP element` (individual)
- Look for: `Read 'MEP Size'='...' from SNAPSHOT` (cluster)

### 5. **System Type Transfer (Cable Trays)**

**Special Mapping:**
- For Cable Trays, `Service Type` maps to `System Type`
- Parameter variation matching: `Service Type` → `System Type`
- Checked in snapshot first, then Revit element

**Verification:**
- Check `transfer_debug.log` for parameter variation matching
- Look for: `Cable Trays detected - mapping 'System Type' -> 'Service Type'`

### 6. **Snapshot Requirements**

**SleeveSnapshots Table:**
- Must contain parameter snapshots for sleeves
- Required for cluster sleeve parameter transfer
- Contains aggregated data for clusters

**Snapshot Parameters:**
- `Size` or `MEP Size` - For cluster sleeves
- `System Type` or `Service Type` - For system type transfer
- `MepParameters` - JSON containing all MEP parameters

**Verification:**
- Check `database_operations.log` for snapshot saves
- Verify snapshots are updated (not just inserted)

### 7. **Element Retrieval**

**Sleeves:**
- Always in active document
- Retrieved using `doc.GetElement(openingId)`

**MEP Elements:**
- Can be in linked documents
- Retrieved using `ElementRetrievalService.GetElementFromDocumentOrLinked()`
- Supports both host and linked documents

### 8. **Document Validation**

**Critical Checks:**
- Document is not null
- Document is not read-only
- Element belongs to correct document
- Parameter element matches opening element (by ID)

**Protection:**
- 17 protection checks implemented in `ParameterTransferService`
- Validates at each critical step
- Prevents document mismatch errors

---

## 💾 Database Requirements

### 1. **Required Tables**

**ClashZones:**
- `Id` (GUID) - Primary key
- `MepElementIdValue` - Integer ID of MEP element
- `MepElementCategory` - Category name (Ducts, Pipes, etc.)
- `MepWidth` - MEP element width (feet)
- `MepHeight` - MEP element height (feet)
- `SleeveInstanceId` - Individual sleeve instance ID
- `ClusterInstanceId` - Cluster sleeve instance ID

**ClusterSleeves:**
- `ClusterInstanceId` - Primary key
- `ClusterGuid` - Deterministic GUID for upserting
- `Category` - Category name
- `ClashZoneIds` - JSON array of clash zone GUIDs

**SleeveSnapshots:**
- `SleeveInstanceId` - Individual sleeve instance ID
- `ClusterInstanceId` - Cluster sleeve instance ID
- `ClashZoneGuid` - GUID for lookup
- `MepParameters` - JSON containing parameter values
- `UpdatedAt` - Timestamp (should update, not just insert)

### 2. **Database Initialization**

**First Run:**
- Database is created automatically
- Tables are created with schema
- R-tree index is created for spatial queries

**Verification:**
- Check `database_operations.log` for table creation
- Verify R-tree index is created

### 3. **Data Consistency**

**Deterministic GUIDs:**
- Cluster sleeves use deterministic GUIDs for upserting
- Prevents duplicate entries
- Ensures updates instead of inserts

**Snapshot Updates:**
- Snapshots should be updated (not inserted) on refresh
- `UpdatedAt` timestamp should change
- Check `database_operations.log` for UPDATE vs INSERT

---

## 📝 Parameter Naming Conventions

### **Critical Rules:**

1. **Spaces vs Underscores:**
   - Most parameters use **spaces**: `MEP Size`, `MEP System Type`
   - Only specific parameters use **underscores**: `MEP_ElementId`, `MEP_UniqueId`

2. **Case Sensitivity:**
   - Parameter names are **case-insensitive** in lookups
   - But use exact case in code for consistency

3. **Variation Matching:**
   - Code handles variations: `System Type` ↔ `System_Type`
   - `MEP System Type` ↔ `System Type`
   - `Service Type` → `System Type` (for Cable Trays)

### **Parameter List:**

| Parameter Name | Format | Type | Critical |
|---------------|--------|------|----------|
| `MEP Size` | Space | Text | No (warning if missing) |
| `MEP System Type` | Space | Text | Yes (error if missing) |
| `MEP System Abbreviation` | Space | Text | No |
| `MEP_ElementId` | Underscore | Integer | Yes (error if missing) |
| `MEP_UniqueId` | Underscore | Text | No |
| `MEP_Count` | Underscore | Integer | No |
| `MEP_Category` | Underscore | Text | No |
| `Sleeve Instance ID` | Space | Integer | No |
| `Cluster Sleeve Instance ID` | Space | Integer | No |
| `MEP Mark` | Space | Text | No |

---

## 📄 Document Structure Requirements

### 1. **Active Document (Sleeves)**

**Requirements:**
- Must be modifiable (not read-only)
- Must not be closed
- Sleeves must be placed in active document
- Universal families must be loaded

**Verification:**
- Check document is not read-only
- Verify families are loaded (Insert → Load Family)

### 2. **Linked Documents (MEP Elements)**

**Requirements:**
- Linked documents must be loaded
- MEP elements can be in any linked document
- `ElementRetrievalService` handles linked file retrieval

**Verification:**
- Manage Links → Check linked files are loaded
- Verify MEP elements are accessible

### 3. **Document Validation**

**Checks Performed:**
- Document is not null
- Document is not read-only
- Element belongs to correct document
- Parameter element matches opening element

---

## ⚙️ Configuration Requirements

### 1. **Optimization Flags**

**Location:** `Services/OptimizationFlags.cs`

**Relevant Flags:**
- `UseBatchParameterLookups` - Enable parameter caching
- `SkipAlreadyTransferredParameters` - Skip if parameter already matches
- `UseBatchedParameterWrites` - Batch parameter writes (performance)

**Recommendation:**
- Enable all flags for production
- Disable for debugging if needed

### 2. **Deployment Mode**

**Location:** `Services/DeploymentConfiguration.cs`

**DeploymentMode:**
- `true` - Disables diagnostic logging (production)
- `false` - Enables diagnostic logging (development)

**Recommendation:**
- Set to `false` during development
- Set to `true` for production deployment

### 3. **Parameter Transfer Configuration**

**Location:** XML configuration files (category-specific)

**Required Files:**
- `Resources/ParameterMappings/Ducts.xml`
- `Resources/ParameterMappings/Pipes.xml`
- `Resources/ParameterMappings/CableTrays.xml`
- `Resources/ParameterMappings/Dampers.xml`

**Structure:**
```xml
<ParameterMappings>
  <Mapping>
    <SourceParameter>Size</SourceParameter>
    <TargetParameter>MEP Size</TargetParameter>
    <TransferType>ReferenceToOpening</TransferType>
    <IsEnabled>true</IsEnabled>
  </Mapping>
</ParameterMappings>
```

---

## 🔧 Troubleshooting

### 1. **MEP Mark Not Applied**

**Symptoms:**
- Sleeves don't have `MEP Mark` values
- Marks are empty or null

**Checks:**
1. Verify `MEP Mark` parameter exists in universal families
2. Check shared parameter file exists
3. Verify families are loaded in project
4. Check `mepmark_debug.log` for errors

**Solutions:**
- Re-add `MEP Mark` parameter to families
- Reload families into project
- Check UI input (project prefix, number format)

### 2. **Remark Selected Not Working**

**Symptoms:**
- Existing marks not updated
- Only some categories updated

**Checks:**
1. Verify remark flag is set for category
2. Check existing mark format is valid
3. Verify category matching (MEP_Category parameter)
4. Check `mepmark_debug.log` for skip messages

**Solutions:**
- Enable remark flag for category
- Verify mark format matches expected pattern
- Check `MEP_Category` parameter on sleeves

### 3. **Parameter Transfer Failing**

**Symptoms:**
- "Transfer failed: 0 successful, X failed"
- Parameters not transferred to sleeves

**Checks:**
1. Check `transfer_debug.log` for detailed errors
2. Verify `MEP_ElementId` parameter exists on sleeves
3. Check MEP elements are accessible (linked files)
4. Verify parameter names match (spaces vs underscores)
5. Check snapshot data exists for cluster sleeves

**Common Issues:**

**Issue: Document Mismatch**
```
Error: Parameter 'MEP Size' belongs to different document
```
**Solution:** Fixed in code - element ID comparison instead of document reference

**Issue: MEP Size Not Transferred**
```
Error: MEP element not found or invalid
```
**Solution:** 
- Verify `MEP_ElementId` is set on sleeve
- Check MEP element exists in linked file
- Verify `ElementRetrievalService` is working

**Issue: Cluster Sleeves Not Getting Parameters**
```
Error: Parameter not found in snapshot
```
**Solution:**
- Run Refresh to create snapshots
- Verify snapshots are updated (not just inserted)
- Check `ClusterGuid` is populated in `ClusterSleeves` table

### 4. **Database Issues**

**Symptoms:**
- Snapshots not found
- Cluster sleeves not updating

**Checks:**
1. Verify database exists and is accessible
2. Check `database_operations.log` for errors
3. Verify tables exist (ClashZones, ClusterSleeves, SleeveSnapshots)
4. Check `ClusterGuid` is populated

**Solutions:**
- Delete database and run Refresh (fresh start)
- Verify deterministic GUID generation
- Check snapshot upsert logic

### 5. **Performance Issues**

**Symptoms:**
- Parameter transfer is slow
- Marks take long to apply

**Checks:**
1. Verify optimization flags are enabled
2. Check `parameter_batching_performance.log`
3. Verify parameter caching is working

**Solutions:**
- Enable `UseBatchParameterLookups`
- Enable `UseBatchedParameterWrites`
- Check for excessive logging (set DeploymentMode = true)

---

## ✅ Best Practices

### 1. **Workflow Order**

**Recommended Sequence:**
1. **Refresh** - Detect clash zones and create database
2. **Place Sleeves** - Place individual and cluster sleeves
3. **Apply MEP Mark** - Apply marks to sleeves
4. **Parameter Transfer** - Transfer MEP parameters to sleeves
5. **Remark Selected** (if needed) - Update marks with new prefix

### 2. **Database Management**

**Before Major Changes:**
- Backup database (copy `.db` file)
- Delete database for fresh start if needed
- Run Refresh to rebuild data

**After Refresh:**
- Verify clash zones are detected
- Check `Refresh_debug.log` for errors
- Verify snapshots are created/updated

### 3. **Parameter Naming**

**Consistency:**
- Always use exact parameter names (spaces vs underscores)
- Check parameter names in shared parameter file
- Verify families have correct parameters

### 4. **Testing**

**Before Production:**
- Test with small dataset first
- Verify marks are applied correctly
- Check parameter transfer works
- Test with linked files
- Test cluster sleeves separately

### 5. **Logging**

**Development:**
- Set `DeploymentMode = false` for detailed logs
- Check `transfer_debug.log` for parameter transfer
- Check `mepmark_debug.log` for mark application
- Check `database_operations.log` for database issues

**Production:**
- Set `DeploymentMode = true` to reduce logging overhead
- Monitor error logs only

### 6. **Error Handling**

**When Errors Occur:**
1. Check relevant log file (`transfer_debug.log`, `mepmark_debug.log`)
2. Verify prerequisites are met
3. Check database state
4. Verify parameter names match
5. Check document structure (active vs linked)

---

## 📚 Additional Resources

### **Log Files:**

| Log File | Purpose | Location |
|----------|---------|----------|
| `transfer_debug.log` | Parameter transfer details | `Logs/R2023/` |
| `mepmark_debug.log` | MEP mark application details | `Logs/R2023/` |
| `database_operations.log` | Database operations | `Logs/R2023/` |
| `parameter_service_debug.log` | Parameter service errors | `Logs/R2023/` |
| `refresh_mep_sizes.log` | MEP size tracking | `Logs/R2023/` |

### **Key Files:**

| File | Purpose |
|------|---------|
| `Services/ParameterTransferService.cs` | Parameter transfer logic |
| `Services/MarkParameterService.cs` | MEP mark application logic |
| `Services/OptimizationFlags.cs` | Feature flags |
| `Services/DeploymentConfiguration.cs` | Deployment settings |
| `Data/Repositories/ClashZoneRepository.cs` | Database operations |
| `Data/Repositories/SleeveSnapshotRepository.cs` | Snapshot operations |

---

## 🎯 Summary Checklist

### **Before Using MEP Mark:**
- [ ] Universal families loaded
- [ ] `MEP Mark` parameter exists in families
- [ ] Shared parameter file exists
- [ ] Database initialized (run Refresh)
- [ ] Clash zones detected

### **Before Using Remark Selected:**
- [ ] Sleeves have existing marks
- [ ] Remark flag enabled for category
- [ ] Project prefix configured
- [ ] Category selected

### **Before Using Parameter Transfer:**
- [ ] `MEP_ElementId` parameter exists on sleeves
- [ ] MEP elements accessible (linked files loaded)
- [ ] Parameter names match (spaces vs underscores)
- [ ] Snapshots exist for cluster sleeves
- [ ] Database tables exist

---

**Last Updated:** 2025-12-03  
**Version:** 1.0

