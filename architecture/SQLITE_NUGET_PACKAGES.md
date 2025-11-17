# SQLite NuGet Package Configuration

## 📦 Required NuGet Packages

### 1. **System.Data.SQLite.Core** (Version 1.0.118.0)
- **Purpose**: Bundled SQLite provider that includes both managed (`System.Data.SQLite.dll`) and native (`SQLite.Interop.dll`) components.
- **Package ID**: `System.Data.SQLite.Core`
- **Version**: `1.0.118.0`
- **Location in Project**: `JSE_RevitAddin_MEP_OPENINGS.csproj`

---

## 🔧 Implementation Details

### Package References in `.csproj`:
```xml
<PackageReference Include="System.Data.SQLite.Core" Version="1.0.118.0" />
```

### Initialization Code:
Connection and dependency verification are handled in `Data/SleeveDbContext.cs` and `Application.CopyNativeSqliteDllToExecutionDirectory()`. No manual `Batteries.Init()` call is required with `System.Data.SQLite.Core`.

---

## ✅ Verification Checklist

- [x] `System.Data.SQLite.Core` version 1.0.118.0 added to project
- [x] `SleeveDbContext` uses `System.Data.SQLite` APIs (`SQLiteConnection`, `SQLiteTransaction`, etc.)
- [x] `Application.CopyNativeSqliteDllToExecutionDirectory` mirrors `System.Data.SQLite.dll` and `x64\SQLite.Interop.dll`
- [x] Packages restored (run `dotnet restore` or restore via Visual Studio)

---

## 🚨 Common Issues and Solutions

### Issue: "Unable to load DLL 'SQLite.Interop.dll'"

**Cause**: Native SQLite runtime not copied to the Revit execution directory

**Solution**:
1. Ensure `System.Data.SQLite.Core` package is restored
2. Verify `x64\SQLite.Interop.dll` exists in build output and Revit Addins folder
3. Restore NuGet packages: `dotnet restore` or right-click solution → "Restore NuGet Packages"
4. Rebuild the project

### Issue: Database file not created

**Cause**: SQLite initialization failing silently

**Solution**:
1. Check logs for SQLite initialization errors
2. Verify `System.Data.SQLite.dll` and `x64\SQLite.Interop.dll` were copied to the execution directory
3. Check file system permissions on Filters directory
4. Verify project name sanitization (no invalid characters)

---

## 📝 Package Installation Commands

### Via Package Manager Console:
```powershell
Install-Package System.Data.SQLite.Core -Version 1.0.118.0
```

### Via .NET CLI:
```bash
dotnet add package System.Data.SQLite.Core --version 1.0.118.0
```

### Via Visual Studio:
1. Right-click project → "Manage NuGet Packages"
2. Search for "System.Data.SQLite.Core" → Install version 1.0.118.0

---

## 🔗 References

- **System.Data.SQLite Documentation**: https://system.data.sqlite.org/index.html/doc/trunk/www/index.wiki
- **Package Page**: https://www.nuget.org/packages/System.Data.SQLite.Core/

---

## ✅ Current Status

**Status**: ✅ **CONFIGURED**  
**Last Updated**: 2025-11-12  
**Packages**: `System.Data.SQLite.Core` added and initialized correctly

