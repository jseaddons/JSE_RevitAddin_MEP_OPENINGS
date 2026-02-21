# Dependency DLL Paths for ISS Installer

## Correct Paths (as per actual build output)

### R23 (.NET Framework 4.8)
**Base path:** `bin\Debug R23\Debug R23\`

| DLL | Path |
|-----|------|
| Main DLL | `bin\Debug R23\Debug R23\JSE_RevitAddin_MEP_OPENINGS.dll` |
| Addin | `bin\Debug R23\Debug R23\JSE_RevitAddin_MEP_OPENINGS.addin` |
| Parameter Service | `bin\Debug R23\Debug R23\JSE_Parameter_Service.dll` |
| CommunityToolkit.Mvvm | `bin\Debug R23\Debug R23\CommunityToolkit.Mvvm.dll` |
| Serilog | `bin\Debug R23\Debug R23\Serilog.dll` |
| Serilog.Sinks.Debug | `bin\Debug R23\Debug R23\Serilog.Sinks.Debug.dll` |
| System.Text.Json | `bin\Debug R23\Debug R23\System.Text.Json.dll` |
| System.Text.Encodings.Web | `bin\Debug R23\Debug R23\System.Text.Encodings.Web.dll` |
| System.Memory | `bin\Debug R23\Debug R23\System.Memory.dll` |
| System.Runtime.CompilerServices.Unsafe | `bin\Debug R23\Debug R23\System.Runtime.CompilerServices.Unsafe.dll` |
| System.Buffers | `bin\Debug R23\Debug R23\System.Buffers.dll` |
| System.Numerics.Vectors | `bin\Debug R23\Debug R23\System.Numerics.Vectors.dll` |
| System.Data.SQLite | `bin\Debug R23\Debug R23\System.Data.SQLite.dll` |
| SQLite.Interop (x64) | `bin\Debug R23\Debug R23\x64\SQLite.Interop.dll` |
| SQLite.Interop (x86) | `bin\Debug R23\Debug R23\x86\SQLite.Interop.dll` |

### R24 (.NET 6)
**Base path:** `bin\Debug R24\Debug R24\`

Same structure as R23 (dependencies in same folder as main DLL)

### R25 (.NET 8)
**Main DLL path:** `bin\Debug R25\Debug R25\`
**Dependencies path:** `bin\Debug R25\Debug R25\publish\`

| DLL | Path |
|-----|------|
| Main DLL | `bin\Debug R25\Debug R25\JSE_RevitAddin_MEP_OPENINGS.dll` |
| Addin | `bin\Debug R25\Debug R25\JSE_RevitAddin_MEP_OPENINGS.addin` |
| Parameter Service | `bin\Debug R25\Debug R25\JSE_Parameter_Service.dll` |
| CommunityToolkit.Mvvm | `bin\Debug R25\Debug R25\publish\CommunityToolkit.Mvvm.dll` |
| Serilog | `bin\Debug R25\Debug R25\publish\Serilog.dll` |
| Serilog.Sinks.Debug | `bin\Debug R25\Debug R25\publish\Serilog.Sinks.Debug.dll` |
| System.Text.Json | `bin\Debug R25\Debug R25\publish\System.Text.Json.dll` |
| Microsoft.Data.Sqlite | `bin\Debug R25\Debug R25\publish\Microsoft.Data.Sqlite.dll` |
| SQLitePCLRaw.batteries_v2 | `bin\Debug R25\Debug R25\publish\SQLitePCLRaw.batteries_v2.dll` |
| SQLitePCLRaw.core | `bin\Debug R25\Debug R25\publish\SQLitePCLRaw.core.dll` |
| SQLitePCLRaw.provider.e_sqlite3 | `bin\Debug R25\Debug R25\publish\SQLitePCLRaw.provider.e_sqlite3.dll` |
| e_sqlite3 | `bin\Debug R25\Debug R25\publish\e_sqlite3.dll` |

### R26 (.NET 8)
**Main DLL path:** `bin\Debug R26\Debug R26\`
**Dependencies path:** `bin\Debug R26\Debug R26\publish\`

Same structure as R25 (dependencies in `publish` subfolder)

## How to Verify

After building, run this PowerShell command to see what DLLs exist:

```powershell
# For R23
Get-ChildItem "bin\Debug R23\Debug R23\*.dll" | Select-Object Name

# For R25 (note the publish subfolder for dependencies)
Get-ChildItem "bin\Debug R25\Debug R25\*.dll" | Select-Object Name
Get-ChildItem "bin\Debug R25\Debug R25\publish\*.dll" | Select-Object Name
```

## Notes

1. **R23/R24** (.NET Framework/.NET 6): All DLLs are in the same folder as the main DLL
2. **R25/R26** (.NET 8): Main DLL is in root, but dependencies are in `publish\` subfolder
3. **JSE_Parameter_Service.dll**: You must manually copy this to each `Debug R##\Debug R##\` folder
4. All other DLLs are copied automatically by MSBuild from NuGet packages
