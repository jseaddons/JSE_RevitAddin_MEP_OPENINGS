$solutionPath = "C:\Jse_Developments\JSE_MEPOPENING_23\JSE_RevitAddin_MEP_OPENINGS.csproj"
$msbuild = "msbuild"  # run from Developer Command Prompt or set full path to msbuild.exe
$configs = @("Release R23","Release R24","Release R25","Release R26")
foreach ($cfg in $configs) {
    Write-Host "Building $cfg"
    & $msbuild $solutionPath /p:Configuration="$cfg" /t:Rebuild /m
    if ($LASTEXITCODE -ne 0) { Write-Error "Build failed for $cfg"; exit $LASTEXITCODE }
}
Write-Host "All builds finished."