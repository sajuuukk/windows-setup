param (
    [string]$InstallPath = "C:\msys64",
    [string]$OutputDir   = "$env:USERPROFILE\emacs-build"
)

$WrapperPath = "$OutputDir\emacs-launcher.cmd"
$MsysBin     = "$InstallPath\mingw64\bin"
$EmacsBin    = "$OutputDir\bin\runemacs.exe"

if (-not (Test-Path $EmacsBin)) {
    Write-Error "runemacs.exe not found at '$EmacsBin'. Run build_emacs.ps1 first (with matching -OutputDir and -InstallPath)."
    Exit 1
}

if (-not (Test-Path $MsysBin)) {
    Write-Error "MinGW64 bin directory not found at '$MsysBin'. Ensure MSYS2 is installed at '$InstallPath'."
    Exit 1
}

$Content = @"
@echo off
REM Add MinGW64 compiler tools to PATH for Native Compilation
set "PATH=$MsysBin;%PATH%"
start "" "$EmacsBin" %*
"@

Set-Content -Path $WrapperPath -Value $Content
Write-Host "Launcher created at: $WrapperPath" -ForegroundColor Green