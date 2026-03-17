<#
.SYNOPSIS
    Final Clean Build Script for GNU Emacs (Master Branch).
    Fixes: "Local changes would be overwritten" by forcing a git reset.
#>

param (
    [string]$InstallPath = "C:\msys64",
    [string]$BuildDir = "$env:USERPROFILE\emacs-source",
    [string]$OutputDir = "$env:USERPROFILE\emacs-build"
)

# Check Admin Privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "This script requires Administrator privileges."
    Break
}

$ErrorActionPreference = "Stop"

# --- Helper Function ---
function Invoke-MsysBash {
    param ([string]$Command)
    $BashPath = "$InstallPath\usr\bin\bash.exe"
    if (-not (Test-Path $BashPath)) { Throw "MSYS2 bash not found at $BashPath" }
    & $BashPath -l -c "$Command"
    if ($LASTEXITCODE -ne 0) { Throw "Command failed with exit code $LASTEXITCODE.`nCommand: $Command" }
}

Write-Host "=== Starting Emacs Build (Clean Master) ===" -ForegroundColor Cyan

# 1. Install/Update MSYS2
if (-not (Test-Path "$InstallPath\usr\bin\bash.exe")) {
    Write-Host "Installing MSYS2..." -ForegroundColor Yellow
    winget install --id "MSYS2.MSYS2" -e --source winget --accept-source-agreements --accept-package-agreements
    Write-Host "Waiting for MSYS2 installation to complete..." -ForegroundColor Yellow
    $WaitSeconds = 0
    while (-not (Test-Path "$InstallPath\usr\bin\bash.exe") -and $WaitSeconds -lt 120) {
        Start-Sleep -Seconds 5
        $WaitSeconds += 5
    }
    if (-not (Test-Path "$InstallPath\usr\bin\bash.exe")) {
        Throw "MSYS2 installation did not complete within 120 seconds. Check $InstallPath and re-run."
    }
}
Write-Host "Updating MSYS2..." -ForegroundColor Yellow
Invoke-MsysBash "pacman -Syu --noconfirm"

# 2. Install Dependencies
$Deps = @("base-devel", "autoconf", "automake", "texinfo", "mingw-w64-x86_64-toolchain", 
          "mingw-w64-x86_64-xpm-nox", "mingw-w64-x86_64-libtiff", "mingw-w64-x86_64-giflib", 
          "mingw-w64-x86_64-libpng", "mingw-w64-x86_64-libjpeg-turbo", "mingw-w64-x86_64-librsvg", 
          "mingw-w64-x86_64-libxml2", "mingw-w64-x86_64-gnutls", "mingw-w64-x86_64-zlib", 
          "mingw-w64-x86_64-harfbuzz", "mingw-w64-x86_64-jansson", "mingw-w64-x86_64-libgccjit", 
          "mingw-w64-x86_64-sqlite3", "git")

Write-Host "Installing dependencies..." -ForegroundColor Yellow
Invoke-MsysBash "pacman -S --needed --noconfirm $($Deps -join ' ')"

# 3. Clone and HARD RESET
# Convert Windows path to MSYS2/bash path (handles any drive letter, e.g. C: -> /c, D: -> /d)
$MsysBuildDir  = $BuildDir  -replace "\\", "/" -replace "^([A-Za-z]):", { "/" + $_.Groups[1].Value.ToLower() }
$MsysOutputDir = $OutputDir -replace "\\", "/" -replace "^([A-Za-z]):", { "/" + $_.Groups[1].Value.ToLower() }

if (-not (Test-Path $BuildDir)) {
    Write-Host "Cloning Emacs (Master)..." -ForegroundColor Yellow
    Invoke-MsysBash "git clone --depth 1 https://git.savannah.gnu.org/git/emacs.git $MsysBuildDir"
} else {
    Write-Host "Resetting Repository (Force Clean)..." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "[!] WARNING: This will run 'git clean -fdx' which permanently deletes ALL untracked" -ForegroundColor Yellow
    Write-Host "    and ignored files in $BuildDir (patches, local configs, build artifacts)." -ForegroundColor Yellow
    $ConfirmClean = Read-Host "    Proceed with hard reset and clean? (y/n)"
    if ($ConfirmClean -ne 'y') {
        Write-Error "Aborted by user. Remove or back up files in $BuildDir and re-run."
    }
    Invoke-MsysBash "cd $MsysBuildDir && git reset --hard && git clean -fdx && git fetch --depth 1 origin master && git checkout master && git pull"
}

# 4. Configure
Write-Host "Generating configure script..." -ForegroundColor Yellow
Invoke-MsysBash "cd $MsysBuildDir && ./autogen.sh"

Write-Host "Configuring..." -ForegroundColor Yellow
# CONFIG_SITE prevents the 'header shadowing' errors by using pre-defined Windows settings
$ConfigureCmd = "./configure --with-native-compilation --with-json --with-modules --with-harfbuzz --without-dbus --without-compress-install"
Invoke-MsysBash "export MSYSTEM=MINGW64 && export PATH=/mingw64/bin:`$PATH && export CONFIG_SITE=$MsysBuildDir/nt/mingw-cfg.site && cd $MsysBuildDir && $ConfigureCmd"

# 5. Build
$Cores = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors
Write-Host "Compiling (Cores: $Cores)..." -ForegroundColor Yellow
Invoke-MsysBash "export MSYSTEM=MINGW64 && export PATH=/mingw64/bin:`$PATH && cd $MsysBuildDir && make -j$Cores"

# 6. Install
Write-Host "Installing..." -ForegroundColor Yellow
Invoke-MsysBash "export MSYSTEM=MINGW64 && export PATH=/mingw64/bin:`$PATH && cd $MsysBuildDir && make install prefix=$MsysOutputDir"

# 7. Copy DLLs (only those actually needed by the Emacs binaries)
Write-Host "Copying dependencies..." -ForegroundColor Yellow
$BinDir = "$OutputDir\bin"
if (-not (Test-Path $BinDir)) { New-Item -ItemType Directory -Force -Path $BinDir | Out-Null }
$MsysBinDir = "$MsysOutputDir/bin"
# Use ldd within MSYS2 to discover the transitive DLL dependencies of runemacs.exe,
# then copy only those DLLs from mingw64/bin — avoids shipping hundreds of toolchain DLLs.
$LddCommand = @"
find $MsysBinDir -name '*.exe' -o -name '*.dll' | xargs ldd 2>/dev/null \
  | grep '/mingw64/bin' \
  | awk '{print `$3}' \
  | sort -u \
  | xargs -I{} cp -n {} $MsysBinDir/
"@
Invoke-MsysBash "export MSYSTEM=MINGW64 && export PATH=/mingw64/bin:`$PATH && $LddCommand"

Write-Host "--- Success! ---" -ForegroundColor Green
Write-Host "Run Emacs: $OutputDir\bin\runemacs.exe" -ForegroundColor Green