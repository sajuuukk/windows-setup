$WrapperPath = "$env:USERPROFILE\emacs-build\emacs-launcher.cmd"
$MsysBin = "C:\msys64\mingw64\bin"
$EmacsBin = "$env:USERPROFILE\emacs-build\bin\runemacs.exe"

$Content = @"
@echo off
REM Add MinGW64 compiler tools to PATH for Native Compilation
set "PATH=$MsysBin;%PATH%"
start "" "$EmacsBin" %*
"@

Set-Content -Path $WrapperPath -Value $Content
Write-Host "Launcher created at: $WrapperPath" -ForegroundColor Green