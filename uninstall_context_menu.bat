@echo off
setlocal

rem Delete only the three registry trees owned by this tool.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $paths=@('HKCU:\Software\Classes\SystemFileAssociations\.xlsx\shell\SowMultiBranchSVNSubmit','HKCU:\Software\Classes\Directory\shell\SowMultiBranchSVNSubmit','HKCU:\Software\Classes\Directory\Background\shell\SowMultiBranchSVNSubmit'); foreach($path in $paths){ if(Test-Path -LiteralPath $path){ Remove-Item -LiteralPath $path -Recurse -Force } }"
if errorlevel 1 exit /b 1

if exist "%SystemRoot%\System32\ie4uinit.exe" "%SystemRoot%\System32\ie4uinit.exe" -show >nul 2>nul
echo Removed the three Sow multi-branch SVN context menus.
if /i not "%~1"=="/quiet" pause
exit /b 0
