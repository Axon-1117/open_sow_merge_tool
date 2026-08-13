@echo off
setlocal

set "TOOL_PATH=%~dp0sow_merge_tool.exe"
if not exist "%TOOL_PATH%" set "TOOL_PATH=%~dp0dist\sow_merge_tool.exe"

if not exist "%TOOL_PATH%" (
  echo ERROR: sow_merge_tool.exe was not found.
  echo %TOOL_PATH%
  exit /b 1
)

rem Keep this batch file ASCII-only. cmd.exe can corrupt its instruction pointer
rem after reading multibyte UTF-8 text. PowerShell constructs the Chinese label
rem from Unicode code points and writes the three exact HKCU registry entries.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $tool=[IO.Path]::GetFullPath($env:TOOL_PATH); $quote=[char]34; $percent=[char]37; $label=[string]::Concat([char[]](0x591A,0x5206,0x652F,0x20,0x53,0x56,0x4E,0x20,0x63D0,0x4EA4)); $items=@(@{Path='HKCU:\Software\Classes\SystemFileAssociations\.xlsx\shell\SowMultiBranchSVNSubmit';Token='1';Multi='Player'},@{Path='HKCU:\Software\Classes\Directory\shell\SowMultiBranchSVNSubmit';Token='1';Multi=$null},@{Path='HKCU:\Software\Classes\Directory\Background\shell\SowMultiBranchSVNSubmit';Token='V';Multi=$null}); foreach($item in $items){ New-Item -Path $item.Path -Force | Out-Null; Set-Item -LiteralPath $item.Path -Value $label; New-ItemProperty -LiteralPath $item.Path -Name 'Icon' -PropertyType String -Value ($quote+$tool+$quote+',0') -Force | Out-Null; New-ItemProperty -LiteralPath $item.Path -Name 'Position' -PropertyType String -Value 'Top' -Force | Out-Null; if($item.Multi){ New-ItemProperty -LiteralPath $item.Path -Name 'MultiSelectModel' -PropertyType String -Value $item.Multi -Force | Out-Null }; $commandPath=$item.Path+'\command'; New-Item -Path $commandPath -Force | Out-Null; $command=$quote+$tool+$quote+' --branch-submit '+$quote+$percent+$item.Token+$quote; Set-Item -LiteralPath $commandPath -Value $command }"
if errorlevel 1 exit /b 1

if exist "%SystemRoot%\System32\ie4uinit.exe" "%SystemRoot%\System32\ie4uinit.exe" -show >nul 2>nul
echo Installed context menus for .xlsx files, folders, and folder backgrounds.
echo Tool: %TOOL_PATH%
if /i not "%~1"=="/quiet" pause
exit /b 0
