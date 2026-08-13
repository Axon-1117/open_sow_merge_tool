@echo off
setlocal
chcp 65001 >nul

set "TOOL_PATH=%~dp0sow_merge_tool.exe"
if not exist "%TOOL_PATH%" set "TOOL_PATH=%~dp0dist\sow_merge_tool.exe"
set "MENU_KEY=HKCU\Software\Classes\SystemFileAssociations\.xlsx\shell\SowMultiBranchSVNSubmit"
set "COMMAND_KEY=%MENU_KEY%\command"

if not exist "%TOOL_PATH%" (
  echo ERROR: 找不到 sow_merge_tool.exe
  echo %TOOL_PATH%
  exit /b 1
)

reg add "%MENU_KEY%" /ve /t REG_SZ /d "多分支 SVN 提交" /f >nul
if errorlevel 1 exit /b 1
reg add "%MENU_KEY%" /v Icon /t REG_SZ /d "\"%TOOL_PATH%\",0" /f >nul
if errorlevel 1 exit /b 1
reg add "%MENU_KEY%" /v MultiSelectModel /t REG_SZ /d "Single" /f >nul
if errorlevel 1 exit /b 1
reg add "%MENU_KEY%" /v Position /t REG_SZ /d "Top" /f >nul
if errorlevel 1 exit /b 1
reg add "%COMMAND_KEY%" /ve /t REG_SZ /d "\"%TOOL_PATH%\" --branch-submit \"%%1\"" /f >nul
if errorlevel 1 exit /b 1

if exist "%SystemRoot%\System32\ie4uinit.exe" "%SystemRoot%\System32\ie4uinit.exe" -show >nul 2>nul
echo 已安装 .xlsx 右键菜单：多分支 SVN 提交
echo 工具：%TOOL_PATH%
if /i not "%~1"=="/quiet" pause
exit /b 0

