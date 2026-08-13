@echo off
setlocal
chcp 65001 >nul

set "MENU_KEY=HKCU\Software\Classes\SystemFileAssociations\.xlsx\shell\SowMultiBranchSVNSubmit"
reg delete "%MENU_KEY%" /f >nul 2>nul
if exist "%SystemRoot%\System32\ie4uinit.exe" "%SystemRoot%\System32\ie4uinit.exe" -show >nul 2>nul

echo 已卸载 .xlsx 右键菜单：多分支 SVN 提交
if /i not "%~1"=="/quiet" pause
exit /b 0

