@echo off
setlocal

echo Removing TortoiseSVN diff/merge tools for .xlsx and .xlsm ...

reg delete "HKCU\Software\TortoiseSVN\DiffTools" /v .xlsx /f
reg delete "HKCU\Software\TortoiseSVN\MergeTools" /v .xlsx /f

reg delete "HKCU\Software\TortoiseSVN\DiffTools\XLSX" /f
reg delete "HKCU\Software\TortoiseSVN\MergeTools\XLSX" /f

reg delete "HKCU\Software\TortoiseSVN\DiffTools" /v .xlsm /f
reg delete "HKCU\Software\TortoiseSVN\MergeTools" /v .xlsm /f
reg delete "HKCU\Software\TortoiseSVN\DiffTools\XLSM" /f
reg delete "HKCU\Software\TortoiseSVN\MergeTools\XLSM" /f

if exist "%~dp0uninstall_context_menu.bat" call "%~dp0uninstall_context_menu.bat" /quiet

echo Done. TortoiseSVN integration and Explorer context menu are removed.
pause
