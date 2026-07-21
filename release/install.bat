@echo off
setlocal

set TOOL_PATH=%~dp0sow_merge_tool.exe

echo Installing TortoiseSVN diff/merge tools for .xlsx and .xlsm ...

for %%E in (.xlsx .xlsm) do (
  reg add "HKCU\Software\TortoiseSVN\DiffTools" /v %%E /t REG_SZ /d "\"%TOOL_PATH%\" --base \"%base\" --mine \"%mine\" --title \"%bname\"" /f
  reg add "HKCU\Software\TortoiseSVN\MergeTools" /v %%E /t REG_SZ /d "\"%TOOL_PATH%\" --base \"%base\" --mine \"%mine\" --theirs \"%theirs\" --merged \"%merged\" --title \"%bname\"" /f
)

for %%K in (XLSX XLSM) do (
  reg add "HKCU\Software\TortoiseSVN\DiffTools\%%K" /v command /t REG_SZ /d "%TOOL_PATH%" /f
  reg add "HKCU\Software\TortoiseSVN\DiffTools\%%K" /v args /t REG_SZ /d "--base %base --mine %mine --title %bname" /f
  reg add "HKCU\Software\TortoiseSVN\MergeTools\%%K" /v command /t REG_SZ /d "%TOOL_PATH%" /f
  reg add "HKCU\Software\TortoiseSVN\MergeTools\%%K" /v args /t REG_SZ /d "--base %base --mine %mine --theirs %theirs --merged %merged --title %bname" /f
)

echo Done. Please restart TortoiseSVN / Explorer if needed.
pause
