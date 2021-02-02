@echo off
cd /d "%~dp0"

if exist "%cd%\_HJT_src.zip" (
  del /f /a "%cd%\_HJT_src.zip" || (echo Unable to delete old src archive! & pause>NUL & exit /b)
)

copy /y Archive\MSCOMCTL.OCX MSCOMCTL.OCX
del MSCOMCTL.OCA

:: Pack
Tools\7zip\7za.exe a -mx1 -y -o"%cd%" -x!*.zip -x!tools/Align4byte/Align4byte.exe -x!tools/ChangeIcon/IC.exe -x!tools/RegTLib/RegTLib.exe -x!tools/TSAwarePatch/TSAwarePatch.exe -x!tools/VersionPatcher/VersionPatcher.exe -x!tools/RemoveSign/RemSign.exe _HJT_src.zip *.* Tools Ico
:: Test
Tools\7zip\7za.exe t "%cd%\_HJT_src.zip"
:: If there was errors
if %errorlevel% neq 0 (pause & exit /B)
