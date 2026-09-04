@echo off
title Philippine Star Front Page Compiler
cls
echo ============================================================
echo         Philippine Star Front Page Compiler
echo ============================================================
echo.
echo Select Drive to scan as source:
echo   [1] D: (Recommended - e.g. D:\copied folders)
echo   [2] C: (e.g. Desktop)
echo   [3] Custom Drive/Path
echo.
set /p drivec="Enter option [1-3] (Default: 1): "

if "%drivec%"=="" set drivec=1
if "%drivec%"=="1" set sdrive=D
if "%drivec%"=="2" set sdrive=C
if "%drivec%"=="3" goto custom_path

echo.
set /p sfolder="Enter source folder name on %sdrive%:\ (Default: copied folders): "
if "%sfolder%"=="" set sfolder=copied folders

echo.
echo Select month to process:
echo   [1] ALL months found in %sdrive%:\%sfolder% (Auto-detect)
echo   [2] Specific Month (e.g. 05 MAY, 06 JUN, 5)
echo.
set /p monthc="Enter option [1-2] (Default: 1): "

if "%monthc%"=="" set monthc=1
if "%monthc%"=="1" goto run_all
if "%monthc%"=="2" goto run_month

:custom_path
echo.
set /p fullsource="Enter full source path (e.g. D:\copied folders\05 MAY): "
echo.
echo Starting compilation...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Compile-Frontpages.ps1" -SourceDir "%fullsource%" -TargetDir "%USERPROFILE%\Desktop\FRONT_PAGE_2015"
goto finish

:run_all
echo.
echo Running compilation for ALL months in %sdrive%:\%sfolder%...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Compile-Frontpages.ps1" -Drive "%sdrive%" -FolderName "%sfolder%"
goto finish

:run_month
echo.
set /p mname="Enter Month (e.g. 05 MAY, MAY, 5): "
echo.
echo Running compilation for %mname% from %sdrive%:\%sfolder%...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Compile-Frontpages.ps1" -Drive "%sdrive%" -FolderName "%sfolder%" -Month "%mname%"
goto finish

:finish
echo.
echo ============================================================
echo Process complete.
pause
