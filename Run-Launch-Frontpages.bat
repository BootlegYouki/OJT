@echo off
title Philippine Star Front Page Viewer
cls
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Launch-Frontpages.ps1"
echo.
echo ============================================================
echo Process finished.
pause
