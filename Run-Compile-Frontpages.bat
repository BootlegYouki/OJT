@echo off
title Philippine Star Front Page Compiler
cls
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Compile-Frontpages.ps1"
echo.
echo ============================================================
echo Process finished.
pause
