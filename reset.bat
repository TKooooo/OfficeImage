@echo off
setlocal
chcp 65001 >nul 2>&1
title Image Zoom Preview - Reset Office block

rem ===========================================================================
rem  Run this when Word/Excel says "this add-in caused a problem, disable it?".
rem  It clears the Office resiliency records that silently skip our add-in
rem  and resets LoadBehavior back to 3.
rem  HKCU only - no administrator rights needed.
rem  Close all Word/Excel windows BEFORE running this.
rem ===========================================================================

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0reset.ps1"

echo.
pause
