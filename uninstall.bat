@echo off
setlocal
chcp 65001 >nul 2>&1
title Image Zoom Preview - Uninstall

rem 需要管理员权限：HKLM 下的 CLSID 注册

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo   Requesting administrator rights ...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"

echo.
pause
