@echo off
setlocal
chcp 65001 >nul 2>&1
title Image Zoom Preview - Install

rem ===========================================================================
rem  需要管理员权限：托管 COM 类必须注册到 HKLM 才能被 mscoree 解析
rem  （写在 HKCU 会导致 CoCreateInstance 返回 0x80070002，Office 静默禁用）
rem ===========================================================================

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo   Requesting administrator rights ...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"

echo.
pause
