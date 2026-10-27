@echo off
setlocal
chcp 65001 >nul

echo ============================================
echo  Image Zoom Add-in - build
echo ============================================

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
    echo [ERROR] .NET Framework 4.x C# compiler not found.
    echo         Expected at %%WINDIR%%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
    if not defined IZ_NOPAUSE pause
    exit /b 1
)

set "FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319"
if not exist "%FW%\System.Drawing.dll" set "FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319"

set "OUT=%~dp0ImageZoomAddin.dll"
rem TEMP OUTPUT MUST KEEP THE SAME FILE NAME:
rem csc uses the /out file name as the assembly name, so building to
rem "ImageZoomAddin.new.dll" makes the assembly "ImageZoomAddin.new",
rem which no longer matches Assembly=ImageZoomAddin in the registry and
rem Office fails to activate it (0x80131040).
rem NOTE: keep batch comments ASCII-only -- cmd reads this file in the
rem console code page, so UTF-8 comments can get truncated and the tail
rem is then executed as a command ("'xxx' is not recognized").
set "TMPDIR=%~dp0.obj"
set "TMPOUT=%TMPDIR%\ImageZoomAddin.dll"
set "SRC=%~dp0src\ImageZoomAddin.cs"

rem ---- optional extra defines, e.g. set IZ_DEFINE=IZ_TEST_NOFG ----
set "DEF="
if defined IZ_DEFINE set "DEF=/define:%IZ_DEFINE%"

rem ---- locate Microsoft's official Extensibility PIA (IDTExtensibility2) ----
set "EXTDLL="
if exist "%WINDIR%\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll" (
    set "EXTDLL=%WINDIR%\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll"
)
if not defined EXTDLL if exist "%CommonProgramFiles(x86)%\Microsoft Shared\MSEnv\PublicAssemblies\extensibility.dll" (
    set "EXTDLL=%CommonProgramFiles(x86)%\Microsoft Shared\MSEnv\PublicAssemblies\extensibility.dll"
)
if not defined EXTDLL if exist "%CommonProgramFiles%\Microsoft Shared\MSEnv\PublicAssemblies\extensibility.dll" (
    set "EXTDLL=%CommonProgramFiles%\Microsoft Shared\MSEnv\PublicAssemblies\extensibility.dll"
)
if not defined EXTDLL (
    echo [ERROR] Extensibility PIA not found.
    echo         Expected: %%WINDIR%%\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll
    echo         This file ships with Office/Visual Studio. Without it the add-in
    echo         cannot be compiled correctly.
    if not defined IZ_NOPAUSE pause
    exit /b 1
)

if not exist "%SRC%" (
    echo [ERROR] Source not found: %SRC%
    if not defined IZ_NOPAUSE pause
    exit /b 1
)

echo Compiler : %CSC%
echo Source   : %SRC%
echo Extensib.: %EXTDLL%
echo Output   : %OUT%
echo.

rem Build into a temp dir first, swap only after a successful compile:
rem while Word/Excel/WPS is running the live DLL is usually mapped, so a
rem direct /out would fail with "file is being used by another process".
if not exist "%TMPDIR%" mkdir "%TMPDIR%" >nul 2>&1
if exist "%TMPOUT%" del /f /q "%TMPOUT%" >nul 2>&1

"%CSC%" /nologo /target:library /platform:anycpu /optimize+ /debug- /codepage:65001 %DEF% ^
 /lib:"%FW%" ^
 /out:"%TMPOUT%" ^
 /r:System.dll ^
 /r:System.Core.dll ^
 /r:System.Drawing.dll ^
 /r:System.Windows.Forms.dll ^
 /r:System.Xml.dll ^
 /r:System.Xml.Linq.dll ^
 /r:System.IO.Compression.dll ^
 /r:System.IO.Compression.FileSystem.dll ^
 /r:Microsoft.CSharp.dll ^
 /r:"%EXTDLL%" ^
 "%SRC%"

if errorlevel 1 goto failed

if not exist "%TMPOUT%" goto failed

rem ---- install the freshly built DLL ----
rem A mapped file cannot be overwritten, but it usually can be renamed:
rem move the old one aside first, then drop the new file in place.
set "PREV=%~dp0ImageZoomAddin.previous.dll"
del /f /q "%PREV%" >nul 2>&1
if exist "%PREV%" set "PREV=%~dp0ImageZoomAddin.previous2.dll"
del /f /q "%PREV%" >nul 2>&1
set "SWAPPED=0"
move /y "%TMPOUT%" "%OUT%" >nul 2>&1
if not errorlevel 1 set "SWAPPED=1"
if "%SWAPPED%"=="0" (
    if exist "%OUT%" move /y "%OUT%" "%PREV%" >nul 2>&1
    move /y "%TMPOUT%" "%OUT%" >nul 2>&1
    if not errorlevel 1 set "SWAPPED=1"
)
if "%SWAPPED%"=="0" (
    echo.
    echo [ERROR] 编译成功，但替换 ImageZoomAddin.dll 失败 - 文件被其它程序占用。
    echo         请关闭所有 Word / Excel / WPS 窗口后重跑本脚本。
    if not defined IZ_NOPAUSE pause
    exit /b 1
)
del /f /q "%~dp0ImageZoomAddin.previous.dll" >nul 2>&1

echo.
echo ============================================
echo  BUILD OK
for %%F in ("%OUT%") do echo  %%~fF   (%%~zF bytes)
echo ============================================
echo  Verify: run  check.bat
echo.
if not defined IZ_NOPAUSE pause
exit /b 0

:failed
echo.
echo ============================================
echo  BUILD FAILED - copy the error lines above
echo ============================================
echo.
if not defined IZ_NOPAUSE pause
exit /b 1
