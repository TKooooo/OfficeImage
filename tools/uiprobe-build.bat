@echo off
rem ===========================================================================
rem  Dev probes (not part of the add-in):
rem    panelprobe  : render the crop panel to png (layout / icons / hover)
rem    actionprobe : attach a real Word instance and call shape/ratio/reset
rem    menuprobe   : open panel -> pick item -> apply -> close panel
rem  Usage: run build.bat first, then _probe\*.exe with the add-in dll path.
rem ===========================================================================
setlocal
set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
set "OUT=%~dp0_probe"
if not exist "%OUT%" mkdir "%OUT%"

set "REF=/r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:Microsoft.CSharp.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll"

"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\panelprobe.exe"  "%~dp0panelprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\actionprobe.exe" "%~dp0actionprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\menuprobe.exe"   "%~dp0menuprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\xlprobe.exe"    "%~dp0xlprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\shapemap.exe"   "%~dp0shapemap.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\xlflow.exe"     "%~dp0xlflow.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\xlmenuprobe.exe" "%~dp0xlmenuprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\xlunits.exe"    "%~dp0xlunits.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 /unsafe %REF% /out:"%OUT%\xlrect.exe" "%~dp0xlrect.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 /unsafe %REF% /out:"%OUT%\xlplace.exe" "%~dp0xlplace.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\rotprobe.exe" "%~dp0rotprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\xlrotate.exe" "%~dp0xlrotate.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\btnprobe.exe" "%~dp0btnprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\compare.exe" "%~dp0compare.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\layeredprobe.exe" "%~dp0layeredprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 /unsafe %REF% /r:System.Xml.Linq.dll /out:"%OUT%\pptprobe.exe" "%~dp0pptprobe.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\pptflow.exe" "%~dp0pptflow.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\pptlive.exe" "%~dp0pptlive.cs"
if errorlevel 1 goto failed
"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /out:"%OUT%\pptaction.exe" "%~dp0pptaction.cs"
if errorlevel 1 goto failed

"%CSC%" /nologo /target:exe /platform:anycpu /codepage:65001 %REF% /r:System.Windows.Forms.dll /out:"%OUT%\uicheck.exe" "%~dp0uicheck.cs"
if errorlevel 1 goto failed

echo BUILD OK  -^> %OUT%
exit /b 0

:failed
echo BUILD FAILED
exit /b 1
