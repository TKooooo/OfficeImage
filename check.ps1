# ImageZoomAddin self-check
# 关键结论：托管 COM 类必须注册在 HKLM 才能被 mscoree 解析。
# 注册在 HKCU\Software\Classes 时 CoCreateInstance 会返回 0x80070002。
$ErrorActionPreference = 'Continue'
$clsid  = '{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}'
$progid = 'ImageZoomAddin.Connect'
$dll    = Join-Path $PSScriptRoot 'ImageZoomAddin.dll'

Write-Host ''
Write-Host '===== 1. DLL =====' -ForegroundColor Cyan
$dllVer = $null
if (Test-Path $dll) {
    $i = Get-Item $dll
    $dllVer = $i.VersionInfo.FileVersion
    Write-Host ("  path     : " + $i.FullName)
    Write-Host ("  size     : " + $i.Length + " bytes   " + $i.LastWriteTime)
    Write-Host ("  version  : " + $dllVer)
} else {
    Write-Host '  MISSING - run build.bat first' -ForegroundColor Red
}

Write-Host ''
Write-Host '===== 2. COM registration (must be in HKLM) =====' -ForegroundColor Cyan
$hklm64 = "HKLM:\SOFTWARE\Classes\CLSID\$clsid\InprocServer32"
$hklm32 = "HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\$clsid\InprocServer32"
$hkcu   = "HKCU:\Software\Classes\CLSID\$clsid\InprocServer32"

foreach ($pair in @(@('HKLM 64-bit', $hklm64, $true), @('HKLM 32-bit (Wow6432Node)', $hklm32, $false))) {
    $label = $pair[0]; $path = $pair[1]; $required = $pair[2]
    if (Test-Path $path) {
        $v = Get-ItemProperty $path
        Write-Host ("  {0}: present" -f $label) -ForegroundColor Green
        Write-Host ("     (default) = " + $v.'(default)')
        Write-Host ("     Class     = " + $v.Class)
        Write-Host ("     Assembly  = " + $v.Assembly)
        Write-Host ("     CodeBase  = " + $v.CodeBase)
        if ($v.'(default)' -ne 'mscoree.dll') {
            Write-Host '     !! InprocServer32 default must be mscoree.dll' -ForegroundColor Red
        }
        $cb = $v.CodeBase
        if ($cb -and $cb -match '^file:///(.+)$') {
            $file = $Matches[1] -replace '/', '\'
            if (-not (Test-Path $file)) { Write-Host ("     !! CodeBase file does not exist: " + $file) -ForegroundColor Red }
        }
    } else {
        if ($required) { Write-Host ("  {0}: MISSING -> run install.bat as administrator" -f $label) -ForegroundColor Red }
        else { Write-Host ("  {0}: absent (only needed for 32-bit Office)" -f $label) -ForegroundColor DarkGray }
    }
}

if (Test-Path $hkcu) {
    Write-Host '  HKCU registration also present (legacy). Harmless, but HKLM is what matters.' -ForegroundColor DarkYellow
}

Write-Host ''
Write-Host '  Office add-in keys:'
foreach ($h in @('Word','Excel','PowerPoint')) {
    $p = "HKCU:\Software\Microsoft\Office\$h\Addins\$progid"
    if (Test-Path $p) {
        $lb = (Get-ItemProperty $p).LoadBehavior
        $note = ''
        if ($lb -eq 2) { $note = '  <- "failed / disabled", must be 3' }
        if ($lb -eq 3) { $note = '  <- OK (load at startup)' }
        if ($lb -eq 0) { $note = '  <- "do not load"' }
        Write-Host ("  {0,-6} LoadBehavior = {1}{2}" -f $h, $lb, $note)
    } else {
        Write-Host ("  {0,-6} NOT REGISTERED" -f $h) -ForegroundColor Red
    }
}

Write-Host ''
Write-Host '  Office resiliency records (these silently skip the add-in):'
$found = $false
foreach ($app in @('Word','Excel','PowerPoint')) {
    foreach ($sub in @('DisabledItems','StartupItems','DisabledAddins')) {
        $d = "HKCU:\Software\Microsoft\Office\16.0\$app\Resiliency\$sub"
        if (-not (Test-Path $d)) { continue }
        $item = Get-Item $d
        foreach ($n in $item.GetValueNames()) {
            $raw = $item.GetValue($n)
            $txt = if ($raw -is [byte[]]) { [System.Text.Encoding]::Unicode.GetString($raw) } else { "$raw" }
            if ($txt -match 'ImageZoom|4F2B8A61') {
                $found = $true
                Write-Host ("  !! {0}\{1} -> {2}" -f $app, $sub, ($txt -replace '[^\x20-\x7E]','')) -ForegroundColor Red
            }
        }
    }
}
if (-not $found) { Write-Host '  clean' -ForegroundColor Green }

Write-Host ''
Write-Host '===== 3. COM activation test (what Office does) =====' -ForegroundColor Cyan
$acts = $false
try {
    $o = New-Object -ComObject $progid
    Write-Host '  OK - COM object created' -ForegroundColor Green
    $acts = $true
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch { }
} catch {
    $h = 0
    try { $h = $_.Exception.HResult } catch { }
    Write-Host ("  FAILED: " + $_.Exception.Message) -ForegroundColor Red
    Write-Host ("  HRESULT: 0x{0:X8}" -f $h) -ForegroundColor Red
    if (($h -band 0xFFFF) -eq 2) {
        Write-Host '  0x80070002 = mscoree cannot resolve the assembly for this CLSID.' -ForegroundColor Yellow
        Write-Host '  Most common cause: registered only in HKCU. Re-run install.bat as administrator.' -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host '===== 4. Add-in log =====' -ForegroundColor Cyan
Write-Host '    "ctor:"        = COM object was created'
Write-Host '    "OnConnection" = Office accepted it as a COM add-in'
Write-Host '    "diag#"        = the 200ms poll loop is running'
Write-Host '    "magnifier:"   = the button was placed on screen'
$log = Join-Path $env:LOCALAPPDATA 'ImageZoomAddin\addin.log'
$hasOnConn = $false
if (Test-Path $log) {
    Write-Host ("  " + $log + "  (" + (Get-Item $log).Length + " bytes)")
    $lines = Get-Content $log
    $hasOnConn = ($lines | Where-Object { $_ -match 'OnConnection' }).Count -gt 0
    $lines | Select-Object -Last 25 | ForEach-Object { Write-Host ("  " + $_) }
} else {
    Write-Host '  NO LOG FILE - the COM object was never created inside Office.' -ForegroundColor Yellow
}

Write-Host ''
Write-Host '===== 5. Running Office processes =====' -ForegroundColor Cyan
$procs = Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue
if ($procs) {
    foreach ($p in $procs) { Write-Host ("  " + $p.ProcessName + " pid=" + $p.Id) }
    Write-Host '  NOTE: restart Office after any registration change.' -ForegroundColor Yellow
} else {
    Write-Host '  none running'
}

Write-Host ''
Write-Host '===== verdict =====' -ForegroundColor Cyan
if (-not (Test-Path $hklm64)) {
    Write-Host '  BROKEN: not registered in HKLM. Run install.bat (as administrator).' -ForegroundColor Red
} elseif (-not $acts) {
    Write-Host '  BROKEN: HKLM registered but activation still fails. Send this output.' -ForegroundColor Red
} elseif (-not $hasOnConn) {
    Write-Host '  Registration is fine, but Office has not loaded the add-in yet.' -ForegroundColor Yellow
    Write-Host '  => Start Word/Excel/PowerPoint, open a document, then run check.bat again.' -ForegroundColor Yellow
} else {
    Write-Host '  OK: registered, activated, and Office ran OnConnection.' -ForegroundColor Green
    if (-not ($lines | Where-Object { $_ -match 'diag#' })) {
        Write-Host '  But the poll loop produced no diag# line - send the log.' -ForegroundColor Yellow
    }
}
Write-Host ''
Read-Host 'Press Enter to close'
