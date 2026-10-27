# =============================================================================
#  ImageZoomAddin - uninstall (administrator required: removes HKLM entries)
# =============================================================================
$ErrorActionPreference = 'Continue'

$Clsid  = '{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}'
$ProgId = 'ImageZoomAddin.Connect'

Write-Host ''
Write-Host '=== Image Zoom Preview : uninstall ===' -ForegroundColor Cyan

# --- CLSID: HKLM 64-bit, HKLM 32-bit, plus legacy HKCU entry ---
$k1 = 'HKLM:\SOFTWARE\Classes\CLSID\' + $Clsid
$k2 = 'HKLM:\SOFTWARE\Classes\Wow6432Node\CLSID\' + $Clsid
$k3 = 'HKCU:\Software\Classes\CLSID\' + $Clsid
if (Test-Path $k1) { Remove-Item $k1 -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $k1) -ForegroundColor Green }
if (Test-Path $k2) { Remove-Item $k2 -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $k2) -ForegroundColor Green }
if (Test-Path $k3) { Remove-Item $k3 -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $k3) -ForegroundColor Green }

# --- ProgID ---
$p1 = 'HKLM:\SOFTWARE\Classes\' + $ProgId
$p2 = 'HKLM:\SOFTWARE\Classes\Wow6432Node\' + $ProgId
$p3 = 'HKCU:\Software\Classes\' + $ProgId
if (Test-Path $p1) { Remove-Item $p1 -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $p1) -ForegroundColor Green }
if (Test-Path $p2) { Remove-Item $p2 -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $p2) -ForegroundColor Green }
if (Test-Path $p3) { Remove-Item $p3 -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $p3) -ForegroundColor Green }

# --- Office add-in registration ---
$w = 'HKCU:\Software\Microsoft\Office\Word\Addins\' + $ProgId
$x = 'HKCU:\Software\Microsoft\Office\Excel\Addins\' + $ProgId
if (Test-Path $w) { Remove-Item $w -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $w) -ForegroundColor Green }
if (Test-Path $x) { Remove-Item $x -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $x) -ForegroundColor Green }

# --- WPS entries (normally never created, cleaned defensively) ---
foreach ($base in @('HKCU:\Software\Kingsoft\Office\WPS\Addins',
                    'HKCU:\Software\Kingsoft\Office\6.0\common\jsaddins',
                    'HKLM:\SOFTWARE\Kingsoft\Office\WPS\Addins',
                    'HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office\WPS\Addins')) {
    $k = "$base\$ProgId"
    if (Test-Path $k) { Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue; Write-Host ('  removed: ' + $k) -ForegroundColor Green }
}

Write-Host ''
Write-Host 'DONE. Restart Word/Excel/PowerPoint to fully unload.' -ForegroundColor Cyan
Write-Host ''
