# ImageZoomAddin - reset Office's "this add-in is broken" state
#
# 用途：当 Word/Excel 弹窗提示 "此加载项出现问题... 是否禁用" 时运行本脚本。
# Office 会把出过问题的加载项写进 Resiliency 下的几个键，之后即使 LoadBehavior=3
# 也会静默跳过加载。本脚本只清这些记录 + 复位 LoadBehavior，不碰其他加载项。
#
# 写的是 HKCU，不需要管理员权限；但运行前请先关闭所有 Word/Excel。

$ErrorActionPreference = 'Continue'
$pattern = 'ImageZoom|4F2B8A61'
$versions = @('16.0', '15.0', '14.0')
$hosts = @('Word', 'Excel', 'PowerPoint')
$subs = @('DisabledItems', 'StartupItems', 'DisabledAddins')
$progid = 'ImageZoomAddin.Connect'
$cleared = 0

Write-Host ''
Write-Host '=== ImageZoomAddin reset ===' -ForegroundColor Cyan

$running = Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue
if ($running) {
    Write-Host ''
    Write-Host '  WARNING: Word/Excel/PowerPoint is still running. Close it first,' -ForegroundColor Yellow
    Write-Host '  otherwise Office will flush its in-memory state back on exit.' -ForegroundColor Yellow
    foreach ($p in $running) { Write-Host ('    ' + $p.ProcessName + ' pid=' + $p.Id) }
}

foreach ($ver in $versions) {
    foreach ($app in $hosts) {
        foreach ($sub in $subs) {
            $path = 'HKCU:\Software\Microsoft\Office\' + $ver + '\' + $app + '\Resiliency\' + $sub
            if (-not (Test-Path $path)) { continue }
            $item = Get-Item $path
            $names = @($item.GetValueNames())
            foreach ($name in $names) {
                $raw = $item.GetValue($name)
                $text = ''
                if ($raw -is [byte[]]) {
                    $text = [System.Text.Encoding]::Unicode.GetString($raw)
                    $text = $text + ' ' + [System.Text.Encoding]::ASCII.GetString($raw)
                } else {
                    $text = [string]$raw
                }
                if ($text -match $pattern) {
                    Remove-ItemProperty -Path $path -Name $name -Force -ErrorAction SilentlyContinue
                    Write-Host ('  cleared  ' + $ver + '\' + $app + '\' + $sub) -ForegroundColor Green
                    $cleared++
                }
            }
        }
    }
}

# 复位 LoadBehavior（Office 加载失败后会把它从 3 改成 2）
$fixed = 0
foreach ($app in $hosts) {
    $bases = @(
        'HKCU:\Software\Microsoft\Office\16.0\' + $app + '\Addins',
        'HKCU:\Software\Microsoft\Office\' + $app + '\Addins'
    )
    foreach ($base in $bases) {
        $key = $base + '\' + $progid
        if (-not (Test-Path $key)) { continue }
        $old = (Get-ItemProperty $key).LoadBehavior
        Set-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3 -Type DWord -Force
        Write-Host ('  LoadBehavior  ' + $key + ' : ' + $old + ' -> 3') -ForegroundColor Green
        $fixed++
    }
}

Write-Host ''
if ($cleared -eq 0 -and $fixed -eq 0) {
    Write-Host '  Nothing to fix (already clean).' -ForegroundColor DarkGray
} else {
    Write-Host ('  Done. cleared ' + $cleared + ' record(s), reset ' + $fixed + ' key(s).') -ForegroundColor Green
}

Write-Host ''
Write-Host '  Next: start Word/Excel/PowerPoint again, then run check.bat to confirm.' -ForegroundColor Cyan
Write-Host ''
Read-Host 'Press Enter to close'
