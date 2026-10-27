# =============================================================================
#  ImageZoomAddin - 安装（需要管理员权限）
#
#  为什么必须写 HKLM：
#    托管 COM 类（InprocServer32 = mscoree.dll）在 HKCU\Software\Classes 下注册时，
#    mscoree 解析 CLSID -> Assembly 映射会失败，CoCreateInstance 返回
#    0x80070002（系统找不到指定的文件），Office 随即静默禁用加载项。
#    同样的键值写进 HKLM\SOFTWARE\Classes 后激活立即成功。
#    该结论为实测结果（见 tools/ 下的验证脚本），与程序集版本号无关。
#
#  同时写 64 位视图与 Wow6432Node，32/64 位 Office 都能加载。
# =============================================================================

$ErrorActionPreference = 'Stop'

$Clsid    = '{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}'
$ProgId   = 'ImageZoomAddin.Connect'
$Friendly = 'Image Zoom Preview'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$dll = Join-Path $scriptDir 'ImageZoomAddin.dll'

Write-Host ''
Write-Host '=== Image Zoom Preview : install ===' -ForegroundColor Cyan

# ---------- 0. 前置检查 ----------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host '  ERROR: administrator rights required (HKLM registration).' -ForegroundColor Red
    Write-Host '         Right-click install.bat -> Run as administrator.'
    exit 1
}

if (-not (Test-Path $dll)) {
    Write-Host "  ERROR: DLL not found: $dll" -ForegroundColor Red
    Write-Host '         Run build.bat first.'
    exit 1
}

# ---------- 1. 读出版本，避免版本号写死 ----------
$ver = $null
try {
    $an = [System.Reflection.AssemblyName]::GetAssemblyName($dll)
    $ver = $an.Version.ToString()
} catch {
    $ver = (Get-Item $dll).VersionInfo.FileVersion
}
if (-not $ver) { $ver = '1.0.0.0' }
while (($ver -split '\.').Count -lt 4) { $ver = $ver + '.0' }

$asmFull = "ImageZoomAddin, Version=$ver, Culture=neutral, PublicKeyToken=null"
$url     = 'file:///' + ($dll -replace '\\', '/')

Write-Host ("  DLL          : {0}" -f $dll)
Write-Host ("  Assembly     : {0}" -f $asmFull)
Write-Host ("  CodeBase     : {0}" -f $url)

# ---------- 2. 写 CLSID（HKLM，64 位视图 + Wow6432Node） ----------
function Install-Clsid([string]$classesRoot) {
    $base = Join-Path $classesRoot ('CLSID\' + $Clsid)
    New-Item -Path $base -Force | Out-Null
    Set-ItemProperty -Path $base -Name '(default)' -Value $ProgId

    $ip = Join-Path $base 'InprocServer32'
    New-Item -Path $ip -Force | Out-Null
    Set-ItemProperty -Path $ip -Name '(default)'      -Value 'mscoree.dll'
    Set-ItemProperty -Path $ip -Name 'ThreadingModel' -Value 'Both'
    Set-ItemProperty -Path $ip -Name 'Class'          -Value $ProgId
    Set-ItemProperty -Path $ip -Name 'Assembly'       -Value $asmFull
    Set-ItemProperty -Path $ip -Name 'RuntimeVersion' -Value 'v4.0.30319'
    Set-ItemProperty -Path $ip -Name 'CodeBase'       -Value $url

    Set-ItemProperty -Path $base -Name 'ProgId' -Value $ProgId

    # RegAsm 也会写这个 .NET 分类，补上更接近标准注册
    New-Item -Path (Join-Path $base 'Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}') -Force | Out-Null

    # ProgID -> CLSID
    $progKey = Join-Path $classesRoot $ProgId
    New-Item -Path $progKey -Force | Out-Null
    Set-ItemProperty -Path $progKey -Name '(default)' -Value $Friendly
    $progClsidKey = Join-Path $progKey 'CLSID'
    New-Item -Path $progClsidKey -Force | Out-Null
    Set-ItemProperty -Path $progClsidKey -Name '(default)' -Value $Clsid
}

Install-Clsid 'HKLM:\SOFTWARE\Classes'
Write-Host '  registered   : HKLM\SOFTWARE\Classes (64-bit)' -ForegroundColor Green
Install-Clsid 'HKLM:\SOFTWARE\Classes\Wow6432Node'
Write-Host '  registered   : HKLM\SOFTWARE\Classes\Wow6432Node (32-bit)' -ForegroundColor Green

# ---------- 3. 清掉旧的 HKCU 注册（旧版写在这里，是无效的） ----------
foreach ($p in @("HKCU:\Software\Classes\CLSID\$Clsid",
                 "HKCU:\Software\Classes\$ProgId")) {
    if (Test-Path $p) {
        Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host ("  removed stale: {0}" -f $p) -ForegroundColor DarkGray
    }
}

# ---------- 4. 给 Word / Excel / PowerPoint 挂加载项 ----------
foreach ($app in @('Word', 'Excel', 'PowerPoint')) {
    $k = "HKCU:\Software\Microsoft\Office\$app\Addins\$ProgId"
    New-Item -Path $k -Force | Out-Null
    New-ItemProperty -Path $k -Name 'FriendlyName' -Value $Friendly -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $k -Name 'Description'  -Value 'WPS-like picture magnifier preview' -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $k -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $k -Name 'CommandLineSafe' -Value 0 -PropertyType DWord -Force | Out-Null
    Write-Host ("  add-in       : {0} LoadBehavior=3" -f $app) -ForegroundColor Green
}

# ---------- 4.5 确保不会挂到 WPS 上 ----------
#   WPS 自带图片放大预览，不需要本加载项参与。
#   这里清掉 WPS 可能存在的同名注册；即便某些 WPS 版本会复用上面那份
#   Office\*\Addins 注册把加载项拉起来，加载项内部也有进程白名单兜底
#   （只放行 WINWORD.EXE / EXCEL.EXE，其余宿主 OnConnection 直接返回）。
$wpsRoots = @(
    'HKCU:\Software\Kingsoft\Office\WPS\Addins',
    'HKCU:\Software\Kingsoft\Office\6.0\common\jsaddins',
    'HKLM:\SOFTWARE\Kingsoft\Office\WPS\Addins',
    'HKLM:\SOFTWARE\WOW6432Node\Kingsoft\Office\WPS\Addins'
)
$wpsCleaned = 0
foreach ($base in $wpsRoots) {
    $k = "$base\$ProgId"
    if (Test-Path $k) {
        Remove-Item $k -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host ("  removed WPS entry: {0}" -f $k) -ForegroundColor DarkGray
        $wpsCleaned++
    }
}
Write-Host "  WPS          : not registered (add-in itself only runs in WINWORD/EXCEL)" -ForegroundColor Green

# ---------- 5. 清除 Office 的"禁用项"记录 ----------
#   Resiliency\DisabledItems : 加载项被禁用
#   Resiliency\StartupItems  : 启动时出过问题的项（Office 会跳过，即使 LoadBehavior=3）
#   两者都是二进制 blob，里面含 DLL 路径或 ProgId，按内容匹配删除
function Clear-Resiliency([string]$app) {
    $verRoot = 'HKCU:\Software\Microsoft\Office'
    $versions = @()
    if (Test-Path $verRoot) {
        $versions = @(Get-ChildItem $verRoot -ErrorAction SilentlyContinue |
                      Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
                      ForEach-Object { $_.PSChildName })
    }
    if ($versions.Count -eq 0) { $versions = @('16.0') }

    foreach ($ver in $versions) {
    foreach ($sub in @('DisabledItems', 'StartupItems', 'DisabledAddins')) {
        $d = "HKCU:\Software\Microsoft\Office\$ver\$app\Resiliency\$sub"
        if (-not (Test-Path $d)) { continue }
        $item = Get-Item $d
        foreach ($n in $item.GetValueNames()) {
            $raw = $item.GetValue($n)
            $txt = ''
            if ($raw -is [byte[]]) {
                $txt = [System.Text.Encoding]::Unicode.GetString($raw)
            } elseif ($raw -is [string]) {
                $txt = $raw
            }
            if ($txt -match 'ImageZoom|4F2B8A61') {
                Remove-ItemProperty -Path $d -Name $n -Force -ErrorAction SilentlyContinue
                Write-Host ("  cleared {0}\{1}: {2}" -f $app, $sub, ($txt -replace '[^\x20-\x7E]', '')) -ForegroundColor Yellow
            }
        }
    }
    }
}
foreach ($app in @('Word', 'Excel', 'PowerPoint')) { Clear-Resiliency $app }

# ---------- 6. 验证 COM 激活 ----------
Write-Host ''
Write-Host '  verifying COM activation ...'
$ok = $false
try {
    $o = New-Object -ComObject $ProgId
    $ok = $true
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) } catch { }
    Write-Host '  COM activation: OK' -ForegroundColor Green
} catch {
    Write-Host ('  COM activation: FAILED -> ' + $_.Exception.Message) -ForegroundColor Red
}

Write-Host ''
if ($ok) {
    Write-Host 'DONE. Close ALL Word/Excel/PowerPoint windows, then reopen and click a picture.' -ForegroundColor Cyan
    Write-Host ('Log: ' + (Join-Path $env:LOCALAPPDATA 'ImageZoomAddin\addin.log'))
} else {
    Write-Host 'Registration written but activation still fails. Run check.bat and send the output.' -ForegroundColor Yellow
}
Write-Host ''
