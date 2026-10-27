; =============================================================================
;  OfficeImage 安装脚本（Inno Setup 6）
;
;  产物：dist\officeimage.exe —— 单文件、自带卸载、无需 .NET 也能启动
;
;  这里做的事情，和 install.ps1 完全等价，只是换成标准安装程序：
;    1. 把 ImageZoomAddin.dll 放进 {app}
;    2. 注册托管 COM 类（CLSID + ProgID），64 位视图和 32 位视图都写
;       —— 32 位 Office 会把 HKLM\Software\Microsoft\Office 重定向到
;          Wow6432Node，少写一个视图就有一半机器加载不到
;    3. 给 Word / Excel / PowerPoint 挂加载项（LoadBehavior=3）
;    4. 清掉 Office 可能残留的"禁用项"记录
;    5. 在"程序和功能"里生成卸载项（Inno 自动完成）
; =============================================================================

#define AppName      "OfficeImage"
#define AppPublisher "OfficeImage"
#define AppExeName   "officeimage.exe"
#define ProgId       "ImageZoomAddin.Connect"
#define ClsidG       "{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"
#define DllName      "ImageZoomAddin.dll"

; 版本号直接读 DLL，改源码里的 AssemblyVersion 之后重新编译即可，不用改这里。
; AssemblyVersion 与 AssemblyFileVersion 在源码里保持一致，安装包才能正确
; 写出 InprocServer32\Assembly（版本对不上 Office 会报 0x80131040 静默退出）。
#define AppVer       GetVersionNumbersString("..\" + DllName)
#define AsmFull      "ImageZoomAddin, Version=" + AppVer + ", Culture=neutral, PublicKeyToken=null"

[Setup]
; AppId 是卸载时的身份标识，一旦发布就不要再改，否则会变成两个"已安装程序"
AppId={{9E7A5D3C-2B14-4C8E-A6F9-3D51B8C40A72}
AppName={#AppName}
AppVersion={#AppVer}
AppVerName={#AppName} {#AppVer}
AppPublisher={#AppPublisher}
AppComments=在 Word / Excel / PowerPoint 中点击图片，浮出放大镜、裁剪、旋转按钮
DefaultDirName={autopf}\OfficeImage
DefaultGroupName=OfficeImage
DisableProgramGroupPage=yes
DisableWelcomePage=no
; 必须管理员：托管 COM 类只有在 HKLM\SOFTWARE\Classes 下注册，mscoree 才能
; 把 CLSID 解析到程序集（写 HKCU 会返回 0x80070002，Office 随即静默禁用加载项）
PrivilegesRequired=admin
; 64 位系统装进 Program Files，32 位系统自动退回 Program Files (x86)
ArchitecturesInstallIn64BitMode=x64compatible
; Windows 7 起（引导程序 PE 子系统版本 6.1，Win7/8/10/11 均可直接双击运行）
MinVersion=6.1
OutputDir=..\dist
OutputBaseFilename=officeimage
SetupIconFile=..\assets\officeimage.ico
WizardImageFile=..\assets\wizard-large.bmp
WizardSmallImageFile=..\assets\wizard-small.bmp
UninstallDisplayName={#AppName} 图片放大预览（Word / Excel / PowerPoint）
UninstallDisplayIcon={app}\officeimage.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ShowLanguageDialog=no
SetupLogging=yes
; 升级时若 Office 正开着（DLL 被占用），提示用户关闭，而不是报错退出
CloseApplications=yes
RestartApplications=no
VersionInfoVersion={#AppVer}
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVer}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription=OfficeImage 安装程序
VersionInfoCopyright=Copyright (C) 2026 {#AppPublisher}

[Languages]
Name: "zh"; MessagesFile: "Languages\ChineseSimplified.isl"

[Files]
Source: "..\{#DllName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\officeimage.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "使用说明.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\OfficeImage 使用说明"; Filename: "{app}\使用说明.txt"
Name: "{group}\卸载 OfficeImage"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\使用说明.txt"; Description: "查看使用说明"; Flags: postinstall shellexec nowait skipifsilent

[Registry]
; -----------------------------------------------------------------------------
;  COM 注册 —— 64 位视图
; -----------------------------------------------------------------------------
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"; ValueType: string; ValueName: ""; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"; ValueType: string; ValueName: "ProgId"; ValueData: "{#ProgId}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#ProgId}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFull}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetCodeBase}"
Root: HKLM64; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"; ValueType: none; ValueName: ""; ValueData: ""
Root: HKLM64; Subkey: "Software\Classes\{#ProgId}"; ValueType: string; ValueName: ""; ValueData: "OfficeImage 图片放大预览"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"

; -----------------------------------------------------------------------------
;  COM 注册 —— 32 位视图（32 位 Office / 32 位系统）
; -----------------------------------------------------------------------------
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"; ValueType: string; ValueName: ""; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"; ValueType: string; ValueName: "ProgId"; ValueData: "{#ProgId}"; Flags:
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags:
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"; Flags:
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "{#ProgId}"; Flags:
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "{#AsmFull}"; Flags:
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"; Flags:
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:GetCodeBase}"; Flags:
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"; ValueType: none; ValueName: ""; ValueData: ""; Flags:
Root: HKLM32; Subkey: "Software\Classes\{#ProgId}"; ValueType: string; ValueName: ""; ValueData: "OfficeImage 图片放大预览"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}"; Flags:

; -----------------------------------------------------------------------------
;  Office 加载项 —— 三个宿主 × 两个视图
;  LoadBehavior=3（加载时启动），CommandLineSafe=0
; -----------------------------------------------------------------------------
Root: HKLM64; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OfficeImage 图片放大预览"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "点击图片后浮出放大镜 / 裁剪 / 旋转按钮"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Flags: uninsdeletekey

Root: HKLM64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OfficeImage 图片放大预览"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "点击图片后浮出放大镜 / 裁剪 / 旋转按钮"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Flags: uninsdeletekey

Root: HKLM64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OfficeImage 图片放大预览"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "点击图片后浮出放大镜 / 裁剪 / 旋转按钮"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#ProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Flags: uninsdeletekey

Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OfficeImage 图片放大预览"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "点击图片后浮出放大镜 / 裁剪 / 旋转按钮"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Flags: uninsdeletekey

Root: HKLM32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OfficeImage 图片放大预览"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "点击图片后浮出放大镜 / 裁剪 / 旋转按钮"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\Excel\Addins\{#ProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Flags: uninsdeletekey

Root: HKLM32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OfficeImage 图片放大预览"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "点击图片后浮出放大镜 / 裁剪 / 旋转按钮"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\{#ProgId}"; ValueType: dword; ValueName: "CommandLineSafe"; ValueData: "0"; Flags: uninsdeletekey

[Code]

const
  cn_ProgId = '{4F2B8A61-9C3D-4C1E-9F1A-7D2E5B6A1C08}';   // 只是给下面的清理代码用
  cn_Apps   = 'Word,Excel,PowerPoint';
  cn_NDP    = 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full';

{ CodeBase 必须是 URL 形式，且用正斜杠；mscoree 按 URL 解析这个值。
  注意：Inno 的 Pascal 里反斜杠是转义符，写 '\' 会被当成"转义后的引号"，
  连字符串都断不干净，所以这里逐字符换成 #92 -> '/' }
function ToUrlPath(const p: String): String;
var
  i: Integer;
  s: String;
begin
  s := '';
  for i := 1 to Length(p) do
  begin
    if p[i] = #92 then
      s := s + '/'
    else
      s := s + Copy(p, i, 1);
  end;
  Result := s;
end;

function GetCodeBase(Param: String): String;
begin
  Result := 'file:///' + ToUrlPath(ExpandConstant('{app}')) + '/ImageZoomAddin.dll';
end;

{ .NET Framework 4.x：4.5 及以后有 Release 值，4.0 只有 Install=1 }
function NdpValue_Release(RootKey: Integer): Cardinal;
var
  v: Cardinal;
begin
  Result := 0;
  if RegQueryDWordValue(RootKey, cn_NDP, 'Release', v) then
    Result := v;
end;

function HasDotNet4: Boolean;
var
  v: Cardinal;
begin
  Result := (NdpValue_Release(HKLM64) >= 378389) or (NdpValue_Release(HKLM32) >= 378389);
  if not Result then
  begin
    if RegQueryDWordValue(HKLM64, cn_NDP, 'Install', v) and (v = 1) then
      Result := True
    else if RegQueryDWordValue(HKLM32, cn_NDP, 'Install', v) and (v = 1) then
      Result := True;
  end;
end;

function AppPathExists(const Exe: String): Boolean;
var
  s: String;
begin
  Result :=
    RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\' + Exe, '', s) or
    RegQueryStringValue(HKLM32, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\' + Exe, '', s);
end;

function DetectOfficeText: String;
begin
  Result := '';
  if AppPathExists('WINWORD.EXE')  then Result := Result + 'Word、';
  if AppPathExists('EXCEL.EXE')    then Result := Result + 'Excel、';
  if AppPathExists('POWERPNT.EXE') then Result := Result + 'PowerPoint、';
  { 去掉末尾多出来的顿号 }
  if Length(Result) > 0 then
    Delete(Result, Length(Result), 1);
  if Result = '' then
    Result := '未检测到 Office（可以先装本插件，装好 Office 后会自动生效）';
end;

{ 把 'ImageZoom' 转成 UTF-16LE 的字节序列，用来在 Office 的二进制"禁用项"里找 }
function WideNeedle(const s: String): AnsiString;
var
  i: Integer;
  r: AnsiString;
begin
  r := '';
  for i := 1 to Length(s) do
    r := r + Copy(s, i, 1) + #0;
  Result := r;
end;

{ Office 把出过问题的加载项记成 Resiliency\DisabledItems 之类的二进制值，
  只要里面出现本插件的名字就删掉整条记录，否则 Office 会一直跳过它 }
procedure ClearDisabledItems;
var
  apps: array[0..2] of String;
  vers: array[0..3] of String;
  subs: array[0..2] of String;
  names: TArrayOfString;
  i, j, k, n: Integer;
  key: String;
  blob: AnsiString;
  needle: AnsiString;
begin
  apps[0] := 'Word'; apps[1] := 'Excel'; apps[2] := 'PowerPoint';
  vers[0] := '16.0'; vers[1] := '15.0'; vers[2] := '14.0'; vers[3] := '12.0';
  subs[0] := 'DisabledItems'; subs[1] := 'StartupItems'; subs[2] := 'DisabledAddins';
  needle := WideNeedle('ImageZoom');
  for i := 0 to 2 do
    for j := 0 to 3 do
      for k := 0 to 2 do
      begin
        key := 'Software\Microsoft\Office\' + vers[j] + '\' + apps[i] + '\Resiliency\' + subs[k];
        if RegGetValueNames(HKCU, key, names) then
        begin
          for n := 0 to GetArrayLength(names) - 1 do
          begin
            if RegQueryBinaryValue(HKCU, key, names[n], blob) then
            begin
              if Pos(needle, blob) > 0 then
                RegDeleteValue(HKCU, key, names[n]);
            end;
          end;
        end;
      end;
end;

function InitializeSetup: Boolean;
var
  msg: String;
begin
  Result := True;
  { 加载项本身是 .NET 4 程序集，没有 .NET 4 装了也不会工作，提前讲清楚 }
  if not HasDotNet4 then
  begin
    { 注意：ISPP 会把"行首"的 #13#10 当成预处理指令，所以拼接时每行都要以
      引号开头，转义序列只能放在上一行末尾 }
    msg := '本插件需要 .NET Framework 4.0 或更高版本，当前系统没有检测到。' + #13#10 + #13#10 +
           'Windows 7 用户请先安装 .NET Framework 4.6.2（微软官网免费下载），' + #13#10 +
           '装好之后再来运行本安装程序。Windows 8 及以上系统通常已经自带。' + #13#10 + #13#10 +
           '现在仍要继续安装吗？（装好 .NET 后插件才会生效）';
    if MsgBox(msg, mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

{ 把"检测到的 Office 组件"放进"准备安装"页的汇总清单。
  不放欢迎页：WelcomeLabel2 是固定大小的静态文本，中文再追加两行会被裁掉。 }
function Pick(const Cond: Boolean; const IfTrue, IfFalse: String): String;
begin
  if Cond then
    Result := IfTrue
  else
    Result := IfFalse;
end;

{ Inno 6 的原型必须一字不差（8 个 const 参数），否则报 Invalid prototype }
function UpdateReadyMemo(const Space, NewLine, MemoUserInfoInfo, MemoDirInfo,
  MemoTypeInfo, MemoComponentsInfo, MemoGroupInfo, MemoTasksInfo: String): String;
var
  extra: String;
begin
  extra := 'Office 组件：' + DetectOfficeText + NewLine +
           '运行环境：' + Pick(HasDotNet4, '.NET Framework 4.x 已就绪', '未检测到 .NET Framework 4.x') + NewLine +
           '安装完成后请重新打开 Office，插件才会加载。';
  Result := MemoDirInfo;
  if MemoTasksInfo <> '' then
    Result := Result + NewLine + NewLine + MemoTasksInfo;
  Result := Result + NewLine + NewLine + extra;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    ClearDisabledItems;
    { 让下一次启动 Office 时不被"修复"提示打扰：清掉安装期间可能残留的禁用记录 }
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    ClearDisabledItems;
    { 兼容早期用 install.ps1 / install.bat 手装过的机器：
      那些版本把加载项注册在 HKCU（虽然对托管 COM 无效，但会留在注册表里），
      顺手清干净，避免卸载后还残留一个指向不存在文件的条目 }
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\CLSID\' + cn_ProgId);
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Classes\ImageZoomAddin.Connect');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\Word\Addins\ImageZoomAddin.Connect');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\Excel\Addins\ImageZoomAddin.Connect');
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\PowerPoint\Addins\ImageZoomAddin.Connect');
  end;
end;
