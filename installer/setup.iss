; =====================================================================
; Word 批量生成器 v2.0 - Inno Setup 安装脚本
; 编译方法：双击 build_installer.bat，或直接运行：
;   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\setup.iss
; 输出：installer\output\Word批量生成器_v2.0_Setup.exe
; =====================================================================

#define AppName      "Word 批量生成器"
#define AppNameShort "Word批量生成器"
#define AppVersion   "2.0.0"
#define AppPublisher "LZS"
#define AppExeName   "Word批量生成器.exe"
#define SourceExe    "..\dist\Word批量生成器.exe"
#define AppIcon      "..\WordBatchGenerator\Resources\app.ico"

[Setup]
; 基本信息
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=https://github.com/lzs
VersionInfoVersion={#AppVersion}
VersionInfoDescription={#AppName} 安装程序
VersionInfoProductName={#AppName}

; 安装目标
DefaultDirName={autopf}\{#AppNameShort}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes

; 安装包图标与外观
SetupIconFile={#AppIcon}
WizardStyle=modern
WizardSizePercent=120

; 输出
OutputDir=output
OutputBaseFilename={#AppNameShort}_v{#AppVersion}_Setup
Compression=lzma2/ultra64
SolidCompression=yes
InternalCompressLevel=ultra64

; 权限（不强制管理员，普通用户也能安装到 LocalAppData）
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; 卸载
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} v{#AppVersion}

; 最低系统要求
MinVersion=10.0

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon";    Description: "创建桌面快捷方式(&D)"; GroupDescription: "附加任务:"
Name: "startmenuicon"; Description: "创建开始菜单快捷方式(&S)"; GroupDescription: "附加任务:"

[Files]
; 主程序（单文件 exe，自包含 .NET 8 运行时，无其他依赖）
Source: "{#SourceExe}"; DestDir: "{app}"; DestName: "{#AppExeName}"; Flags: ignoreversion

[Icons]
; 开始菜单
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"; Tasks: startmenuicon
; 桌面快捷方式（图标显式指定，彻底解决大文件桌面图标问题）
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"; Tasks: desktopicon
; 卸载快捷方式（开始菜单）
Name: "{group}\卸载 {#AppName}"; Filename: "{uninstallexe}"

[Run]
; 安装完成后可选择立即启动
Filename: "{app}\{#AppExeName}"; Description: "立即启动 {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 卸载时清理程序自身产生的缓存（AppData 数据由用户自行保留，不强制删除）
Type: filesandordirs; Name: "{localappdata}\{#AppNameShort}\WebView2_Cache"
