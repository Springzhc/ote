; 出口销售合同管理系统安装脚本
; 生成命令: iscc setup.iss

[Setup]
AppName=出口销售合同管理系统
AppVersion=1.0
AppPublisher=您的公司名称
AppPublisherURL=https://github.com/Springzhc/ote
AppSupportURL=https://github.com/Springzhc/ote/issues
AppUpdatesURL=https://github.com/Springzhc/ote
DefaultDirName={autopf}\出口销售合同管理系统
DefaultGroupName=出口销售合同管理系统
OutputDir=d:\Mycode\ote\installer
OutputBaseFilename=ContractManagementSystem_Setup
Compression=lzma2
SolidCompression=yes

[Languages]
Name: "chinese"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]
Source: "d:\Mycode\ote\dist\ContractManagementSystem.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "d:\Mycode\ote\data\*"; DestDir: "{app}\data"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\出口销售合同管理系统"; Filename: "{app}\ContractManagementSystem.exe"
Name: "{commondesktop}\出口销售合同管理系统"; Filename: "{app}\ContractManagementSystem.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\出口销售合同管理系统"; Filename: "{app}\ContractManagementSystem.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\ContractManagementSystem.exe"; Description: "{cm:LaunchProgram,出口销售合同管理系统}"; Flags: nowait postinstall skipifsilent