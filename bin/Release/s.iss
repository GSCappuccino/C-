; 脚本由 Inno Setup 脚本向导 生成！
; 有关创建 Inno Setup 脚本文件的详细资料请查阅帮助文档！

#define MyAppName "商标书式管理系统"
#define MyAppVersion "1.0"
#define MyAppPublisher "G-S"
#define MyAppExeName "shangbiao.exe"

[Setup]
; 注: AppId的值为单独标识该应用程序。
; 不要为其他安装程序使用相同的AppId值。
; (生成新的GUID，点击 工具|在IDE中生成GUID。)
AppId={{7E6BAFC0-440E-4C2A-B915-911996D18DD6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=C:\Users\G-S\Desktop\exe,输出目录
OutputBaseFilename=setup_64
Compression=lzma
SolidCompression=yes

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]
Source: "G:\working\MyApp\bin\Release\shangbiao.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\Release\BarChart.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\Release\DevComponents.DotNetBar.SuperGrid.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\Release\DevComponents.DotNetBar2.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\Release\DevComponents.SuperGrid.Design.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\Release\Microsoft.Office.Interop.Word.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\G-S\Desktop\file\DB\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 注意: 不要在任何共享系统文件上使用“Flags: ignoreversion”

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

