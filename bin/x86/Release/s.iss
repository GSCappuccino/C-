; �ű��� Inno Setup �ű��� ���ɣ�
; �йش��� Inno Setup �ű��ļ�����ϸ��������İ����ĵ���

#define MyAppName "�̱���ʽ����ϵͳ"
#define MyAppVersion "1.0"
#define MyAppPublisher "G-S"
#define MyAppExeName "shangbiao.exe"

[Setup]
; ע: AppId��ֵΪ������ʶ��Ӧ�ó���
; ��ҪΪ������װ����ʹ����ͬ��AppIdֵ��
; (�����µ�GUID����� ����|��IDE������GUID��)
AppId={{813621E9-3CDD-4CB1-91B5-FDB666688B30}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=C:\Users\G-S\Desktop\exe,���Ŀ¼
OutputBaseFilename=setup_86
Compression=lzma
SolidCompression=yes

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]
Source: "G:\working\MyApp\bin\x86\Release\shangbiao.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\G-S\Desktop\file\DB\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "G:\working\MyApp\bin\x86\Release\BarChart.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\x86\Release\DevComponents.DotNetBar.SuperGrid.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\x86\Release\DevComponents.DotNetBar2.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\x86\Release\DevComponents.SuperGrid.Design.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "G:\working\MyApp\bin\x86\Release\Microsoft.Office.Interop.Word.dll"; DestDir: "{app}"; Flags: ignoreversion
; ע��: ��Ҫ���κι���ϵͳ�ļ���ʹ�á�Flags: ignoreversion��

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

