[Setup]
AppName=LynxAutomator
AppVersion=0.001
DefaultDirName={pf}\LynxAutomator
DefaultGroupName=LynxAutomator
OutputBaseFilename=LynxAutomator-Setup
Compression=lzma
SolidCompression=yes

[Files]
Source: "C:\Users\WWF-POR113\Desktop\PythonApps\BIWBapp\dist\LynxAutomator_v001alpha.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\LynxAutomator"; Filename: "{app}\LynxAutomator_v001alpha.exe"
Name: "{commondesktop}\LynxAutomator"; Filename: "{app}\LynxAutomator_v001alpha.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked