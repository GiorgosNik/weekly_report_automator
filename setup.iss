[Setup]
AppName=Weekly Report Automator     
OutputDir=./installer
AppVersion=1.0.0
OutputBaseFilename=Weekly Report Automator
DefaultDirName={autopf}\Weekly Report Automator    
Compression=lzma
SolidCompression=yes
SetupIconFile=./assets/calendar_office_day.ico  

[Files]
Source: "./dist/*"; DestDir: "{app}"; Flags: ignoreversion
Source: "./assets/calendar_office_day.ico"; DestDir: "{app}/assets"; Flags: ignoreversion
Source: "./template/finalPresentationTemplate.pptx"; DestDir: "{app}/template"; Flags: ignoreversion
Source: "./template/firstBureauTemplate.pptx"; DestDir: "{app}/template"; Flags: ignoreversion
Source: "./template/otherBureausTemplate.pptx"; DestDir: "{app}/template"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional tasks"; Flags: unchecked

[Icons]
Name: "{group}\Weekly Report Automator"; Filename: "{app}\Weekly_Report_Automator.exe"
Name: "{autodesktop}\Weekly Report Automator"; Filename: "{app}\Weekly_Report_Automator.exe"; IconFilename: "{app}/assets/calendar_office_day.ico"; Tasks: desktopicon