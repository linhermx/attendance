#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#ifndef LauncherSourceDir
  #define LauncherSourceDir AddBackslash(SourcePath) + "..\dist\attendance_launcher"
#endif

#ifndef OutputDir
  #define OutputDir AddBackslash(SourcePath) + "..\dist\installer"
#endif

#ifndef SetupIconFilePath
  #define SetupIconFilePath AddBackslash(SourcePath) + "..\src\attendance\resources\attendance.ico"
#endif

#define MyAppId "LINHER.Attendance"
#define MyAppName "LINHER Attendance"
#define MyAppPublisher "LINHER"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\LINHER\Attendance
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir={#OutputDir}
OutputBaseFilename=attendance_setup
SetupIconFile={#SetupIconFilePath}
UninstallDisplayIcon={app}\attendance_launcher.exe

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; Flags: unchecked

[Files]
Source: "{#LauncherSourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\LINHER Attendance"; Filename: "{app}\attendance_launcher.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\LINHER Attendance"; Filename: "{app}\attendance_launcher.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\attendance_launcher.exe"; Description: "Abrir LINHER Attendance"; Flags: nowait postinstall skipifsilent
